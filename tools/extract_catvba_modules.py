import math
import re
import struct
from pathlib import Path


FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


class CFB:
    def __init__(self, path):
        self.path = Path(path)
        self.data = self.path.read_bytes()
        self.sector_size = 1 << struct.unpack_from("<H", self.data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self.data, 32)[0]
        self.num_fat = struct.unpack_from("<I", self.data, 44)[0]
        self.dir_start = struct.unpack_from("<I", self.data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", self.data, 56)[0]
        self.mini_fat_start = struct.unpack_from("<I", self.data, 60)[0]
        self.difat_start = struct.unpack_from("<I", self.data, 68)[0]
        self.num_difat = struct.unpack_from("<I", self.data, 72)[0]
        self.difat = list(struct.unpack_from("<109I", self.data, 76))
        self._load_difat()
        self._load_fat()
        self._load_dirs()
        self._load_minifat()

    def sector(self, sector_id):
        offset = 512 + sector_id * self.sector_size
        return self.data[offset : offset + self.sector_size]

    def chain(self, start, table=None):
        table = self.fat if table is None else table
        sector_id = start
        seen = set()
        result = []
        while sector_id not in (FREESECT, ENDOFCHAIN, NOSTREAM) and sector_id not in seen:
            seen.add(sector_id)
            result.append(sector_id)
            if sector_id >= len(table):
                break
            sector_id = table[sector_id]
        return result

    def _load_difat(self):
        sector_id = self.difat_start
        for _ in range(self.num_difat):
            if sector_id in (FREESECT, ENDOFCHAIN, NOSTREAM):
                break
            sector = self.sector(sector_id)
            self.difat.extend(struct.unpack_from("<127I", sector, 0))
            sector_id = struct.unpack_from("<I", sector, 508)[0]
        self.difat = [
            value
            for value in self.difat
            if value not in (FREESECT, ENDOFCHAIN, FATSECT, DIFSECT, NOSTREAM)
        ]

    def _load_fat(self):
        fat_bytes = bytearray()
        for sector_id in self.difat[: self.num_fat]:
            fat_bytes.extend(self.sector(sector_id))
        self.fat = list(struct.unpack("<" + "I" * (len(fat_bytes) // 4), fat_bytes))

    def read_regular_stream(self, start, size=None):
        data = bytearray()
        for sector_id in self.chain(start):
            data.extend(self.sector(sector_id))
        return bytes(data if size is None else data[:size])

    def _load_dirs(self):
        blob = self.read_regular_stream(self.dir_start)
        self.dirs = []
        self.paths = {}
        for offset in range(0, len(blob), 128):
            entry = blob[offset : offset + 128]
            if len(entry) < 128:
                continue
            name_len = struct.unpack_from("<H", entry, 64)[0]
            name = ""
            if name_len >= 2:
                name = entry[: name_len - 2].decode("utf-16le", errors="ignore")
            left, right, child = struct.unpack_from("<III", entry, 68)
            self.dirs.append(
                {
                    "name": name,
                    "type": entry[66],
                    "left": left,
                    "right": right,
                    "child": child,
                    "start": struct.unpack_from("<I", entry, 116)[0],
                    "size": struct.unpack_from("<Q", entry, 120)[0],
                }
            )

        def walk(index, prefix):
            if index == NOSTREAM or index >= len(self.dirs):
                return
            entry = self.dirs[index]
            walk(entry["left"], prefix)
            path = "/".join(prefix + [entry["name"]])
            self.paths[path] = entry
            if entry["type"] in (1, 5):
                walk(entry["child"], prefix + [entry["name"]])
            walk(entry["right"], prefix)

        if self.dirs:
            walk(self.dirs[0]["child"], [])
            self.root = self.dirs[0]
        else:
            self.root = None

    def _load_minifat(self):
        minifat_bytes = bytearray()
        for sector_id in self.chain(self.mini_fat_start):
            minifat_bytes.extend(self.sector(sector_id))
        self.minifat = (
            list(struct.unpack("<" + "I" * (len(minifat_bytes) // 4), minifat_bytes))
            if minifat_bytes
            else []
        )
        self.ministream = (
            self.read_regular_stream(self.root["start"], self.root["size"]) if self.root else b""
        )

    def read_stream(self, path):
        entry = self.paths[path]
        size = entry["size"]
        if entry["type"] == 2 and size < self.mini_cutoff and self.minifat:
            data = bytearray()
            sector_id = entry["start"]
            seen = set()
            while sector_id not in (FREESECT, ENDOFCHAIN, NOSTREAM) and sector_id not in seen:
                seen.add(sector_id)
                offset = sector_id * self.mini_sector_size
                data.extend(self.ministream[offset : offset + self.mini_sector_size])
                if sector_id >= len(self.minifat):
                    break
                sector_id = self.minifat[sector_id]
            return bytes(data[:size])
        return self.read_regular_stream(entry["start"], size)


def decompress_vba(data):
    if not data or data[0] != 1:
        raise ValueError("Not a VBA compressed container")
    position = 1
    output = bytearray()
    while position + 2 <= len(data):
        header = struct.unpack_from("<H", data, position)[0]
        position += 2
        chunk_size = (header & 0x0FFF) + 3
        chunk_end = min(len(data), position + chunk_size - 2)
        compressed = header & 0x8000
        if not compressed:
            output.extend(data[position:chunk_end])
            position = chunk_end
            continue

        chunk = bytearray()
        while position < chunk_end:
            flags = data[position]
            position += 1
            for bit in range(8):
                if position >= chunk_end:
                    break
                if not flags & (1 << bit):
                    chunk.append(data[position])
                    position += 1
                else:
                    if position + 2 > chunk_end:
                        position = chunk_end
                        break
                    token = struct.unpack_from("<H", data, position)[0]
                    position += 2
                    current = len(chunk)
                    bit_count = max(4, int(math.ceil(math.log(max(current, 1), 2))))
                    length_mask = 0xFFFF >> bit_count
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    source = current - offset
                    if source < 0:
                        raise ValueError("Invalid copy token")
                    for _ in range(length):
                        chunk.append(chunk[source])
                        source += 1
        output.extend(chunk)
        position = chunk_end
    return bytes(output)


def best_decompress_stream(blob):
    best = None
    for offset, marker in enumerate(blob):
        if marker != 1:
            continue
        try:
            decompressed = decompress_vba(blob[offset:])
        except Exception:
            continue
        head = decompressed[:8000]
        score = 0
        for pattern in (
            b"Attribute VB_Name",
            b"Option Explicit",
            b"Public Sub",
            b"Private Sub",
            b"Function ",
            b"End Sub",
        ):
            if pattern in head:
                score += 20
        if score and (best is None or score > best[0]):
            best = (score, offset, decompressed)
    if not best:
        return None, None
    return best[1], best[2]


def safe_name(name):
    return re.sub(r"[^A-Za-z0-9_. -]+", "_", name).strip(" .") or "module"


def extract_file(catvba_path, out_root):
    cfb = CFB(catvba_path)
    out_dir = Path(out_root) / Path(catvba_path).stem
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = []

    for path in sorted(cfb.paths):
        name = path.split("/")[-1]
        if "/VBA/" not in path:
            continue
        if name.lower() in {"dir", "_vba_project", "project", "projectwm", "projectlk"}:
            continue
        if name.startswith("__SRP_"):
            continue
        entry = cfb.paths[path]
        if entry["type"] != 2 or entry["size"] < 100:
            continue

        offset, decompressed = best_decompress_stream(cfb.read_stream(path))
        if decompressed is None:
            continue
        text = decompressed.decode("latin-1", errors="replace")
        match = re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', text)
        module_name = match.group(1) if match else name
        ext = ".cls" if re.search(r"Attribute\s+VB_(PredeclaredId|Base)\s*=", text) else ".bas"
        target = out_dir / f"{safe_name(module_name)}{ext}"
        target.write_text(text, encoding="utf-8", errors="replace")
        manifest.append(f"{module_name}{ext}\tstream={path}\toffset={offset}\tchars={len(text)}")

    (out_dir / "_manifest.txt").write_text("\n".join(manifest), encoding="utf-8")
    print(f"{Path(catvba_path).name}: extracted {len(manifest)} modules to {out_dir}")


def main():
    files = [
        r"LaunchScript\LaunchScript\BA_KBE_GCC_DDP_OCT14_2020d.catvba",
        r"LaunchScript\LaunchScript\BA_KBE_GCC_SAVEFILEBASE_OCT14_2020.catvba",
        r"LaunchScript\LaunchScript\BA_KBE_GCC_PROD_OCT14_2020.catvba",
        r"LaunchScript\LaunchScript\BA_KBE_GCC_CHECK_OCT14_2020.catvba",
    ]
    for file_path in files:
        extract_file(file_path, "extracted_vba")


if __name__ == "__main__":
    main()
