using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace WebServicesAccessLibrary;

[ComVisible(false)]
internal class ExcelAddin : IExcelAddIn
{
	public void AutoOpen()
	{
		ComServer.DllRegisterServer();
	}

	public void AutoClose()
	{
		ComServer.DllUnregisterServer();
	}
}
