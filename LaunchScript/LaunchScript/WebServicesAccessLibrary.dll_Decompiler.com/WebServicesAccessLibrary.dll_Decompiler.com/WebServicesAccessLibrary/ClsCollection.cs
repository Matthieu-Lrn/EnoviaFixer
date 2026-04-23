using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace WebServicesAccessLibrary;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class ClsCollection : Hashtable
{
	private string sKeyPrefix;

	public string Prefix
	{
		get
		{
			return sKeyPrefix;
		}
		set
		{
			sKeyPrefix = value;
		}
	}

	public object GetKeys
	{
		get
		{
			//IL_0000: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Expected O, but got Unknown
			Collection val = new Collection();
			foreach (object key in base.Keys)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(key);
				val.Add(RuntimeHelpers.GetObjectValue(objectValue), (string)null, (object)null, (object)null);
			}
			return val;
		}
	}

	public object GetItems
	{
		get
		{
			//IL_0000: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Expected O, but got Unknown
			Collection val = new Collection();
			foreach (object value in base.Values)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(value);
				val.Add(RuntimeHelpers.GetObjectValue(objectValue), (string)null, (object)null, (object)null);
			}
			return val;
		}
	}

	public string GetKey
	{
		get
		{
			if (iIndex <= Count)
			{
				object getKeys = GetKeys;
				object[] obj = new object[1] { iIndex };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				object obj3 = NewLateBinding.LateGet(getKeys, (Type)null, "item", obj, (string[])null, (Type[])null, obj2);
				if (array2[0])
				{
					iIndex = (short)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(short));
				}
				return Conversions.ToString(obj3);
			}
			return "";
		}
	}

	public object GetItem
	{
		get
		{
			if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(iIndex)))
			{
				object getItems = GetItems;
				object[] obj = new object[1] { iIndex };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				object obj3 = NewLateBinding.LateGet(getItems, (Type)null, "item", obj, (string[])null, (Type[])null, obj2);
				if (array2[0])
				{
					iIndex = RuntimeHelpers.GetObjectValue(array[0]);
				}
				return RuntimeHelpers.GetObjectValue(obj3);
			}
			return RuntimeHelpers.GetObjectValue(this[RuntimeHelpers.GetObjectValue(iIndex)]);
		}
	}

	public object GetItemByIndex
	{
		get
		{
			object getItems = GetItems;
			object[] obj = new object[1] { iIndex };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(getItems, (Type)null, "item", obj, (string[])null, (Type[])null, obj2);
			if (array2[0])
			{
				iIndex = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
			}
			return RuntimeHelpers.GetObjectValue(obj3);
		}
	}

	public object GetItemByKey => RuntimeHelpers.GetObjectValue(this[RuntimeHelpers.GetObjectValue(iIndex)]);

	public bool Exists => ContainsKey(sIndex);

	public ClsCollection()
	{
		base.Clear();
		sKeyPrefix = "";
	}

	public void SetItem(object iIndex, object Value)
	{
		if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(iIndex)))
		{
			this[this.get_GetKey(Conversions.ToShort(iIndex))] = RuntimeHelpers.GetObjectValue(Value);
		}
		else
		{
			this[Operators.ConcatenateObject((object)sKeyPrefix, iIndex)] = RuntimeHelpers.GetObjectValue(Value);
		}
	}

	public void SetItemByIndex(int iIndex, object Value)
	{
		this[RuntimeHelpers.GetObjectValue(this.get_GetItemByIndex(iIndex))] = RuntimeHelpers.GetObjectValue(Value);
	}

	public void SetItemByKey(object iIndex, object Value)
	{
		this[Operators.ConcatenateObject((object)sKeyPrefix, iIndex)] = RuntimeHelpers.GetObjectValue(Value);
	}

	public override void Add(object key, object value)
	{
		base.Add(Operators.ConcatenateObject((object)sKeyPrefix, key), RuntimeHelpers.GetObjectValue(value));
	}

	public override void Remove(object iIndex)
	{
		if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(iIndex)))
		{
			base.Remove(this.get_GetKey(Conversions.ToShort(iIndex)));
		}
		else
		{
			base.Remove(RuntimeHelpers.GetObjectValue(iIndex));
		}
	}

	public void RemoveByKey(object Key)
	{
		base.Remove(RuntimeHelpers.GetObjectValue(Key));
	}

	public void RemoveByIndex(int iIndex)
	{
		base.Remove(RuntimeHelpers.GetObjectValue(this.get_GetItemByIndex(iIndex)));
	}

	public void RemoveAll()
	{
		Clear();
	}
}
