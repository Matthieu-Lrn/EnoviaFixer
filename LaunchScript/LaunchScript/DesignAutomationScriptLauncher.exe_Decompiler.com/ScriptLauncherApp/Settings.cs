using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ScriptLauncherApp;

[Serializable]
[XmlRoot("Settings")]
public class Settings
{
	[Serializable]
	public class Setting<T>
	{
		[XmlAttribute]
		public T Value { get; set; }

		[XmlAttribute]
		public string Description { get; set; }

		public Setting(T sValue, string sDescription)
		{
			Description = "";
			Value = sValue;
			Description = sDescription;
		}

		public Setting()
		{
			Description = "";
		}
	}

	[Serializable]
	public class ScriptArgument
	{
		[XmlAttribute]
		public string Name { get; set; }

		[XmlAttribute]
		public string Value { get; set; }

		[XmlArray("ArgumentPaths")]
		[XmlArrayItem("ArgumentPath")]
		public List<string> ArgumentPathList { get; set; }

		public ScriptArgument()
		{
			Name = "";
			Value = "";
			ArgumentPathList = null;
		}
	}

	[XmlElement]
	public Setting<bool> IsActive { get; set; }

	[XmlArray("ScriptPaths")]
	[XmlArrayItem("ScriptPath")]
	public List<string> ScriptPathList { get; set; }

	[XmlElement]
	public Setting<string> NonActiveMessage { get; set; }

	[XmlArray("Arguments")]
	[XmlArrayItem("Argument")]
	public List<ScriptArgument> ArgumentList { get; set; }

	public Settings()
	{
		IsActive = null;
		ScriptPathList = null;
		NonActiveMessage = null;
		ArgumentList = null;
	}

	public static Settings Initialize(string sFilePath)
	{
		Settings result = null;
		StreamReader streamReader = null;
		if (File.Exists(sFilePath))
		{
			try
			{
				Module1.oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.EXECUTING, "Setting file exist", sFilePath, bSaveNow: false, bExit: false, null, bDisplayMsg: false, bReferToTracking: true, (MsgBoxStyle)16);
				streamReader = new StreamReader(sFilePath);
				result = (Settings)new XmlSerializer(typeof(Settings)).Deserialize(streamReader);
				streamReader.Close();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				streamReader?.Close();
				result = null;
				Module1.oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.CRITICAL, "Setting file could not be open or deserialized.", "", bSaveNow: true, bExit: true, ex2, bDisplayMsg: true, bReferToTracking: true, (MsgBoxStyle)16);
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			Module1.oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.CRITICAL, "Setting file can't be found.", sFilePath, bSaveNow: true, bExit: true, null, bDisplayMsg: true, bReferToTracking: true, (MsgBoxStyle)16);
		}
		return result;
	}
}
