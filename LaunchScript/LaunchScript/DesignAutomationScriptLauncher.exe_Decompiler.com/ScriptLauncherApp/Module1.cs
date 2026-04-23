using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ScriptLauncherApp;

[StandardModule]
internal sealed class Module1
{
	public static Tracking oTracking = null;

	[STAThread]
	public static void Main()
	{
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		string path = "";
		try
		{
			path = Path.GetTempPath();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox((object)"Can't retrieve the user temp folder. Process terminated", (MsgBoxStyle)16, (object)"Data transfer");
			Environment.Exit(1);
			ProjectData.ClearProjectError();
		}
		oTracking = new Tracking(Path.Combine(path, "LaunchScript_Tracking.txt"));
		oTracking.AddHeader("PROCESS");
		Settings settings = Settings.Initialize(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "LaunchScript_Settings.xml"));
		if (!settings.IsActive.Value)
		{
			Interaction.MsgBox((object)settings.NonActiveMessage.Value, (MsgBoxStyle)64, (object)"Launch Script");
			return;
		}
		Process process = new Process();
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.AppendLine("The following scripts don't exist:");
		foreach (string scriptPath in settings.ScriptPathList)
		{
			if (File.Exists(scriptPath))
			{
				oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.EXECUTING, "Script name is", scriptPath, bSaveNow: false, bExit: false, null, bDisplayMsg: false, bReferToTracking: true, (MsgBoxStyle)16);
				process.StartInfo.FileName = scriptPath;
				break;
			}
			stringBuilder.AppendLine(" -" + scriptPath);
		}
		if (Operators.CompareString(process.StartInfo.FileName, "", false) == 0)
		{
			oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.CRITICAL, stringBuilder.ToString(), "", bSaveNow: true, bExit: true, null, bDisplayMsg: true, bReferToTracking: true, (MsgBoxStyle)16);
		}
		stringBuilder = new StringBuilder();
		foreach (Settings.ScriptArgument argument in settings.ArgumentList)
		{
			string text = "";
			StringBuilder stringBuilder2 = new StringBuilder();
			if (Operators.CompareString(argument.Value.Trim(), "", false) != 0)
			{
				text = argument.Value;
			}
			else
			{
				stringBuilder2.AppendLine("The following files don't exist:");
				foreach (string argumentPath in argument.ArgumentPathList)
				{
					if (File.Exists(argumentPath))
					{
						text = argumentPath;
						break;
					}
					stringBuilder2.AppendLine(" -" + argumentPath);
				}
			}
			if (Operators.CompareString(text, "", false) == 0)
			{
				oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.CRITICAL, stringBuilder2.ToString(), "", bSaveNow: true, bExit: true, null, bDisplayMsg: true, bReferToTracking: true, (MsgBoxStyle)16);
			}
			if (stringBuilder.Length != 0)
			{
				stringBuilder.Append(" ");
			}
			stringBuilder.Append(argument.Name);
			stringBuilder.Append(text);
		}
		process.StartInfo.Arguments = stringBuilder.ToString();
		oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.EXECUTING, "Script arguments", stringBuilder.ToString(), bSaveNow: false, bExit: false, null, bDisplayMsg: false, bReferToTracking: true, (MsgBoxStyle)16);
		try
		{
			process.Start();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.CRITICAL, "Problem launching the script.", "", bSaveNow: true, bExit: true, ex4, bDisplayMsg: true, bReferToTracking: true, (MsgBoxStyle)16);
			ProjectData.ClearProjectError();
		}
		finally
		{
			oTracking.AddTrackingStep(MethodBase.GetCurrentMethod(), Tracking.InfoType.EXECUTING, "Done", "", bSaveNow: true, bExit: true, null, bDisplayMsg: false, bReferToTracking: true, (MsgBoxStyle)16);
		}
	}
}
