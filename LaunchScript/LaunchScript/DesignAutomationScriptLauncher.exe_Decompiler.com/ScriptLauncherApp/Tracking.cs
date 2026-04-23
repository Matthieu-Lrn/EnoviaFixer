using System;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ScriptLauncherApp;

public class Tracking
{
	public enum InfoType
	{
		EXECUTING,
		WARNING_SILENT,
		WARNING_REPORT,
		CRITICAL
	}

	private StringBuilder sStrBuilder;

	private int mNbLine;

	private const int iBufferSize = 1;

	private string mFilePath { get; set; }

	public string FilePath => mFilePath;

	public Tracking(string FilePath, bool Overwrite = true)
	{
		sStrBuilder = new StringBuilder();
		mNbLine = 0;
		mFilePath = "";
		mFilePath = FilePath;
		if (File.Exists(mFilePath) && Overwrite)
		{
			File.Delete(mFilePath);
		}
	}

	public void AddLine(string sText)
	{
		AppendLine(sText);
	}

	public void AddEmptyLine()
	{
		AppendLine("");
	}

	public void AddInformation(string sText1, string sText2, bool bSaveNow = false)
	{
		int[] iColumnWidth = new int[2] { 40, 40 };
		string[] sString = new string[2] { sText1, sText2 };
		ConcatenateMultiColumnLine(sString, iColumnWidth, bSaveNow);
	}

	public void AddTrackingStep(MethodBase oMethodBase, InfoType sInfoType, string sMsg1, string sMsg2 = "", bool bSaveNow = false, bool bExit = false, Exception ex = null, bool bDisplayMsg = false, bool bReferToTracking = true, MsgBoxStyle oMsgBoxStyle = 16)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		AddTrackingStep(oMethodBase.ReflectedType.Name + "." + oMethodBase.Name, sInfoType, sMsg1, sMsg2, bSaveNow, bExit, ex, bDisplayMsg, bReferToTracking, oMsgBoxStyle);
	}

	private void AddTrackingStep(string sMethodName, InfoType sInfoType, string sMsg1, string sMsg2 = "", bool bSaveNow = false, bool bExit = false, Exception ex = null, bool bDisplayMsg = false, bool bReferToTracking = true, MsgBoxStyle oMsgBoxStyle = 16)
	{
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		int[] iColumnWidth = new int[5] { 27, 20, 40, 40, 30 };
		string text = Strings.Format((object)DateAndTime.Now, "yyyy/MM/dd HH:mm:ss:fff");
		string[] sString = new string[5]
		{
			text,
			sInfoType.ToString(),
			sMethodName,
			sMsg1,
			sMsg2
		};
		ConcatenateMultiColumnLine(sString, iColumnWidth, bSaveNow);
		if (ex != null)
		{
			AddInformation(ex.Message, "", bSaveNow: true);
		}
		if (bDisplayMsg)
		{
			StringBuilder stringBuilder = new StringBuilder();
			stringBuilder.AppendLine(sMsg1);
			if (Operators.CompareString(sMsg2, string.Empty, false) != 0)
			{
				stringBuilder.AppendLine(sMsg2);
			}
			if (bReferToTracking)
			{
				stringBuilder.AppendLine("\r\nRefer to " + mFilePath);
			}
			Interaction.MsgBox((object)stringBuilder.ToString(), (MsgBoxStyle)checked(0 + oMsgBoxStyle), (object)"Data Transfer");
		}
		if (bExit)
		{
			Environment.Exit(1);
		}
	}

	private void ConcatenateMultiColumnLine(string[] sString, int[] iColumnWidth, bool bSaveNow)
	{
		int num = 0;
		string text = "";
		int num2 = Information.LBound((Array)sString, 1);
		int num3 = Information.UBound((Array)sString, 1);
		checked
		{
			for (int i = num2; i <= num3; i++)
			{
				num += iColumnWidth[i];
				text += sString[i];
				text = text.PadRight(num);
			}
			AppendLine(text);
			Save(bSaveNow);
		}
	}

	private void AppendLine(string sLine)
	{
		sStrBuilder.AppendLine(sLine.TrimEnd(new char[0]));
		checked
		{
			mNbLine++;
		}
	}

	public void AddHeader(string sName, bool bSaveNow = false)
	{
		string text = "";
		AppendLine(text.PadRight(150, '#'));
		AppendLine(text);
		AppendLine(text.PadLeft(75) + sName);
		AppendLine(text);
		AppendLine(text.PadRight(150, '#'));
		Save(bSaveNow);
	}

	public void CloseHeader(bool bSaveNow = false)
	{
		string text = "";
		AppendLine(text.PadRight(150, '#'));
		Save(bSaveNow);
	}

	public void Save(bool bSaveNow)
	{
		if (mNbLine >= 1 || bSaveNow)
		{
			StreamWriter streamWriter = new StreamWriter(mFilePath, append: true);
			streamWriter.WriteLine(sStrBuilder.ToString().TrimEnd(new char[0]));
			streamWriter.Close();
			sStrBuilder.Clear();
			mNbLine = 0;
		}
	}
}
