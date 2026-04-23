using System;
using System.Collections;
using System.ComponentModel;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WebServicesAccessLibrary.My;

namespace WebServicesAccessLibrary;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class WebServiceAccessTool
{
	private class EnhancedWebClient : WebClient
	{
		private int _TimeOutInMilliseconds;

		public int TimeOutInSeconds
		{
			get
			{
				return checked((int)Math.Round((double)_TimeOutInMilliseconds / 1000.0));
			}
			set
			{
				_TimeOutInMilliseconds = checked(value * 1000);
			}
		}

		public EnhancedWebClient()
		{
			_TimeOutInMilliseconds = 100000;
		}

		protected override WebRequest GetWebRequest(Uri uri)
		{
			WebRequest webRequest = base.GetWebRequest(uri);
			webRequest.Timeout = _TimeOutInMilliseconds;
			return webRequest;
		}
	}

	private const string sAuthApps = "I:\\V5_KBE_Tools\\Production\\05_KBE_CATScript\\03_ENOVIA_Connection_Tools\\2-PROD\\AuthorizedApps.xml";

	private const string sHTTP = "https://";

	private const string SubDomainBDIPROD = "bdi";

	private const string SubDomainBDIDEVT = "bdi-dev";

	private const string sOnsiteDomainName = "ca.aero.bombardier.net";

	private const string sOffsiteDomainName = "space.aero.bombardier.net";

	private const string sPortNumber = "";

	private const string WebService_NIEO = "%SERVERNAME%/api/reports/nieo/%PARTNUMBER%";

	private const string WebService_ERDN = "%SERVERNAME%/api/reports/edrn/%BASENUMBER%/%REVISION%";

	private const string WebService_RV = "%SERVERNAME%/api/reports/rv/%RVNUMBER%/%REVISION%";

	private const string WebService_UserInfo = "%SERVERNAME%/api/reports/echecker/userinfo/%USERID%";

	private const string WebService_StringVsEBOM = "%SERVERNAME%/api/reports/assy/%PARTNUMBER%/%PROJECTNUMBER%";

	private const string WebService_PLMAction = "%SERVERNAME%/api/reports/enovia/plmaction/%BASENUMBER%/%REVISION%";

	private const string WebService_BADeliverable = "%SERVERNAME%/api/echecker/reports/enovia/BADeliverable/%BASENUMBER%/%REVISION%";

	private const string WebService_BADeliverableDocuments = "%SERVERNAME%/api/echecker/reports/enovia/attachedDocumentsByBADeliverable/%ACTIONNUMBER%";

	private const string WebService_PLMActionDocuments = "%SERVERNAME%/api/echecker/reports/enovia/attachedDocumentsByPLMAction/%PLMACTION%";

	private const string WebService_DocumentAttributs = "%SERVERNAME%/api/reports/enovia/document/%BASENUMBER%/%REVISION%";

	private const string WebService_SearchDocumentsByAttributs = "%SERVERNAME%/api/echecker/reports/enovia/getdocuments/";

	private const string WebService_AssemblyMatrices = "%SERVERNAME%/api/reports/enovia/assyRelations/%BASENUMBER%";

	private const string WebService_DocumentsByBaseNumber = "%SERVERNAME%/api/echecker/reports/enovia/documentsByBaseNumber/%BASENUMBER%";

	private const string WebService_NHA = "%SERVERNAME%/api/reports/enovia/NHA/%PARTNUMBER%";

	private const string Webservice_OIDNHA = "%SERVERNAME%/api/echecker/reports/enovia/nha/%PARTNUMBER%";

	private const string WebService_FTVCDLPartInfo = "%SERVERNAME%/api/echecker/ftvcdl/parts/%PARTNUMBER%";

	private const string WebService_EDRNBox21 = "%SERVERNAME%/api/echecker/reports/enovia/fillEDRNBox21/?partNumber=%PARTNUMBER%&rev=%REVISION%&history=%HISTORY%";

	private const string WebService_LogUsage = "%SERVERNAME%/api/echecker/logtoolusage/";

	private const string WebService_LogKPI = "%SERVERNAME%/api/echecker/KPI/";

	private const string WebService_GETKPITOP2000 = "%SERVERNAME%/api/echecker/KPI/TOP2000/";

	private const string WebService_PartRefByPartNumber = "%SERVERNAME%/api/echecker/reports/enovia/part/getPartRefByPartNumber/%PARTNUMBER%";

	private const string WebService_PartInstanceByInstanceName = "%SERVERNAME%/api/echecker/reports/enovia/part/getPartInstanceByInstanceNumber/%INSTANCENUMBER%";

	private const string WebService_PrcOidByPrcName = "%SERVERNAME%/api/echecker/reports/enovia/prc/getPrcByVid/%ID%";

	private const string WebService_EBOMStringContent = "%SERVERNAME%/api/reports/ebom/configstrings/%FAMILYID%/%MONUMENTID%/%VARIANTID%/%ENVELOPID%";

	private const string Webservice_EnoviaDocumentsLinkedToPart = "%SERVERNAME%/api/echecker/reports/enovia/getDocumentLinksToParts/%PartNumber%/%Rev%";

	private const string WebService_FMLRawmaterialAll = "%SERVERNAME%/api/echecker/fml/rawmaterial/allfml/";

	private const string Webservice_FMLListByDescription = "%SERVERNAME%/api/echecker/fml/rawmaterial/fmlListByDescription/%Description%";

	private const string WebService_FMLRawmaterial = "%SERVERNAME%/api/echecker/fml/rawmaterial/?";

	private const string Webservice_EnoviaLoginContext = "%SERVERNAME%/api/echecker/reports/enovia/enovialogincontexts/?enoviauserid=%LOGINID%&enoviadatabase=%ENOMDBNAME%";

	private const string JSONwebservice_BDIPARTInfo = "%SERVERNAME%/api/parts/";

	private const string JSONwebservice_BDIPARTInfoRev = "%SERVERNAME%/api/parts/?num=%BASENUMBER%&rev=%REVISION%";

	private const string JSONwebservice_BDIDOCInfo = "%SERVERNAME%/api/documents/";

	private const string JSONwebservice_BDIDOCInfoPartRev = "%SERVERNAME%/api/documents/?num=%BASENUMBER%&rev=%REVISION%";

	private const string JSONwebservice_BDIDOCRelInfo = "%SERVERNAME%/api/documents/%DOCID%/rvDocRels";

	private const string JSONwebservice_BDIDOCEffectivity = "%SERVERNAME%/api/documents/%DocumentNumber%/%Rev%/effectivity";

	private const string JSONwebservice_BDISearchInfo = "%SERVERNAME%/quicksearch/%BDIobjectType%/_search?q=keys:";

	private const string JSONwebservice_BDIUserInfo = "%SERVERNAME%/api/users/me";

	private const string JSONwebservice_BDIProjectInfo = "%SERVERNAME%/api/projects/";

	private const string JSONwebservice_RVWorkFlow = "%SERVERNAME%/api/rvs/%RVID%/workflow";

	private const string JSONwebservice_DOCWorkFlow = "%SERVERNAME%/api/documents/%DOCID%/workflow";

	private const string JSONwebservice_ListOfAttachmentsDPS = "%SERVERNAME%/bdicommon/attachments/210/%DocumentID%/";

	private const string JSONWebservice_DownLoadAttachment = "%SERVERNAME%/bdicommon/attachments/%AttachmentID%/file";

	private const string JSONWebservice_DocmentEffectivity = "%SERVERNAME%/api/documents/%PartNumber%/%Rev%/effectivity";

	private const string JSONWebservice_BDIPartChildren = "%SERVERNAME%/api/parts/%PartID%/childrens/";

	private const string JSONWebservice_BDIPartDPSDocumments = "%SERVERNAME%/api/parts/%PartID%/documents/";

	private const string JSONWebservice_BDIPartChildrenList = "%SERVERNAME%/ebom/strings/getPartChildrenListWithoutStringLevel?eBom_LibItemId=%PARTID%&inAcFilter_id=%ACID%&outAcFilter_id=%ACID%";

	private const string JSONWebservice_BDIPARTStrings = "%SERVERNAME%/api/parts/%PARTID%/strings";

	private string sHeaderID;

	public string webServiceErrorMessage;

	private Hashtable oCache;

	private int _ParentWindowHWND;

	private string _sWebserviceLogPayload;

	private const int iLogMaxLineCnt = 100;

	private int iLogLineCnt;

	private string _WebtransactionTime;

	private bool _ResultFromCache;

	private string _LocationID;

	private bool _UseDevtDBForGet;

	public const string ENOMDBDefaultName = "ENOM6PRD";

	private string _WebClientReturnStatus;

	private string _WebResponseExceptionStatus;

	private string _GeneralException;

	private int _FrmLeft;

	private int _FrmTop;

	[CompilerGenerated]
	[AccessedThroughProperty("BckGdWrk")]
	private BackgroundWorker _BckGdWrk;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sDomainName;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sUserCode;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber;

	[SpecialName]
	private string _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName;

	[SpecialName]
	private object _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName;

	[SpecialName]
	private Hashtable _0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd;

	public string LocationID
	{
		get
		{
			if (Operators.CompareString(_LocationID, "", false) == 0)
			{
				try
				{
					string text = Dns.GetHostEntry(Dns.GetHostName()).AddressList[0].ToString();
					StreamReader streamReader = new StreamReader(MySettingsProperty.Settings.sLocInfoFile);
					string text2 = "";
					string text3 = "";
					string text4;
					int num = default(int);
					do
					{
						text4 = streamReader.ReadLine();
						if (LikeOperator.LikeString("loc." + text, Strings.Trim(Strings.Split(text4, "=", -1, (CompareMethod)0)[0]) + ".*", (CompareMethod)0) && Strings.Trim(Strings.Split(text4, "=", -1, (CompareMethod)0)[0]).Length > text2.Length)
						{
							text2 = Strings.Trim(Strings.Split(text4, "=", -1, (CompareMethod)0)[0]);
							text3 = Strings.Trim(Strings.Split(text4, "=", -1, (CompareMethod)0)[1]);
						}
						num = checked(num + 1);
					}
					while (!(text4 == null || num > 500));
					_LocationID = text3.ToUpper();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (Operators.CompareString(_LocationID, "", false) == 0)
				{
					_LocationID = Strings.UCase(Environment.GetEnvironmentVariable("V5START_LOCID"));
				}
			}
			return _LocationID;
		}
	}

	public string WebserviceToolVersion
	{
		get
		{
			AssemblyName name = Assembly.GetExecutingAssembly().GetName();
			return name.Name + name.Version.ToString();
		}
	}

	public string UsedbyApplication { get; set; }

	public string V5StartMode { get; set; }

	public string V5StartACCode { get; set; }

	public string V5StartENODB { get; set; }

	public string UserFullName { get; set; }

	public int TimeOutInSeconds { get; set; }

	public int MaximumReadConnectionAttempts { get; set; }

	public int MaximumWriteConnectionAttempts { get; set; }

	public string DefaultDownloadDirectoryPath { get; set; }

	public bool EnableToolUsageLog { get; set; }

	public string ServerName => sServerName(UseDevtDBForGet);

	public bool UseDevtDB { get; set; }

	public bool UseDevtDBForGet
	{
		get
		{
			if (UseDevtDB)
			{
				return _UseDevtDBForGet;
			}
			return UseDevtDB;
		}
		set
		{
			_UseDevtDBForGet = value;
		}
	}

	private string _ENOMDBNAME { get; set; }

	public string ENOMDBNAME
	{
		get
		{
			if (Operators.CompareString(_ENOMDBNAME, "", false) == 0)
			{
				return "ENOM6PRD";
			}
			return _ENOMDBNAME;
		}
		set
		{
			_ENOMDBNAME = value;
		}
	}

	public int ParentWindowHWND
	{
		get
		{
			return _ParentWindowHWND;
		}
		set
		{
			_ParentWindowHWND = value;
		}
	}

	public string WebClientReturnStatus => _WebClientReturnStatus;

	public string WebResponseExceptionStatus => _WebResponseExceptionStatus;

	public string GeneralException => _GeneralException;

	public int FrmLeft
	{
		get
		{
			return _FrmLeft;
		}
		set
		{
			_FrmLeft = value;
		}
	}

	public int FrmTop
	{
		get
		{
			return _FrmTop;
		}
		set
		{
			_FrmTop = value;
		}
	}

	public int ProcessID { get; set; }

	private virtual BackgroundWorker BckGdWrk
	{
		[CompilerGenerated]
		get
		{
			return _BckGdWrk;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			DoWorkEventHandler value2 = BckGdWrk_DoWork;
			BackgroundWorker bckGdWrk = _BckGdWrk;
			if (bckGdWrk != null)
			{
				bckGdWrk.DoWork -= value2;
			}
			_BckGdWrk = value;
			bckGdWrk = _BckGdWrk;
			if (bckGdWrk != null)
			{
				bckGdWrk.DoWork += value2;
			}
		}
	}

	public WebServiceAccessTool()
	{
		sHeaderID = "WebServiceAccessTool_DEVT";
		webServiceErrorMessage = "Cannot connect to BDI !";
		_ParentWindowHWND = 0;
		_LocationID = "";
		UsedbyApplication = "";
		V5StartMode = "";
		V5StartACCode = "";
		V5StartENODB = "";
		UserFullName = "";
		TimeOutInSeconds = 100;
		MaximumReadConnectionAttempts = 4;
		MaximumWriteConnectionAttempts = 1;
		DefaultDownloadDirectoryPath = "";
		EnableToolUsageLog = true;
		UseDevtDB = false;
		_UseDevtDBForGet = false;
		_WebClientReturnStatus = "";
		_WebResponseExceptionStatus = "";
		_GeneralException = "";
		_FrmLeft = 0;
		_FrmTop = 0;
		ProcessID = -1;
		oCache = new Hashtable();
		sHeaderID = MySettingsProperty.Settings.sHeaderID;
	}

	public void ClearCache()
	{
		oCache.Clear();
	}

	public string GetMachineOS()
	{
		_GeneralException = "";
		string result = "";
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(Interaction.GetObject("winmgmts:", (string)null), (Type)null, "InstancesOf", new object[1] { "Win32_OperatingSystem" }, (string[])null, (Type[])null, (bool[])null));
			IEnumerator enumerator = ((IEnumerable)objectValue).GetEnumerator();
			try
			{
				if (enumerator.MoveNext())
				{
					result = Strings.Trim(Conversions.ToString(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(enumerator.Current), (Type)null, "Caption", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			object objectValue = null;
		}
		return result;
	}

	private string GetCurrentDNSDomainName()
	{
		return ((ActiveDirectoryPartition)Domain.GetComputerDomain()).Name.ToUpper();
	}

	public ClsCollection GetToolLogUsage(string sLogType, bool bFromModifDate, string dateRange_l = "", string dateRange_h = "", string sPartNumber = "", string sRevision = "", string sMacroNameAndVersion = "", string sFunctionName = "", string sV5StartLocation = "", bool GetUncachedResult = false)
	{
		ClsCollection oCol = new ClsCollection();
		string text = "";
		DateTime dateTime = ((!(!Information.IsDBNull((object)dateRange_h) & (Operators.CompareString(dateRange_h, "", false) != 0))) ? DateTime.Parse("2017-11-01") : DateTime.Parse(dateRange_l));
		DateTime dateTime2 = ((!(!Information.IsDBNull((object)dateRange_h) & (Operators.CompareString(dateRange_h, "", false) != 0))) ? DateAndTime.DateAdd("d", 2.0, (object)DateAndTime.Now) : DateTime.Parse(dateRange_h));
		if (Operators.CompareString(sLogType, "", false) != 0)
		{
			text = "log_type=" + sLogType;
		}
		text = text + "&modifTime=" + bFromModifDate.ToString().ToUpper();
		if (Operators.CompareString(sMacroNameAndVersion, "", false) != 0)
		{
			text = text + "&macro_name_and_version=" + sMacroNameAndVersion;
		}
		if (Operators.CompareString(sFunctionName, "", false) != 0)
		{
			text = text + "&function_name=" + sFunctionName;
		}
		if (Operators.CompareString(sV5StartLocation, "", false) != 0)
		{
			text = text + "&v5start_location=" + sV5StartLocation;
		}
		if (Operators.CompareString(sPartNumber, "", false) != 0)
		{
			text = text + "&part_number=" + sPartNumber;
		}
		if (Operators.CompareString(sRevision, "", false) != 0)
		{
			text = text + "&part_rev=" + sRevision;
		}
		if (Operators.CompareString(dateTime.ToString(), "", false) != 0)
		{
			text = text + "&dateRange_l=" + dateTime.ToString("yyyy-MM-dd") + "T00:00:00Z";
		}
		if (Operators.CompareString(dateTime2.ToString(), "", false) != 0)
		{
			text = text + "&dateRange_h=" + dateTime2.ToString("yyyy-MM-dd") + "T00:00:00Z";
		}
		string text2 = Strings.Replace("%SERVERNAME%/api/echecker/logtoolusage/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text2 = text2 + "?" + Strings.Replace(text, " ", "%20", 1, -1, (CompareMethod)0);
		string text3 = sGetWebResult(text2, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text3, "", false) == 0) | (Operators.CompareString(text3, webServiceErrorMessage, false) == 0))
		{
			result = oCol;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			try
			{
				xmlDocument.LoadXml(text3);
				oCol = new ClsCollection();
				_PopulateClsColfromXMLdata(ref oCol, xmlDocument.SelectSingleNode("/data").ChildNodes);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			result = oCol;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2);
			}
		}
		return result;
	}

	public void LogToolUsage(string sLog_Type, string sLog_Status, string sMacro_name_and_version, string sMacro_Mode, string sFunction_name, string sPartNumber, string sPartRev, object iOptionalPartIteration = null, string sOptionalValue = "", string sOptionalArgument1 = "", string sOptionalArgument2 = "", string sOptionalArgument3 = "", string sOptionalArgument4 = "", string sOptionalArgument5 = "", string sOptionalArgument6 = "", string sOptionalArgument7 = "", string sOptionalArgument8 = "", string sOptionalArgument9 = "", string sOptionalArgument10 = "")
	{
		_GeneralException = "";
		string sLogXML = "";
		string text = MonitorLogUsage(ref sLogXML, sLog_Type, sLog_Status, sMacro_name_and_version, sMacro_Mode, sFunction_name, sPartNumber, sPartRev, RuntimeHelpers.GetObjectValue(iOptionalPartIteration), sOptionalValue, sOptionalArgument1, sOptionalArgument2, sOptionalArgument3, sOptionalArgument4, sOptionalArgument5, sOptionalArgument6, sOptionalArgument7, sOptionalArgument8, sOptionalArgument9, sOptionalArgument10);
		if (Operators.CompareString(text, "", false) == 0)
		{
			sLogXML = FinalizeXMlbeforePost(sLogXML);
			string strURL = Strings.Replace("%SERVERNAME%/api/echecker/logtoolusage/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
			_SendXMLToServer(sLogXML, strURL);
		}
		else
		{
			_GeneralException = text;
		}
	}

	private string MonitorLogUsage(ref string sLogXML, string sLog_Type, string sLog_Status, string sMacro_name_and_version, string sMacro_Mode, string sFunction_name, string sPartNumber, string sPartRev, object iOptionalPartIteration = null, string sOptionalValue = "", string sOptionalArgument1 = "", string sOptionalArgument2 = "", string sOptionalArgument3 = "", string sOptionalArgument4 = "", string sOptionalArgument5 = "", string sOptionalArgument6 = "", string sOptionalArgument7 = "", string sOptionalArgument8 = "", string sOptionalArgument9 = "", string sOptionalArgument10 = "")
	{
		if ((Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sDomainName, "", false) == 0) | (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sUserCode, "", false) == 0))
		{
			object objectValue = RuntimeHelpers.GetObjectValue(Interaction.CreateObject("WScript.Network", ""));
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sDomainName = Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "UserDomain", new object[0], (string[])null, (Type[])null, (bool[])null));
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sUserCode = Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "UserName", new object[0], (string[])null, (Type[])null, (bool[])null));
			objectValue = null;
		}
		if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset, "", false) == 0)
		{
			_ = TimeZone.CurrentTimeZone;
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString("c");
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset = Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset, checked(Strings.InStrRev(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset, ":", -1, (CompareMethod)0) - 1));
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset = Conversions.ToString(Interaction.IIf(Versioned.IsNumeric((object)Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset, 1)), (object)("+" + _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset), (object)_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset));
		}
		if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName, "", false) == 0)
		{
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName = Interaction.Environ("computername");
		}
		if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem, "", false) == 0)
		{
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem = GetMachineOS();
		}
		string text = LocationID;
		if (Operators.CompareString(V5StartMode, "", false) == 0)
		{
			V5StartMode = Strings.UCase(Environment.GetEnvironmentVariable("CAT_LEVEL"));
		}
		if (Operators.CompareString(V5StartACCode, "", false) == 0)
		{
			V5StartACCode = Strings.UCase(Environment.GetEnvironmentVariable("CAT_CUST_VERS"));
		}
		if (Operators.CompareString(V5StartENODB, "", false) == 0)
		{
			V5StartENODB = Strings.UCase(Environment.GetEnvironmentVariable("ENO_DBUNAME"));
		}
		if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber, "", false) == 0)
		{
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber = Environment.GetEnvironmentVariable("V5START_USERID");
			if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber, "", false) == 0)
			{
				object objectValue2 = RuntimeHelpers.GetObjectValue(Interaction.CreateObject("WScript.Network", ""));
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber = Conversions.ToString(NewLateBinding.LateGet(objectValue2, (Type)null, "UserName", new object[0], (string[])null, (Type[])null, (bool[])null));
				objectValue2 = null;
			}
		}
		if (Operators.CompareString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName, "", false) == 0)
		{
			_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName = GetNetworkUserFullNameFromLogin(Environment.GetEnvironmentVariable("USERDOMAIN"), _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber);
		}
		if (Operators.CompareString(sMacro_Mode, "", false) == 0)
		{
			if (Operators.ConditionalCompareObjectEqual(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName, (object)"", false))
			{
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName = MySettingsProperty.Settings.sHeaderID;
				bool flag = true;
				if (Operators.ConditionalCompareObjectEqual((object)flag, LikeOperator.LikeObject(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName, (object)"*DEVT*", (CompareMethod)0), false))
				{
					_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName = "DEVT";
				}
				else if (Operators.ConditionalCompareObjectEqual((object)flag, LikeOperator.LikeObject(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName, (object)"*TEST*", (CompareMethod)0), false))
				{
					_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName = "TEST";
				}
				else if (Operators.ConditionalCompareObjectEqual((object)flag, LikeOperator.LikeObject(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName, (object)"*PROD*", (CompareMethod)0), false))
				{
					_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName = "PROD";
				}
				else
				{
					_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName = "N/A";
				}
			}
			sMacro_Mode = Conversions.ToString(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sLocalMacroName);
		}
		string text2 = "";
		if (Information.IsDBNull((object)sLog_Type))
		{
			text2 += ",sLog_Type";
		}
		if (Information.IsDBNull((object)sMacro_name_and_version))
		{
			text2 += ",sMacro_name_and_version";
		}
		if (Information.IsDBNull((object)sFunction_name))
		{
			text2 += ",sFunction_name";
		}
		if (Information.IsDBNull((object)sPartNumber))
		{
			text2 += ",sPartNumber";
		}
		if (Information.IsDBNull((object)sPartRev))
		{
			text2 += ",sPartRev";
		}
		if (iOptionalPartIteration != null && !Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(iOptionalPartIteration)))
		{
			text2 += ",iOptionalPartIteration";
		}
		if (Operators.CompareString(text2, "", false) != 0)
		{
			text2 = "Following parameter(s) cannot be null : " + text2;
		}
		else
		{
			if (Strings.Len(sLog_Type) > 25)
			{
				sLog_Type = Strings.Left(sLog_Type, 25);
			}
			if (Strings.Len(sLog_Status) > 50)
			{
				sLog_Status = Strings.Left(sLog_Status, 50);
			}
			if (Strings.Len(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName) > 25)
			{
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName = Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName, 25);
			}
			if (Strings.Len(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem) > 50)
			{
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem = Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem, 50);
			}
			if (Strings.Len(sMacro_name_and_version) > 50)
			{
				sMacro_name_and_version = Strings.Left(sMacro_name_and_version, 50);
			}
			if (Strings.Len(sMacro_Mode) > 5)
			{
				sMacro_Mode = Strings.Left(sMacro_Mode, 5);
			}
			if (Strings.Len(sFunction_name) > 50)
			{
				sFunction_name = Strings.Left(sFunction_name, 50);
			}
			if (Strings.Len(sPartNumber) > 50)
			{
				sPartNumber = Strings.Left(sPartNumber, 50);
			}
			if (Strings.Len(sPartRev) > 2)
			{
				sPartRev = Strings.Left(sPartRev, 2);
			}
			if (Strings.Len(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber) > 8)
			{
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber = Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber, 8);
			}
			if (Strings.Len(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName) > 50)
			{
				_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName = Strings.Left(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName, 50);
			}
			if (Strings.Len(text) > 5)
			{
				text = Strings.Left(text, 5);
			}
			if (Strings.Len(V5StartMode) > 5)
			{
				V5StartMode = Strings.Left(V5StartMode, 5);
			}
			if (Strings.Len(V5StartACCode) > 10)
			{
				V5StartACCode = Strings.Left(V5StartACCode, 10);
			}
			if (Strings.Len(V5StartENODB) > 10)
			{
				V5StartENODB = Strings.Left(V5StartENODB, 10);
			}
			sLogXML += " <part ";
			if (!string.IsNullOrEmpty(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset))
			{
				sLogXML = sLogXML + " zone_offset=\"" + _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sTimeZoneOffset + "\"";
			}
			if (!string.IsNullOrEmpty(sLog_Type))
			{
				sLogXML = sLogXML + " log_type=\"" + ConvertForbiddenCharacters(sLog_Type) + "\"";
			}
			if (!string.IsNullOrEmpty(sLog_Status))
			{
				sLogXML = sLogXML + " log_status=\"" + ConvertForbiddenCharacters(sLog_Status) + "\"";
			}
			if (!string.IsNullOrEmpty(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName))
			{
				sLogXML = sLogXML + " machine_name=\"" + _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sMachineName + "\"";
			}
			if (!string.IsNullOrEmpty(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem))
			{
				sLogXML = sLogXML + "  operating_system=\"" + _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024sOperatingSystem + "\"";
			}
			if (!string.IsNullOrEmpty(sMacro_name_and_version))
			{
				sLogXML = sLogXML + " macro_name_and_version=\"" + ConvertForbiddenCharacters(sMacro_name_and_version) + "\"";
			}
			if (!string.IsNullOrEmpty(sMacro_Mode))
			{
				sLogXML = sLogXML + " macro_mode=\"" + ConvertForbiddenCharacters(sMacro_Mode) + "\"";
			}
			if (!string.IsNullOrEmpty(sFunction_name))
			{
				sLogXML = sLogXML + " function_name=\"" + ConvertForbiddenCharacters(sFunction_name) + "\"";
			}
			if (!string.IsNullOrEmpty(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName))
			{
				sLogXML = sLogXML + " usr_fn=\"" + ConvertForbiddenCharacters(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserFullName) + "\"";
			}
			if (!string.IsNullOrEmpty(_0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber))
			{
				sLogXML = sLogXML + " usr_bn=\"" + _0024STATIC_0024MonitorLogUsage_00242014E10EEEEEEEE1CEEEEEEEEEEE_0024UserBadgeNumber + "\"";
			}
			if (!string.IsNullOrEmpty(sPartNumber))
			{
				sLogXML = sLogXML + " part_number=\"" + sPartNumber + "\"";
			}
			if (!string.IsNullOrEmpty(sPartRev))
			{
				sLogXML = sLogXML + " part_rev=\"" + sPartRev + "\"";
			}
			if (iOptionalPartIteration != null && !string.IsNullOrEmpty(Conversions.ToString(iOptionalPartIteration)))
			{
				sLogXML = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject((object)(sLogXML + " part_iter=\""), iOptionalPartIteration), (object)"\""));
			}
			if (Operators.CompareString(sOptionalValue, "", false) != 0 && !string.IsNullOrEmpty(sOptionalValue))
			{
				sLogXML = sLogXML + " value=\"" + ConvertForbiddenCharacters(sOptionalValue) + "\"";
			}
			sLogXML += " ><v5start ";
			if (Operators.CompareString(text, "", false) != 0 && !string.IsNullOrEmpty(text))
			{
				sLogXML = sLogXML + " v5start_location=\"" + text + "\"";
			}
			if (Operators.CompareString(V5StartMode, "", false) != 0 && !string.IsNullOrEmpty(V5StartMode))
			{
				sLogXML = sLogXML + " v5start_mode=\"" + V5StartMode + "\"";
			}
			if (Operators.CompareString(V5StartACCode, "", false) != 0 && !string.IsNullOrEmpty(V5StartACCode))
			{
				sLogXML = sLogXML + " v5start_ac_code=\"" + V5StartACCode + "\"";
			}
			if (Operators.CompareString(V5StartENODB, "", false) != 0 && !string.IsNullOrEmpty(V5StartENODB))
			{
				sLogXML = sLogXML + " v5start_enovia_db=\"" + V5StartENODB + "\"";
			}
			sLogXML += " /><arg ";
			if (Operators.CompareString(sOptionalArgument1, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument1))
			{
				sLogXML = sLogXML + " arg1=\"" + ConvertForbiddenCharacters(sOptionalArgument1) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument2, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument2))
			{
				sLogXML = sLogXML + " arg2=\"" + ConvertForbiddenCharacters(sOptionalArgument2) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument3, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument3))
			{
				sLogXML = sLogXML + " arg3=\"" + ConvertForbiddenCharacters(sOptionalArgument3) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument4, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument4))
			{
				sLogXML = sLogXML + " arg4=\"" + ConvertForbiddenCharacters(sOptionalArgument4) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument5, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument5))
			{
				sLogXML = sLogXML + " arg5=\"" + ConvertForbiddenCharacters(sOptionalArgument5) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument6, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument6))
			{
				sLogXML = sLogXML + " arg6=\"" + ConvertForbiddenCharacters(sOptionalArgument6) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument7, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument7))
			{
				sLogXML = sLogXML + " arg7=\"" + ConvertForbiddenCharacters(sOptionalArgument7) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument8, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument8))
			{
				sLogXML = sLogXML + " arg8=\"" + ConvertForbiddenCharacters(sOptionalArgument8) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument9, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument9))
			{
				sLogXML = sLogXML + " arg9=\"" + ConvertForbiddenCharacters(sOptionalArgument9) + "\"";
			}
			if (Operators.CompareString(sOptionalArgument10, "", false) != 0 && !string.IsNullOrEmpty(sOptionalArgument10))
			{
				sLogXML = sLogXML + " arg10=\"" + ConvertForbiddenCharacters(sOptionalArgument10) + "\"";
			}
			sLogXML += "  /></part>";
		}
		return text2;
	}

	private string FinalizeXMlbeforePost(string sXMl)
	{
		return "<?xml version=\"1.0\" encoding=\"utf-8\"?> <data>" + sXMl + "</data>";
	}

	public void KPI_POST(string sXMLPayload)
	{
		_GeneralException = "";
		if (Operators.CompareString(sXMLPayload, "", false) != 0)
		{
			string text = Strings.Replace("%SERVERNAME%/api/echecker/KPI/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
			_SendXMLToServer(sXMLPayload, text);
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
			}
		}
		else
		{
			_GeneralException = "blank XML Payload";
		}
	}

	public void KPI_PUT(string sXMLPayload)
	{
		_GeneralException = "";
		if (Operators.CompareString(sXMLPayload, "", false) != 0)
		{
			string text = Strings.Replace("%SERVERNAME%/api/echecker/KPI/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
			_SendXMLToServer(sXMLPayload, text, "PUT");
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
			}
		}
		else
		{
			_GeneralException = "blank XML Payload";
		}
	}

	public void KPI_DELETE(string sXMLPayload)
	{
		_GeneralException = "";
		if (Operators.CompareString(sXMLPayload, "", false) != 0)
		{
			string text = Strings.Replace("%SERVERNAME%/api/echecker/KPI/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
			_SendXMLToServer(sXMLPayload, text, "DELETE");
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
			}
		}
		else
		{
			_GeneralException = "blank XML Payload";
		}
	}

	public ClsCollection KPI_GET(string sdpsnumber, string sdpsrev, string sArguments, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace("%SERVERNAME%/api/echecker/KPI/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = text + "?" + Strings.Replace(sArguments, " ", "%20", 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			return clsCollection;
		}
		XmlDocument xmlDocument = new XmlDocument();
		string key = "";
		xmlDocument.LoadXml(text2);
		XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/part");
		foreach (XmlElement item in xmlNodeList)
		{
			ClsCollection clsCollection2 = new ClsCollection();
			foreach (XmlAttribute attribute in item.Attributes)
			{
				if (Operators.CompareString(Strings.UCase(attribute.Name), "KPI_ID", false) == 0)
				{
					key = "#" + attribute.Value;
				}
				clsCollection2.Add(attribute.Name, attribute.Value);
			}
			XmlNodeList xmlNodeList2 = item.SelectNodes("kpiDesc");
			foreach (XmlElement item2 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute2 in item2.Attributes)
				{
					clsCollection2.Add(attribute2.Name, attribute2.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("dpsDesc");
			foreach (XmlElement item3 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute3 in item3.Attributes)
				{
					clsCollection2.Add(attribute3.Name, attribute3.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("edrnDesc");
			foreach (XmlElement item4 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute4 in item4.Attributes)
				{
					clsCollection2.Add(attribute4.Name, attribute4.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("rvDesc");
			foreach (XmlElement item5 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute5 in item5.Attributes)
				{
					clsCollection2.Add(attribute5.Name, attribute5.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("errorDesc");
			foreach (XmlElement item6 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute6 in item6.Attributes)
				{
					clsCollection2.Add(attribute6.Name, attribute6.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("auditor");
			foreach (XmlElement item7 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute7 in item7.Attributes)
				{
					clsCollection2.Add("auditor_" + attribute7.Name, attribute7.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("designer");
			foreach (XmlElement item8 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute8 in item8.Attributes)
				{
					clsCollection2.Add("designer_" + attribute8.Name, attribute8.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("manager");
			foreach (XmlElement item9 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute9 in item9.Attributes)
				{
					clsCollection2.Add("manager_" + attribute9.Name, attribute9.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("fileDesc");
			foreach (XmlElement item10 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute10 in item10.Attributes)
				{
					clsCollection2.Add(attribute10.Name, attribute10.Value);
				}
			}
			xmlNodeList2 = item.SelectNodes("documentDesc");
			foreach (XmlElement item11 in xmlNodeList2)
			{
				foreach (XmlAttribute attribute11 in item11.Attributes)
				{
					clsCollection2.Add(attribute11.Name, attribute11.Value);
				}
			}
			if (!clsCollection.Contains(key))
			{
				clsCollection.Add(key, clsCollection2);
			}
		}
		return clsCollection;
	}

	private void LogWebserviceUsage(string sFunction_name, string sURL, string sPartNumber = "", string sPartRev = "", object iOptionalPartIteration = null, string sOptionalValue = "", string sOptionalArgument4 = "", string sOptionalArgument5 = "", string sOptionalArgument6 = "", string sOptionalArgument7 = "", string sOptionalArgument8 = "")
	{
		checked
		{
			iLogLineCnt++;
			string sOptionalArgument9 = "";
			string sLog_Status;
			if ((Operators.CompareString(_WebClientReturnStatus, "", false) != 0) | (Operators.CompareString(_WebResponseExceptionStatus, "", false) != 0) | (Operators.CompareString(_GeneralException, "", false) != 0))
			{
				sLog_Status = "ERROR";
				sOptionalArgument9 = "WebClientReturnStatus=" + _WebClientReturnStatus + "|WebResponseExceptionStatus" + _WebResponseExceptionStatus + "|GeneralException" + _GeneralException;
			}
			else
			{
				sLog_Status = "OK";
			}
			if (!_ResultFromCache)
			{
				MonitorLogUsage(ref _sWebserviceLogPayload, "LOG", sLog_Status, WebserviceToolVersion, "", sFunction_name, sPartNumber, sPartRev, RuntimeHelpers.GetObjectValue(iOptionalPartIteration), sOptionalValue, _WebtransactionTime + "s", WebClientReturnStatus + WebResponseExceptionStatus, sURL, sOptionalArgument4, sOptionalArgument5, sOptionalArgument6, sOptionalArgument7, sOptionalArgument8, sOptionalArgument9, "Called From App:" + UsedbyApplication);
			}
			if ((iLogLineCnt >= 100) & (Operators.CompareString(_sWebserviceLogPayload, "", false) != 0))
			{
				_sWebserviceLogPayload = FinalizeXMlbeforePost(_sWebserviceLogPayload);
				string strURL = Strings.Replace("%SERVERNAME%/api/echecker/logtoolusage/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
				_SendXMLToServer(_sWebserviceLogPayload, strURL);
				_sWebserviceLogPayload = "";
				iLogLineCnt = 0;
			}
		}
	}

	public void LogPendingWebserviceUsageInCache()
	{
		if (Operators.CompareString(_sWebserviceLogPayload, "", false) != 0)
		{
			_sWebserviceLogPayload = FinalizeXMlbeforePost(_sWebserviceLogPayload);
			string strURL = Strings.Replace("%SERVERNAME%/api/echecker/logtoolusage/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
			_SendXMLToServer(_sWebserviceLogPayload, strURL);
			_sWebserviceLogPayload = "";
			iLogLineCnt = 0;
		}
	}

	private string ConvertForbiddenCharacters(string s)
	{
		return Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace(s, "&", "&amp;", 1, -1, (CompareMethod)0), "<", " &lt;", 1, -1, (CompareMethod)0), ">", "&gt;", 1, -1, (CompareMethod)0), "\"", "&quot;", 1, -1, (CompareMethod)0), "\v", "\n", 1, -1, (CompareMethod)0);
	}

	private string GetNetworkUserFullNameFromLogin(string sDomainName, string sLoginCode)
	{
		string text = "";
		try
		{
			string text2 = ".";
			text = Conversions.ToString(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(Interaction.GetObject("WINMGMTS:\\\\" + text2 + "\\ROOT\\cimv2", (string)null)), (Type)null, "Get", new object[1] { "Win32_UserAccount.Domain=\"" + sDomainName + "\",Name=\"" + sLoginCode + "\"" }, (string[])null, (Type[])null, (bool[])null)), (Type)null, "FullName", new object[0], (string[])null, (Type[])null, (bool[])null));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (Operators.CompareString(text, "", false) == 0)
		{
			object objectValue = RuntimeHelpers.GetObjectValue(Interaction.CreateObject("WScript.Network", ""));
			text = Conversions.ToString(NewLateBinding.LateGet(Interaction.GetObject(Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject((object)"WinNT://", NewLateBinding.LateGet(objectValue, (Type)null, "UserDomain", new object[0], (string[])null, (Type[])null, (bool[])null)), (object)"/"), NewLateBinding.LateGet(objectValue, (Type)null, "UserName", new object[0], (string[])null, (Type[])null, (bool[])null)), (object)",user")), (string)null), (Type)null, "FullName", new object[0], (string[])null, (Type[])null, (bool[])null));
			objectValue = null;
		}
		return text;
	}

	public string JSON_BDIPartInfo(string sBDIPartID, ref bool bexitbyuser, [Optional][DefaultParameterValue(false)] ref bool bIsCI, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/parts/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0) + sBDIPartID;
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string text2 = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		try
		{
			Hashtable hashtable = JsonConvert.DeserializeObject<Hashtable>(text2);
			bIsCI = Conversions.ToBoolean(hashtable["isCi"]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sBDIPARTID=" + sBDIPartID);
		}
		return text2;
	}

	public string JSON_BDIPartInfoFromPnRev(string sBDIPartNumber, string sRevision, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = ((Operators.CompareString(sRevision, "", false) != 0) ? Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/parts/?num=%BASENUMBER%&rev=%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sBDIPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRevision, 1, -1, (CompareMethod)0) : Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/parts/?num=%BASENUMBER%&rev=%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sBDIPartNumber, 1, -1, (CompareMethod)0), "&rev=%REVISION%", sRevision, 1, -1, (CompareMethod)0));
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string text2 = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		try
		{
			JsonConvert.DeserializeObject<Hashtable>(text2);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (EnableFUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sBDIPARTID=" + sBDIPartNumber);
		}
		return text2;
	}

	public string JSON_BDIGetUserInfo(ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "")
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/users/me", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult: true);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public string JSON_BDIDocumentInfo(string sBDIPartID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/documents/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0) + sBDIPartID;
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "BDIDocID=" + sBDIPartID);
		}
		return result;
	}

	public string JSON_BDIDocumentInfoFromPnRv(string sBDIPartNumber, string sRevision, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = ((Operators.CompareString(sRevision, "", false) != 0) ? Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/documents/?num=%BASENUMBER%&rev=%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sBDIPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRevision, 1, -1, (CompareMethod)0) : Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/documents/?num=%BASENUMBER%&rev=%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sBDIPartNumber, 1, -1, (CompareMethod)0), "&rev=%REVISION%", sRevision, 1, -1, (CompareMethod)0));
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string text2 = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		try
		{
			JsonConvert.DeserializeObject<Hashtable>(text2);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "BDIDocPartNumber=" + sBDIPartNumber);
		}
		return text2;
	}

	public string JSON_BDIDocumentRelationInfo(string sBDIPartID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/documents/%DOCID%/rvDocRels", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%DOCID%", sBDIPartID, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "BDIDocID=" + sBDIPartID);
		}
		return result;
	}

	public string JSON_RVWorkflow(string sRVid, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/rvs/%RVID%/workflow", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%RVID%", sRVid, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sRVid=" + sRVid);
		}
		return result;
	}

	public string JSON_DOCWorkFlow(string sDocId, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/documents/%DOCID%/workflow", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%DOCID%", sDocId, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sDocId=" + sDocId);
		}
		return result;
	}

	public string JSON_DOCEffectivity(string sDocumentNumber, string sDocumentRev, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/documents/%DocumentNumber%/%Rev%/effectivity", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%DocumentNumber%", sDocumentNumber, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%Rev%", sDocumentRev, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sDocumentNumber, sDocumentRev);
		}
		return result;
	}

	public string JSON_TailNumberFromProjectNumber(string sProjectNbs, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = "";
		string text2 = Strings.Replace("%SERVERNAME%/api/projects/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0) + sProjectNbs;
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			if (Operators.CompareString(sUserName, "", false) == 0)
			{
				sUserName = Environment.UserName;
			}
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string text3 = sGetWebResult(text2, sEncodedUsrAndPwd);
		try
		{
			Hashtable hashtable = JsonConvert.DeserializeObject<Hashtable>(text3);
			if (hashtable != null && hashtable.Contains("Project_Tail"))
			{
				text = Conversions.ToString(hashtable["Project_Tail"]);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			_GeneralException = ex2.Message;
			ProjectData.ClearProjectError();
		}
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2, "", "", null, "", "sProjectNbs=" + sProjectNbs, "TailNumber=" + text);
		}
		return text;
	}

	public ClsCollection JSON_GETBDIID(string sObjectType, string sObjectName, string sObjrev, ref bool bexitbyuser, ref object oResultList, int SearchResultMaxSize = 200, [Optional][DefaultParameterValue("")] ref string JResult, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_00e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Expected O, but got Unknown
		_GeneralException = "";
		Collection val = new Collection();
		Collection val2 = new Collection();
		Hashtable hashtable = new Hashtable();
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace("%SERVERNAME%/quicksearch/%BDIobjectType%/_search?q=keys:", "%BDIobjectType%", sObjectType, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0) + sObjectName + sObjrev + "&size=" + Conversions.ToString(SearchResultMaxSize);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		JResult = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		object objectValue;
		JArray val3;
		try
		{
			objectValue = RuntimeHelpers.GetObjectValue(JsonConvert.DeserializeObject(JResult));
			objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(NewLateBinding.LateIndexGet(objectValue, new object[1] { "hits" }, (string[])null), new object[1] { "hits" }, (string[])null));
			val3 = (JArray)objectValue;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			_GeneralException = ex2.Message;
			ProjectData.ClearProjectError();
			goto IL_0536;
		}
		checked
		{
			if (((JContainer)val3).Count > 0)
			{
				int num = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(objectValue, (Type)null, "count", new object[0], (string[])null, (Type[])null, (bool[])null), (object)1));
				for (int i = 0; i <= num; i++)
				{
					Hashtable hashtable2 = JsonConvert.DeserializeObject<Hashtable>(val3[i][(object)"_source"].ToString());
					try
					{
						if (!hashtable2.Contains("title"))
						{
							continue;
						}
						if (Operators.CompareString(sObjrev, "", false) != 0)
						{
							if (Operators.CompareString(Strings.UCase(Strings.Replace(Strings.Replace(Strings.Replace(Conversions.ToString(hashtable2["title"]), " ", "", 1, -1, (CompareMethod)0), "[", "", 1, -1, (CompareMethod)0), "]", "", 1, -1, (CompareMethod)0)), Strings.UCase(sObjectName + sObjrev), false) == 0 && !clsCollection.Contains(RuntimeHelpers.GetObjectValue(hashtable2["title"])))
							{
								clsCollection.Add(RuntimeHelpers.GetObjectValue(hashtable2["title"]), RuntimeHelpers.GetObjectValue(hashtable2["id"]));
							}
						}
						else if (Operators.CompareString(Strings.UCase(Strings.Replace(Strings.Split(Conversions.ToString(hashtable2["title"]), "[", -1, (CompareMethod)0)[0], " ", "", 1, -1, (CompareMethod)0)), Strings.UCase(sObjectName), false) == 0)
						{
							string text2 = "";
							if (Strings.InStr(Conversions.ToString(hashtable2["title"]), "[", (CompareMethod)0) > 0)
							{
								text2 = Strings.Right(Conversions.ToString(hashtable2["title"]), Strings.Len(RuntimeHelpers.GetObjectValue(hashtable2["title"])) - Strings.InStr(Conversions.ToString(hashtable2["title"]), "[", (CompareMethod)0));
								text2 = Strings.Trim(Strings.Left(text2, Strings.InStr(text2, "]", (CompareMethod)0) - 1));
							}
							if (!clsCollection.Contains(RuntimeHelpers.GetObjectValue(hashtable2["title"])))
							{
								clsCollection.Add(RuntimeHelpers.GetObjectValue(hashtable2["title"]), RuntimeHelpers.GetObjectValue(hashtable2["id"]));
							}
							if (!hashtable.Contains(text2))
							{
								hashtable.Add(text2, "");
								val2.Add((object)(text2 + "|" + Conversions.ToString(hashtable2["id"])), (string)null, (object)null, (object)null);
								val.Add((object)(Strings.Replace(Strings.Replace(Strings.Replace(Conversions.ToString(hashtable2["title"]), "[", "", 1, -1, (CompareMethod)0), "]", "", 1, -1, (CompareMethod)0), " ", "", 1, -1, (CompareMethod)0) + "|" + Conversions.ToString(hashtable2["id"])), (string)null, (object)null, (object)null);
							}
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (val2.Count > 1)
			{
				int num2 = val2.Count - 1;
				for (int i = 1; i <= num2; i++)
				{
					int num3 = i + 1;
					int count = val2.Count;
					for (int j = num3; j <= count; j++)
					{
						if (Operators.CompareString(Strings.Split(Conversions.ToString(val2[i]), "|", -1, (CompareMethod)0)[0], Strings.Split(Conversions.ToString(val2[j]), "|", -1, (CompareMethod)0)[0], false) > 0)
						{
							string text3 = Conversions.ToString(val2[j]);
							val2.Remove(j);
							val2.Add((object)text3, (string)null, (object)i, (object)null);
							string text4 = Conversions.ToString(val[j]);
							val.Remove(j);
							val.Add((object)text4, (string)null, (object)i, (object)null);
						}
					}
				}
			}
			oResultList = val;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sObjectName, sObjrev);
			}
			goto IL_0536;
		}
		IL_0536:
		return clsCollection;
	}

	public string JSON_ListOfAttachmentsDPS(string sDocId, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/bdicommon/attachments/210/%DocumentID%/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%DocumentID%", sDocId, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sDocId=" + sDocId);
		}
		return result;
	}

	public string JSON_DocumentEffectivity(string DocumentNumber, string Rev, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false, bool GetEffTxtOnly = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/documents/%PartNumber%/%Rev%/effectivity", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%PartNumber%", DocumentNumber, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%Rev%", Rev, 1, -1, (CompareMethod)0);
		if (GetEffTxtOnly)
		{
			text += "/display?format=text";
		}
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "DocumentNumber=" + DocumentNumber, "Rev=" + Rev);
		}
		return result;
	}

	public string JSON_BDIPartChildren(string PartID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/parts/%PartID%/childrens/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%PartID%", PartID, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public string JSON_BDIPartDPSDocumments(string PartID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/parts/%PartID%/documents/", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%PartID%", PartID, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public string JSON_BDIPARTFirstLevelChildren(string PartID, string ACID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/ebom/strings/getPartChildrenListWithoutStringLevel?eBom_LibItemId=%PARTID%&inAcFilter_id=%ACID%&outAcFilter_id=%ACID%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		if (Operators.CompareString(ACID, "", false) == 0)
		{
			text = text.Split(new char[1] { '%' }).First() + "%PartID%";
		}
		text = Strings.Replace(Strings.Replace(text, "%PARTID%", PartID, 1, -1, (CompareMethod)0), "%ACID%", ACID, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public string JSON_BDIPARTStrings(string PartID, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/parts/%PARTID%/strings", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%PARTID%", PartID, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public string DownloadAttachment(string sAttachmentId, ref bool bexitbyuser, string sUserName = "", string sEncodedUsrAndPwd = "")
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/bdicommon/attachments/%AttachmentID%/file", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%AttachmentID%", sAttachmentId, 1, -1, (CompareMethod)0);
		if (Operators.CompareString(sEncodedUsrAndPwd, "", false) == 0)
		{
			sEncodedUsrAndPwd = GetUserPassword(sUserName, ref bexitbyuser);
		}
		string result = sGetWebResult(text, sEncodedUsrAndPwd, GetUncachedResult: true);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, "", "sAttachmentId=" + sAttachmentId);
		}
		return result;
	}

	public ClsCollection GetNIEOData(string sPartNumber, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		ClsCollection clsCollection2 = new ClsCollection();
		ClsCollection clsCollection3 = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/nieo/%PARTNUMBER%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		checked
		{
			ClsCollection result;
			if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
			{
				result = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.LoadXml(text2);
				XmlNode xmlNode = xmlDocument.SelectSingleNode("/data");
				foreach (XmlElement childNode in xmlNode.ChildNodes)
				{
					clsCollection.Add(childNode.Name, new ClsCollection());
					foreach (object attribute in childNode.Attributes)
					{
						object objectValue = RuntimeHelpers.GetObjectValue(attribute);
						object obj = clsCollection.get_GetItem((object)childNode.Name);
						object[] array = new object[2];
						object obj2 = objectValue;
						array[0] = NewLateBinding.LateGet(obj2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
						object obj3 = objectValue;
						array[1] = NewLateBinding.LateGet(obj3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null);
						object[] array2 = array;
						bool[] array3;
						NewLateBinding.LateCall(obj, (Type)null, "Add", array, (string[])null, (Type[])null, array3 = new bool[2] { true, true }, true);
						if (array3[0])
						{
							NewLateBinding.LateSetComplex(obj2, (Type)null, "name", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
						}
						if (array3[1])
						{
							NewLateBinding.LateSetComplex(obj3, (Type)null, "Value", new object[1] { array2[1] }, (string[])null, (Type[])null, true, false);
						}
					}
					int num = 0;
					foreach (XmlElement childNode2 in childNode.ChildNodes)
					{
						num++;
						clsCollection2 = new ClsCollection();
						string text3 = ((Operators.CompareString(childNode2.Name, "signatories", false) != 0) ? (childNode2.Name + Conversions.ToString(num)) : childNode2.Name);
						foreach (object attribute2 in childNode2.Attributes)
						{
							object objectValue = RuntimeHelpers.GetObjectValue(attribute2);
							clsCollection2.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
						}
						int num2 = 0;
						foreach (XmlElement childNode3 in childNode2.ChildNodes)
						{
							num2++;
							clsCollection3 = new ClsCollection();
							foreach (object attribute3 in childNode3.Attributes)
							{
								object objectValue = RuntimeHelpers.GetObjectValue(attribute3);
								clsCollection3.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
							}
							clsCollection2.Add(childNode3.Name + Conversions.ToString(num2), clsCollection3);
						}
						object[] array2;
						bool[] array3;
						NewLateBinding.LateCall(clsCollection.get_GetItem((object)childNode.Name), (Type)null, "add", array2 = new object[2] { text3, clsCollection2 }, (string[])null, (Type[])null, array3 = new bool[2] { true, true }, true);
						if (array3[0])
						{
							text3 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(string));
						}
						if (array3[1])
						{
							clsCollection2 = (ClsCollection)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(ClsCollection));
						}
					}
				}
				result = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber);
				}
			}
			return result;
		}
	}

	public ClsCollection GetEDRNData(string sdpsnumber, string sdpsrev, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/edrn/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sdpsnumber, 1, -1, (CompareMethod)0), "%REVISION%", sdpsrev, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/edrn");
			clsCollection.Add(xmlNode.Name, new ClsCollection());
			foreach (object attribute in xmlNode.Attributes)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(attribute);
				clsCollection.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sdpsnumber, sdpsrev);
			}
		}
		return result;
	}

	public ClsCollection GetBDIData(string sdpsnumber, string sdpsrev, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/edrn/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sdpsnumber, 1, -1, (CompareMethod)0), "%REVISION%", sdpsrev, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		checked
		{
			ClsCollection result;
			if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
			{
				result = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.LoadXml(text2);
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/relatedParts/part");
				ClsCollection clsCollection2 = new ClsCollection();
				foreach (XmlElement item in xmlNodeList)
				{
					ClsCollection clsCollection3 = new ClsCollection();
					foreach (object attribute in item.Attributes)
					{
						object objectValue = RuntimeHelpers.GetObjectValue(attribute);
						clsCollection3.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection2.Add("RelatedPart_" + Conversions.ToString(clsCollection2.Count + 1), clsCollection3);
				}
				clsCollection.Add("RelatedParts", clsCollection2);
				ClsCollection clsCollection4 = new ClsCollection();
				XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/dps_attributes");
				foreach (object attribute2 in xmlNode.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute2);
					clsCollection4.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				clsCollection.Add("dps_attributes", clsCollection4);
				ClsCollection clsCollection5 = new ClsCollection();
				xmlNode = xmlDocument.SelectSingleNode("/data/workflowStatus");
				foreach (object attribute3 in xmlNode.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute3);
					clsCollection5.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				xmlNodeList = xmlDocument.SelectNodes("/data/workflowStatus/signatories");
				ClsCollection clsCollection6 = new ClsCollection();
				foreach (XmlElement item2 in xmlNodeList)
				{
					ClsCollection clsCollection7 = new ClsCollection();
					foreach (object attribute4 in item2.Attributes)
					{
						object objectValue = RuntimeHelpers.GetObjectValue(attribute4);
						clsCollection7.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection6.Add("signatory_" + Conversions.ToString(clsCollection6.Count + 1), clsCollection7);
				}
				clsCollection5.Add("signatories", clsCollection6);
				clsCollection.Add("workflowStatus", clsCollection5);
				result = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sdpsnumber, sdpsrev);
				}
			}
			return result;
		}
	}

	public ClsCollection GetBADeliverableDataFromDocument(string sPartNumber, string sRevision, ref object oCol, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/BADeliverable/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%BASENUMBER%", sPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRevision, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/part");
			clsCollection.Add(xmlNode.Name, new ClsCollection());
			foreach (object attribute in xmlNode.Attributes)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(attribute);
				clsCollection.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
			}
			oCol = clsCollection;
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber, sRevision);
			}
		}
		return result;
	}

	public ClsCollection GetAttachedDocumentstoBADeliverable(string sBADeliverable)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/attachedDocumentsByBADeliverable/%ACTIONNUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%ACTIONNUMBER%", sBADeliverable, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/document");
			string text3 = "";
			string text4 = "";
			foreach (object item in xmlNodeList)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(item);
				text3 = "";
				text4 = "";
				foreach (object item2 in (IEnumerable)NewLateBinding.LateGet(objectValue, (Type)null, "Attributes", new object[0], (string[])null, (Type[])null, (bool[])null))
				{
					object objectValue2 = RuntimeHelpers.GetObjectValue(item2);
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"FIELD_PART_NUMBER", false))
					{
						text3 = Conversions.ToString(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"FIELD_DOCUMENT_REVISION", false))
					{
						text4 = Conversions.ToString(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					if ((Operators.CompareString(text3, "", false) != 0) & (Operators.CompareString(text4, "", false) != 0))
					{
						clsCollection.Add(text3, text4);
					}
				}
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sBADeliverable);
			}
		}
		return result;
	}

	public ClsCollection GetPLMActionDataFromDocument(string sPartNumber, string sRevision, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/enovia/plmaction/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%BASENUMBER%", sPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRevision, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/plmaction");
			foreach (object attribute in xmlNode.Attributes)
			{
				object objectValue = RuntimeHelpers.GetObjectValue(attribute);
				clsCollection.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber, sRevision);
			}
		}
		return result;
	}

	public ClsCollection GetENOVIADocumentAttributs(string sPartNumber, string sRevision, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/enovia/document/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%BASENUMBER%", sPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRevision, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode oXMLNode = xmlDocument.SelectSingleNode("/data/document");
			clsCollection = ExtractDocAttributesFromXML(oXMLNode);
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber, sRevision);
			}
		}
		return result;
	}

	private ClsCollection ExtractDocAttributesFromXML(object oXMLNode)
	{
		ClsCollection clsCollection = new ClsCollection();
		checked
		{
			foreach (object item in (IEnumerable)NewLateBinding.LateGet(oXMLNode, (Type)null, "attributes", new object[0], (string[])null, (Type[])null, (bool[])null))
			{
				object objectValue = RuntimeHelpers.GetObjectValue(item);
				object obj = NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
				if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_PART_NUMBER", false))
				{
					clsCollection.Add("Part Number", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_BASE_NUMBER", false))
				{
					clsCollection.Add("Base Number", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DASH_NUMBER", false))
				{
					clsCollection.Add("Dash Number", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DOCUMENT_REVISION", false))
				{
					clsCollection.Add("BA Document Revision", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DOCUMENT_ITERATION", false))
				{
					clsCollection.Add("DOCUMENT_ITERATION", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_STATUS", false))
				{
					clsCollection.Add("Revision Status", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_CONFIDENTIAL", false))
				{
					clsCollection.Add("CONFIDENTIAL", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_TITLE", false))
				{
					clsCollection.Add("Title", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DATASET_TYPE", false))
				{
					clsCollection.Add("Dataset Type", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MFG_PROGRESS", false))
				{
					clsCollection.Add("MFG Process", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ATA_CHAPTER_SECTION_SNS", false))
				{
					clsCollection.Add("ATA Chapter Section / SNS", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ALLOY", false))
				{
					clsCollection.Add("Alloy", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_CI", false))
				{
					clsCollection.Add("CI", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_CONDITION_STATEMENT", false))
				{
					clsCollection.Add("Condition Statement", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DAL", false))
				{
					clsCollection.Add("DAL", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DENSITY", false))
				{
					clsCollection.Add("Density", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DESIGN_AUTHORITY_PROGRAM", false))
				{
					clsCollection.Add("Design Authority Program", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DOCUMENT_TYPE", false))
				{
					clsCollection.Add("Document Type", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ENG_MAKE_FROM", false))
				{
					clsCollection.Add("Eng. Make From", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_EXTENSION", false))
				{
					clsCollection.Add("EXTENSION", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_FAR", false))
				{
					clsCollection.Add("FAR", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_FINAL_CONDITION", false))
				{
					clsCollection.Add("Final Condition", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_FINISH_CODE", false))
				{
					clsCollection.Add("Finish Code", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_FORM", false))
				{
					clsCollection.Add("Form", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_GRADE_COMPOSITION", false))
				{
					clsCollection.Add("Grade/Composition", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_INSIDE_DIAMETER", false))
				{
					clsCollection.Add("Inside Diameter", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - (string.IsNullOrEmpty(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))) ? 1 : 0)) + 2), new object[2]
					{
						0,
						NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_INTERCHANGEABILITY_CODE", false))
				{
					clsCollection.Add("Interchangeability Code", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_INTERCHANGEABILITY_PARTS", false))
				{
					clsCollection.Add("Interchangeability Parts", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - ((Operators.CompareString(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)), "0", false) == 0) ? 1 : 0)) + 2), new object[2]
					{
						null,
						Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_IRCS", false))
				{
					clsCollection.Add("I/R/C/S", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_KEY_CHARACTERISTICS", false))
				{
					clsCollection.Add("KEY_CHARACTERISTICS", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LAST_MODIFIED_BY", false))
				{
					clsCollection.Add("LAST_MODIFIED_BY", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LENGTH", false))
				{
					clsCollection.Add("Length", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - (string.IsNullOrEmpty(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))) ? 1 : 0)) + 2), new object[2]
					{
						0,
						NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LFCRN", false))
				{
					clsCollection.Add("LFCRN", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LIFE_LIMITED_PART", false))
				{
					clsCollection.Add("LIFE_LIMITED_PART", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LRU", false))
				{
					clsCollection.Add("LRU", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MAJOR_SUPPLIER_CODE", false))
				{
					clsCollection.Add("Major Supplier Code", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_CLASS", false))
				{
					clsCollection.Add("Material Class", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_DESC_PROD", false))
				{
					clsCollection.Add("Material Description Production", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_FORM", false))
				{
					clsCollection.Add("Material Form", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_SPEC_PROD", false))
				{
					clsCollection.Add("Material Specification Production", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_TYPE", false))
				{
					clsCollection.Add("Material Type", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MESH_CELL_SIZE", false))
				{
					clsCollection.Add("Mesh Cell Size", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MTBF", false))
				{
					clsCollection.Add("MTBF", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MTBUR", false))
				{
					clsCollection.Add("MTBUR", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_NIEO_NUM", false))
				{
					clsCollection.Add("NIEO_NUM", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_NOMENCLATURE", false))
				{
					clsCollection.Add("Nomenclature", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_OID", false))
				{
					clsCollection.Add("OID", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_OUTSIDE_DIAMETER", false))
				{
					clsCollection.Add("Outside Diameter", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - (string.IsNullOrEmpty(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))) ? 1 : 0)) + 2), new object[2]
					{
						0,
						NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_PCCN", false))
				{
					clsCollection.Add("PCCN", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_PROTECTED_INTERFACE", false))
				{
					clsCollection.Add("Protected Interface", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_QAN", false))
				{
					clsCollection.Add("QAN", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_QUALITY_LEVEL", false))
				{
					clsCollection.Add("QUALITY_LEVEL", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_CREATION_DATE", false))
				{
					clsCollection.Add("REV_CREATION_DATE", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_CREATOR", false))
				{
					clsCollection.Add("REV_CREATOR", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_LAST_MOD_DATE", false))
				{
					clsCollection.Add("REV_LAST_MOD_DATE", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_ORGANIZATION", false))
				{
					clsCollection.Add("Revision Organization", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_OWNER", false))
				{
					clsCollection.Add("REV_OWNER", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_REV_PROJECT", false))
				{
					clsCollection.Add("Revision Project", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SECURITYCHECK", false))
				{
					clsCollection.Add("Security Check", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SERIALIZATION", false))
				{
					clsCollection.Add("Serialized", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SIZE", false))
				{
					clsCollection.Add("Size", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_STANDARD_SPEC_DIE", false))
				{
					clsCollection.Add("Standard Spec Die", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SUB_STATUS", false))
				{
					clsCollection.Add("Sub-Status", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SUPPLIER_NAME_CAGE_CODE", false))
				{
					clsCollection.Add("Supplier Name And CAGE Code", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_TD_MATERIAL_CODE", false))
				{
					clsCollection.Add("TD Material Code", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_THICKNESS", false))
				{
					clsCollection.Add("Thickness", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_WALL", false))
				{
					clsCollection.Add("Wall", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - (string.IsNullOrEmpty(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))) ? 1 : 0)) + 2), new object[2]
					{
						0,
						NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_WIDTH", false))
				{
					clsCollection.Add("Width", RuntimeHelpers.GetObjectValue(Interaction.Choose((double)(unchecked(0 - (string.IsNullOrEmpty(Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null))) ? 1 : 0)) + 2), new object[2]
					{
						0,
						NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)
					})));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_WIRE_GAUGE", false))
				{
					clsCollection.Add("Wire Gauge", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DEFINING_PART", false))
				{
					clsCollection.Add("Defining Part", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_3D_ONLY", false))
				{
					clsCollection.Add("3D Only", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_COLOR_CODED", false))
				{
					clsCollection.Add("Color Coded", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_PROPULSION_SYSTEM", false))
				{
					clsCollection.Add("Propulsion System", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_TYPE", false))
				{
					clsCollection.Add("Type", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_STYLE", false))
				{
					clsCollection.Add("Style", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_SPECIFICATIONS", false))
				{
					clsCollection.Add("Material Specifications", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_MATERIAL_DESCRIPTION", false))
				{
					clsCollection.Add("Material Description", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_FT_LOCATION_ZONE", false))
				{
					clsCollection.Add("FT Location Zone", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_SHAREABLE", false))
				{
					clsCollection.Add("Shareable", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_DOCUMENT_ORGANIZATION", false))
				{
					clsCollection.Add("Document Organization", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_LINKED_TO_PARTS", false))
				{
					clsCollection.Add("Linked to Parts", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_OID", false))
				{
					clsCollection.Add("OID", RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
			}
			return clsCollection;
		}
	}

	public object GetAssyPositionMatricies(string sPartNumber, bool GetUncachedResult = false)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		//IL_03fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0404: Expected O, but got Unknown
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00da: Expected O, but got Unknown
		//IL_04ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d1: Expected O, but got Unknown
		_GeneralException = "";
		Collection val = new Collection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/enovia/assyRelations/%BASENUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%BASENUMBER%", sPartNumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		object result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = val;
		}
		else
		{
			double[] array = new double[12];
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/childPart");
			foreach (XmlElement item in xmlNodeList)
			{
				try
				{
					if (!val.Contains(item.GetAttribute("FIELD_PART_NUMBER")))
					{
						val.Add((object)new Collection(), item.GetAttribute("FIELD_PART_NUMBER"), (object)null, (object)null);
						NewLateBinding.LateCall(val[item.GetAttribute("FIELD_PART_NUMBER")], (Type)null, "Add", new object[2]
						{
							item.GetAttribute("FIELD_PART_NUMBER"),
							item.GetAttribute("FIELD_PART_NUMBER")
						}, (string[])null, (Type[])null, (bool[])null, true);
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					_GeneralException = ex2.Message;
					ProjectData.ClearProjectError();
				}
				foreach (object attribute in item.FirstChild.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute);
					object obj = NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
					if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM01", false))
					{
						array[0] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM02", false))
					{
						array[1] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM03", false))
					{
						array[2] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM04", false))
					{
						array[3] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM05", false))
					{
						array[4] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM06", false))
					{
						array[5] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM07", false))
					{
						array[6] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM08", false))
					{
						array[7] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM09", false))
					{
						array[8] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM10", false))
					{
						array[9] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM11", false))
					{
						array[10] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					else if (Operators.ConditionalCompareObjectEqual(obj, (object)"FIELD_ARM12", false))
					{
						array[11] = Conversions.ToDouble(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
				}
				Collection val2 = new Collection();
				val2.Add((object)(item.GetAttribute("FIELD_PART_NUMBER") + "." + Conversions.ToString(NewLateBinding.LateGet(val[item.GetAttribute("FIELD_PART_NUMBER")], (Type)null, "Count", new object[0], (string[])null, (Type[])null, (bool[])null))), (string)null, (object)null, (object)null);
				val2.Add(RuntimeHelpers.GetObjectValue(array.Clone()), (string)null, (object)null, (object)null);
				object[] array2;
				bool[] array3;
				NewLateBinding.LateCall(val[item.GetAttribute("FIELD_PART_NUMBER")], (Type)null, "Add", array2 = new object[2]
				{
					val2,
					val2[1]
				}, (string[])null, (Type[])null, array3 = new bool[2] { true, false }, true);
				if (array3[0])
				{
					val2 = (Collection)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(Collection));
				}
			}
			result = val;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber);
			}
		}
		return result;
	}

	public ClsCollection GetStringVsEBOMData(string sPartNumber, string sProjectNumber, [Optional][DefaultParameterValue("")] ref string StringVsEBOM_ListValidPartsAsString, [Optional][DefaultParameterValue("")] ref string StringVsEBOM_ListInValidPartsAsString, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/assy/%PARTNUMBER%/%PROJECTNUMBER%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0), "%PROJECTNUMBER%", sProjectNumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		checked
		{
			ClsCollection result;
			if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
			{
				result = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				string text3 = "";
				string text4 = "";
				string text5 = "";
				string text6 = "";
				xmlDocument.LoadXml(text2);
				XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/item");
				foreach (object attribute in xmlNode.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute);
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"NUMBER", false))
					{
						text3 = Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"LATEST_REV", false))
					{
						text4 = Conversions.ToString(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
					}
					clsCollection.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/relatedStrings/string");
				ClsCollection clsCollection2 = new ClsCollection();
				string text7 = "";
				foreach (XmlElement item in xmlNodeList)
				{
					ClsCollection clsCollection3 = new ClsCollection();
					ClsCollection clsCollection4 = new ClsCollection();
					foreach (object attribute2 in item.Attributes)
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(attribute2);
						if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"EBOM_STRING", false))
						{
							text7 = Conversions.ToString(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
						}
						clsCollection4.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection3.Add("Attributes", clsCollection4);
					ClsCollection clsCollection5 = new ClsCollection();
					XmlNodeList xmlNodeList2 = item.SelectNodes("children/childPart");
					foreach (XmlElement item2 in xmlNodeList2)
					{
						ClsCollection clsCollection6 = new ClsCollection();
						foreach (object attribute3 in item2.Attributes)
						{
							object objectValue3 = RuntimeHelpers.GetObjectValue(attribute3);
							clsCollection6.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
						}
						clsCollection5.Add("Children_" + Conversions.ToString(clsCollection5.Count + 1), clsCollection6);
					}
					clsCollection3.Add("ChildPart_" + Conversions.ToString(clsCollection3.Count + 1), clsCollection5);
					if (!clsCollection2.get_Exists(text7))
					{
						clsCollection2.Add(text7, clsCollection3);
					}
					else
					{
						clsCollection2.Add("String_" + Conversions.ToString(clsCollection2.Count + 1), clsCollection3);
					}
				}
				clsCollection.Add("RelatedStrings", clsCollection2);
				if (xmlNodeList.Count > 0)
				{
					text5 = "List of Valid Parts for " + text3 + " " + text4 + ":";
					text6 = "List of Invalid Parts for " + text3 + " " + text4 + ":";
				}
				xmlNodeList = xmlDocument.SelectNodes("/data/validParts/part");
				ClsCollection clsCollection7 = new ClsCollection();
				foreach (XmlElement item3 in xmlNodeList)
				{
					ClsCollection clsCollection8 = new ClsCollection();
					ClsCollection clsCollection9 = new ClsCollection();
					foreach (object attribute4 in item3.Attributes)
					{
						object objectValue4 = RuntimeHelpers.GetObjectValue(attribute4);
						clsCollection9.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection8.Add("Attributes", clsCollection9);
					if (Operators.CompareString(text5, "", false) != 0)
					{
						text5 += "\r\n";
					}
					text5 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject((object)(text5 + "Part: "), clsCollection9.get_GetItem((object)"PART_NUMBER")), (object)", EBOM Cnt: "), clsCollection9.get_GetItem((object)"BDI_COUNT")), (object)", ASSY Cnt: "), clsCollection9.get_GetItem((object)"ENO_COUNT")));
					ClsCollection clsCollection10 = new ClsCollection();
					XmlNodeList xmlNodeList2 = item3.GetElementsByTagName("applicableStrings");
					if (xmlNodeList2.Count == 1)
					{
						XmlElement xmlElement = (XmlElement)xmlNodeList2.Item(0);
						xmlNodeList2 = xmlElement.GetElementsByTagName("string");
					}
					if (xmlNodeList2.Count != 0)
					{
						foreach (object item4 in xmlNodeList2)
						{
							object objectValue5 = RuntimeHelpers.GetObjectValue(item4);
							ClsCollection clsCollection6 = new ClsCollection();
							foreach (object item5 in (IEnumerable)NewLateBinding.LateGet(objectValue5, (Type)null, "attributes", new object[0], (string[])null, (Type[])null, (bool[])null))
							{
								object objectValue6 = RuntimeHelpers.GetObjectValue(item5);
								clsCollection6.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue6, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue6, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
							}
							clsCollection10.Add("String_" + Conversions.ToString(clsCollection10.Count + 1), clsCollection6);
							text5 = Conversions.ToString(Operators.ConcatenateObject((object)(text5 + "\r\n\t--> Found in: "), clsCollection6.get_GetItem((object)"EBOM_STRING")));
						}
						clsCollection8.Add("ApplicableString_" + Conversions.ToString(clsCollection8.Count + 1), clsCollection10);
					}
					clsCollection7.Add("ValidPart_" + Conversions.ToString(clsCollection7.Count + 1), clsCollection8);
				}
				clsCollection.Add("ValidParts", clsCollection7);
				xmlNodeList = xmlDocument.GetElementsByTagName("invalidParts");
				if (xmlNodeList.Count == 1)
				{
					XmlElement xmlElement = (XmlElement)xmlNodeList.Item(0);
					xmlNodeList = xmlElement.GetElementsByTagName("part");
				}
				ClsCollection clsCollection11 = new ClsCollection();
				if (xmlNodeList.Count != 0)
				{
					foreach (XmlElement item6 in xmlNodeList)
					{
						ClsCollection clsCollection8 = new ClsCollection();
						ClsCollection clsCollection9 = new ClsCollection();
						foreach (object attribute5 in item6.Attributes)
						{
							object objectValue7 = RuntimeHelpers.GetObjectValue(attribute5);
							clsCollection9.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue7, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue7, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
						}
						clsCollection8.Add("Attributes", clsCollection9);
						if (Operators.CompareString(text6, "", false) != 0)
						{
							text6 += "\r\n";
						}
						text6 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject((object)(text6 + "Part: "), clsCollection9.get_GetItem((object)"PART_NUMBER")), (object)", EBOM Cnt: "), clsCollection9.get_GetItem((object)"BDI_COUNT")), (object)", ASSY Cnt: "), clsCollection9.get_GetItem((object)"ENO_COUNT")));
						ClsCollection clsCollection10 = new ClsCollection();
						XmlNodeList xmlNodeList2 = item6.GetElementsByTagName("applicableStrings");
						if (xmlNodeList2.Count == 1)
						{
							XmlElement xmlElement = (XmlElement)xmlNodeList2.Item(0);
							xmlNodeList2 = xmlElement.GetElementsByTagName("string");
						}
						if (xmlNodeList2.Count != 0)
						{
							foreach (object item7 in xmlNodeList2)
							{
								object objectValue8 = RuntimeHelpers.GetObjectValue(item7);
								ClsCollection clsCollection6 = new ClsCollection();
								foreach (object item8 in (IEnumerable)NewLateBinding.LateGet(objectValue8, (Type)null, "attributes", new object[0], (string[])null, (Type[])null, (bool[])null))
								{
									object objectValue9 = RuntimeHelpers.GetObjectValue(item8);
									clsCollection6.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue9, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue9, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
								}
								clsCollection10.Add("String_" + Conversions.ToString(clsCollection10.Count + 1), clsCollection6);
								text6 = Conversions.ToString(Operators.ConcatenateObject((object)(text6 + "\r\n\t--> Found in: "), clsCollection6.get_GetItem((object)"EBOM_STRING")));
							}
							clsCollection8.Add("ApplicableString_" + Conversions.ToString(clsCollection8.Count + 1), clsCollection10);
						}
						clsCollection11.Add("InValidPart_" + Conversions.ToString(clsCollection11.Count + 1), clsCollection8);
					}
				}
				clsCollection.Add("InValidParts", clsCollection11);
				StringVsEBOM_ListValidPartsAsString = text5;
				StringVsEBOM_ListInValidPartsAsString = text6;
				result = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber, "", null, "", "Project=" + sProjectNumber);
				}
			}
			return result;
		}
	}

	public string GetEBOMStringContentRawData(string sFamilyID, string sMonumentID, string sVariantID, string sEnvelopID, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = Strings.Replace("%SERVERNAME%/api/reports/ebom/configstrings/%FAMILYID%/%MONUMENTID%/%VARIANTID%/%ENVELOPID%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%FAMILYID%", sFamilyID, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%MONUMENTID%", sMonumentID, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%VARIANTID%", sVariantID, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%ENVELOPID%", sEnvelopID, 1, -1, (CompareMethod)0);
		string result = sGetWebResult(text, "", GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
		}
		return result;
	}

	public ClsCollection GetDocumentByBaseNumber(string sBasenumber, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/documentsByBaseNumber/%BASENUMBER%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sBasenumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/Document");
			foreach (XmlElement item in xmlNodeList)
			{
				string text3 = item.GetAttribute("FIELD_PART_NUMBER") + item.GetAttribute("FIELD_DOCUMENT_REVISION");
				clsCollection.Add(text3, new ClsCollection());
				foreach (object attribute in item.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute);
					object obj = clsCollection.get_GetItem((object)text3);
					object[] array = new object[2];
					object obj2 = objectValue;
					array[0] = NewLateBinding.LateGet(obj2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
					object obj3 = objectValue;
					array[1] = NewLateBinding.LateGet(obj3, (Type)null, "value", new object[0], (string[])null, (Type[])null, (bool[])null);
					object[] array2 = array;
					bool[] array3;
					NewLateBinding.LateCall(obj, (Type)null, "add", array, (string[])null, (Type[])null, array3 = new bool[2] { true, true }, true);
					if (array3[0])
					{
						NewLateBinding.LateSetComplex(obj2, (Type)null, "name", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
					}
					if (array3[1])
					{
						NewLateBinding.LateSetComplex(obj3, (Type)null, "value", new object[1] { array2[1] }, (string[])null, (Type[])null, true, false);
					}
				}
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sBasenumber);
			}
		}
		return result;
	}

	public ClsCollection GetRVData(string sRVnumber, string sRVrev, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/rv/%RVNUMBER%/%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%RVNUMBER%", sRVnumber, 1, -1, (CompareMethod)0), "%REVISION%", sRVrev, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		checked
		{
			ClsCollection result;
			if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
			{
				result = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.LoadXml(text2);
				XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/rv");
				foreach (object attribute in xmlNode.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute);
					clsCollection.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/rv/authority");
				string text3 = "";
				string text4 = "";
				string text5 = "";
				ClsCollection clsCollection2 = new ClsCollection();
				ClsCollection clsCollection3 = new ClsCollection();
				foreach (XmlElement item in xmlNodeList)
				{
					ClsCollection clsCollection4 = new ClsCollection();
					foreach (object attribute2 in item.Attributes)
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(attribute2);
						clsCollection4.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					if (Operators.CompareString(text3, "", false) != 0)
					{
						text3 += "; ";
					}
					text3 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject((object)text3, clsCollection4.get_GetItem((object)"auth_type")), (object)" "), clsCollection4.get_GetItem((object)"auth_num")));
					XmlNodeList xmlNodeList2 = item.SelectNodes("dps");
					ClsCollection clsCollection5 = new ClsCollection();
					foreach (XmlElement item2 in xmlNodeList2)
					{
						ClsCollection clsCollection6 = new ClsCollection();
						text4 = "";
						foreach (object attribute3 in item2.Attributes)
						{
							object objectValue3 = RuntimeHelpers.GetObjectValue(attribute3);
							Type typeFromHandle = typeof(Strings);
							object[] array = new object[1];
							object obj = objectValue3;
							array[0] = NewLateBinding.LateGet(obj, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
							object[] array2 = array;
							bool[] array3;
							object obj2 = NewLateBinding.LateGet((object)null, typeFromHandle, "LCase", array, (string[])null, (Type[])null, array3 = new bool[1] { true });
							if (array3[0])
							{
								NewLateBinding.LateSetComplex(obj, (Type)null, "name", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
							}
							object obj3 = Operators.CompareObjectEqual(obj2, (object)"code", false);
							Type typeFromHandle2 = typeof(Strings);
							object[] array4 = new object[1];
							obj = objectValue3;
							array4[0] = NewLateBinding.LateGet(obj, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null);
							array2 = array4;
							obj2 = NewLateBinding.LateGet((object)null, typeFromHandle2, "UCase", array4, (string[])null, (Type[])null, array3 = new bool[1] { true });
							if (array3[0])
							{
								NewLateBinding.LateSetComplex(obj, (Type)null, "Value", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
							}
							object obj4 = Operators.CompareObjectEqual(obj2, (object)"REF", false);
							Type typeFromHandle3 = typeof(Strings);
							object[] array5 = new object[1];
							obj = objectValue3;
							array5[0] = NewLateBinding.LateGet(obj, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null);
							array2 = array5;
							obj2 = NewLateBinding.LateGet((object)null, typeFromHandle3, "UCase", array5, (string[])null, (Type[])null, array3 = new bool[1] { true });
							if (array3[0])
							{
								NewLateBinding.LateSetComplex(obj, (Type)null, "Value", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
							}
							object obj5 = Operators.OrObject(obj4, Operators.CompareObjectEqual(obj2, (object)"D", false));
							Type typeFromHandle4 = typeof(Strings);
							object[] array6 = new object[1];
							obj = objectValue3;
							array6[0] = NewLateBinding.LateGet(obj, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null);
							array2 = array6;
							obj2 = NewLateBinding.LateGet((object)null, typeFromHandle4, "UCase", array6, (string[])null, (Type[])null, array3 = new bool[1] { true });
							if (array3[0])
							{
								NewLateBinding.LateSetComplex(obj, (Type)null, "Value", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
							}
							object obj6 = Operators.AndObject(obj3, Operators.OrObject(obj5, Operators.CompareObjectEqual(obj2, (object)"A", false)));
							Type typeFromHandle5 = typeof(Strings);
							object[] array7 = new object[1];
							obj = objectValue3;
							array7[0] = NewLateBinding.LateGet(obj, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null);
							array2 = array7;
							obj2 = NewLateBinding.LateGet((object)null, typeFromHandle5, "LCase", array7, (string[])null, (Type[])null, array3 = new bool[1] { true });
							if (array3[0])
							{
								NewLateBinding.LateSetComplex(obj, (Type)null, "name", new object[1] { array2[0] }, (string[])null, (Type[])null, true, false);
							}
							if (Conversions.ToBoolean(Operators.OrObject(obj6, Operators.AndObject(Operators.CompareObjectEqual(obj2, (object)"dps_number", false), (object)clsCollection5.get_Exists(Conversions.ToString(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)))))))
							{
								break;
							}
							clsCollection6.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
							if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue3, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"dps_number", false))
							{
								text4 = Conversions.ToString(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
							}
							else if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue3, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null), (object)"dps_rev", false))
							{
								text5 = Conversions.ToString(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null));
							}
						}
						int num = 0;
						if (Operators.CompareString(text4, "", false) != 0)
						{
							XmlNodeList xmlNodeList3 = item2.SelectNodes("signature");
							foreach (XmlElement item3 in xmlNodeList3)
							{
								num++;
								ClsCollection clsCollection7 = new ClsCollection();
								foreach (object attribute4 in item3.Attributes)
								{
									object objectValue4 = RuntimeHelpers.GetObjectValue(attribute4);
									clsCollection7.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
								}
								if (clsCollection7.Count > 0)
								{
									clsCollection6.Add("Sign_" + Conversions.ToString(num), clsCollection7);
								}
							}
							if (clsCollection6.Count > 0)
							{
								clsCollection5.Add(text4, clsCollection6);
							}
						}
						if (((Operators.CompareString(text4, "", false) != 0) & (Operators.CompareString(text5, "", false) != 0)) && !clsCollection3.get_Exists(text4 + text5))
						{
							clsCollection3.Add(text4 + text5, clsCollection6);
						}
					}
					clsCollection4.Add("dps", clsCollection5);
					clsCollection2.Add(Conversions.ToString(clsCollection2.Count + 1), clsCollection4);
				}
				if (!clsCollection.get_Exists("rv_listauthorities"))
				{
					clsCollection.Add("rv_listauthorities", text3);
				}
				clsCollection.Add("authority", clsCollection2);
				clsCollection.Add("dps", clsCollection3);
				result = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sRVnumber, sRVrev);
				}
			}
			return result;
		}
	}

	public ClsCollection GetPLMActionDocuments(string sPLMAction, ref object oDocumentCollection, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/attachedDocumentsByPLMAction/%PLMACTION%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%PLMACTION%", sPLMAction, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/document");
			ClsCollection clsCollection2 = new ClsCollection();
			foreach (XmlElement item in xmlNodeList)
			{
				string attribute = item.GetAttribute("FIELD_PART_NUMBER");
				string attribute2 = item.GetAttribute("FIELD_DOCUMENT_REVISION");
				clsCollection2.Add(attribute, attribute2);
			}
			oDocumentCollection = clsCollection2;
			result = clsCollection2;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
			}
		}
		return result;
	}

	public ClsCollection GetBDIFullData(string sdpsnumber, string sdpsrev, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/edrn/%BASENUMBER%/%REVISION%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%BASENUMBER%", sdpsnumber, 1, -1, (CompareMethod)0), "%REVISION%", sdpsrev, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		checked
		{
			ClsCollection result;
			if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
			{
				result = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.LoadXml(text2);
				XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/edrn");
				ClsCollection clsCollection2 = new ClsCollection();
				foreach (object attribute in xmlNode.Attributes)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(attribute);
					clsCollection2.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				clsCollection.Add("EDRN", clsCollection2);
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/relatedParts/part");
				ClsCollection clsCollection3 = new ClsCollection();
				foreach (XmlElement item in xmlNodeList)
				{
					ClsCollection clsCollection4 = new ClsCollection();
					foreach (object attribute2 in item.Attributes)
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(attribute2);
						clsCollection4.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue2, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection3.Add("RelatedPart_" + Conversions.ToString(clsCollection3.Count + 1), clsCollection4);
				}
				clsCollection.Add("RelatedParts", clsCollection3);
				xmlNode = xmlDocument.SelectSingleNode("/data/dps_attributes");
				ClsCollection clsCollection5 = new ClsCollection();
				foreach (object attribute3 in xmlNode.Attributes)
				{
					object objectValue3 = RuntimeHelpers.GetObjectValue(attribute3);
					clsCollection5.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				if (!clsCollection5.get_Exists("FIELD_PROGRAM"))
				{
					clsCollection5.Add("FIELD_PROGRAM", "Global 7000/8000 Program");
				}
				clsCollection.Add("dps_attributes", clsCollection5);
				xmlNode = xmlDocument.SelectSingleNode("/data/workflowStatus");
				ClsCollection clsCollection6 = new ClsCollection();
				foreach (object attribute4 in xmlNode.Attributes)
				{
					object objectValue4 = RuntimeHelpers.GetObjectValue(attribute4);
					clsCollection6.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
				}
				xmlNodeList = xmlDocument.SelectNodes("/data/workflowStatus/signatories/signatory");
				ClsCollection clsCollection7 = new ClsCollection();
				foreach (XmlElement item2 in xmlNodeList)
				{
					ClsCollection clsCollection8 = new ClsCollection();
					foreach (object attribute5 in item2.Attributes)
					{
						object objectValue5 = RuntimeHelpers.GetObjectValue(attribute5);
						clsCollection8.Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue5, (Type)null, "name", new object[0], (string[])null, (Type[])null, (bool[])null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue5, (Type)null, "Value", new object[0], (string[])null, (Type[])null, (bool[])null)));
					}
					clsCollection7.Add("signatory_" + Conversions.ToString(clsCollection7.Count + 1), clsCollection8);
				}
				clsCollection6.Add("signatories", clsCollection7);
				clsCollection.Add("workflowStatus", clsCollection6);
				result = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sdpsnumber, sdpsrev);
				}
			}
			return result;
		}
	}

	public string GetPartRefOID(string sPartNumber, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = "";
		string text2 = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/part/getPartRefByPartNumber/%PARTNUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0);
		string text3 = sGetWebResult(text2, "", GetUncachedResult);
		string result;
		if ((Operators.CompareString(text3, "", false) == 0) | (Operators.CompareString(text3, webServiceErrorMessage, false) == 0))
		{
			result = text;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text3);
			text = ((XmlElement)xmlDocument.SelectSingleNode("/data/Part")).GetAttribute("FIELD_OID");
			result = text;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2, sPartNumber);
			}
		}
		return result;
	}

	public string GetPrcOID(string sPRC, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		string text = "";
		string text2 = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/prc/getPrcByVid/%ID%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%ID%", sPRC, 1, -1, (CompareMethod)0);
		string text3 = sGetWebResult(text2, "", GetUncachedResult);
		string result;
		if ((Operators.CompareString(text3, "", false) == 0) | (Operators.CompareString(text3, webServiceErrorMessage, false) == 0))
		{
			result = text;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text3);
			text = ((XmlElement)xmlDocument.SelectSingleNode("/data/PRC")).GetAttribute("FIELD_OID");
			result = text;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2);
			}
		}
		return result;
	}

	public ClsCollection GetOIDinfoOfInstance(string sInstanceName, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/part/getPartInstanceByInstanceNumber/%INSTANCENUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%INSTANCENUMBER%", sInstanceName, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/Instance");
			foreach (XmlAttribute attribute in xmlNode.Attributes)
			{
				clsCollection.Add(attribute.Name, attribute.Value);
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sInstanceName);
			}
		}
		return result;
	}

	public ClsCollection GetNHAPartNumber(string sPartNumber, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/enovia/NHA/%PARTNUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/partNHA");
			foreach (XmlElement item in xmlNodeList)
			{
				ClsCollection clsCollection2 = new ClsCollection();
				foreach (XmlAttribute attribute in item.Attributes)
				{
					clsCollection2.Add(attribute.Name, attribute.Value);
				}
				clsCollection.Add(checked(clsCollection.Count + 1), clsCollection2);
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber);
			}
		}
		return result;
	}

	public ClsCollection GetAvailableInstances(string sPartNumber, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/nha/%PARTNUMBER%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/part");
			foreach (XmlElement item in xmlNodeList)
			{
				ClsCollection clsCollection2 = new ClsCollection();
				foreach (XmlAttribute attribute in item.Attributes)
				{
					clsCollection2.Add(attribute.Name, attribute.Value);
				}
				clsCollection.Add(checked(clsCollection.Count + 1), clsCollection2);
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, sPartNumber);
			}
		}
		return result;
	}

	public ClsCollection GetUserInfo(string sUserId, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/api/reports/echecker/userinfo/%USERID%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%USERID%", sUserId, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		ClsCollection result;
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			result = clsCollection;
		}
		else
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text2);
			XmlNode xmlNode = xmlDocument.SelectSingleNode("/data/UserInfo");
			foreach (XmlAttribute attribute in xmlNode.Attributes)
			{
				clsCollection.Add(attribute.Name, attribute.Value);
			}
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/UserInfo/Group");
			int num = 1;
			clsCollection.Add("Groups", new ClsCollection());
			ClsCollection clsCollection2 = (ClsCollection)clsCollection.get_GetItem((object)"Groups");
			foreach (XmlElement item in xmlNodeList)
			{
				clsCollection2.Add("Group" + num, new ClsCollection());
				foreach (XmlAttribute attribute2 in item.Attributes)
				{
					object[] array;
					XmlAttribute xmlAttribute2;
					bool[] array2;
					NewLateBinding.LateCall(clsCollection2.get_GetItem((object)("Group" + num)), (Type)null, "Add", array = new object[2]
					{
						attribute2.Name,
						(xmlAttribute2 = attribute2).Value
					}, (string[])null, (Type[])null, array2 = new bool[2] { false, true }, true);
					if (array2[1])
					{
						xmlAttribute2.Value = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(string));
					}
				}
				num = checked(num + 1);
			}
			result = clsCollection;
			if (EnableToolUsageLog)
			{
				LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text);
			}
		}
		return result;
	}

	public ClsCollection GetEnoviaLoginContext(string sLoginID, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection clsCollection = new ClsCollection();
		string text = Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/enovialogincontexts/?enoviauserid=%LOGINID%&enoviadatabase=%ENOMDBNAME%", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%LOGINID%", sLoginID, 1, -1, (CompareMethod)0);
		text = Strings.Replace(text, "%ENOMDBNAME%", ENOMDBNAME, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			return clsCollection;
		}
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(text2);
		XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/loginContext");
		foreach (XmlElement item in xmlNodeList)
		{
			ClsCollection clsCollection2 = new ClsCollection();
			foreach (XmlAttribute attribute in item.Attributes)
			{
				clsCollection2.Add(attribute.Name, attribute.Value);
			}
			clsCollection.Add("LoginContext." + Conversions.ToString(checked(clsCollection.Count + 1)), clsCollection2);
		}
		return clsCollection;
	}

	public bool GetStdPartStatusBasedOnFTVCDLAndAttributes(string sPartNumber, ref string sOutputString, [Optional][DefaultParameterValue(null)] ref ClsCollection oResultCollection, bool GetUncachedResult = false)
	{
		ClsCollection clsCollection = new ClsCollection();
		string text = sPartNumber;
		int num = Information.UBound((Array)Strings.Split(MySettingsProperty.Settings.sFTVCDLForbiddenChars, "|", -1, (CompareMethod)0), 1);
		checked
		{
			for (int i = 0; i <= num; i++)
			{
				if (Strings.InStr(1, text, Strings.Split(MySettingsProperty.Settings.sFTVCDLForbiddenChars, "|", -1, (CompareMethod)0)[i], (CompareMethod)0) != 0)
				{
					text = Strings.Left(text, Strings.InStr(1, text, Strings.Split(MySettingsProperty.Settings.sFTVCDLForbiddenChars, "|", -1, (CompareMethod)0)[i], (CompareMethod)0) - 1);
				}
			}
			string text2 = Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/ftvcdl/parts/%PARTNUMBER%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%PARTNUMBER%", text, 1, -1, (CompareMethod)0);
			string text3 = sGetWebResult(text2, "", GetUncachedResult);
			bool result;
			if ((Operators.CompareString(text3, "", false) == 0) | (Operators.CompareString(text3, webServiceErrorMessage, false) == 0))
			{
				result = false;
				oResultCollection = clsCollection;
			}
			else
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.LoadXml(text3);
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/data/part");
				bool flag = false;
				foreach (XmlElement item in xmlNodeList)
				{
					clsCollection.Add("Part" + Conversions.ToString(clsCollection.Count + 1), new ClsCollection());
					foreach (XmlAttribute attribute in item.Attributes)
					{
						object[] array;
						XmlAttribute xmlAttribute2;
						bool[] array2;
						NewLateBinding.LateCall(clsCollection.get_GetItem((object)("Part" + Conversions.ToString(clsCollection.Count))), (Type)null, "add", array = new object[2]
						{
							attribute.Name,
							(xmlAttribute2 = attribute).Value
						}, (string[])null, (Type[])null, array2 = new bool[2] { false, true }, true);
						if (array2[1])
						{
							xmlAttribute2.Value = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(string));
						}
					}
				}
				XmlNodeList xmlNodeList2 = xmlDocument.SelectNodes("/data/part[@replaceBy and @number=\"" + sPartNumber + "\"]");
				if (xmlNodeList2.Count <= 0)
				{
					foreach (XmlElement item2 in xmlNodeList)
					{
						if (Operators.CompareString(item2.GetAttribute("number"), sPartNumber, false) != 0)
						{
							continue;
						}
						string text4 = Strings.UCase(item2.GetAttribute("acsn"));
						if (Operators.CompareString(text4, "HW GLOSSARY", false) != 0)
						{
							if (Operators.CompareString(text4, "NON-LOGISTICS HW", false) == 0)
							{
								flag = true;
								sOutputString = "NON-LOGISTICS HW|OK";
							}
						}
						else if (((Operators.CompareString(Strings.Trim(item2.GetAttribute("logisticsComment")), "", false) == 0) | (Operators.CompareString(Strings.UCase(Strings.Replace(item2.GetAttribute("logisticsComment"), " ", "", 1, -1, (CompareMethod)0)), "FTV4BASELINE", false) == 0)) & (Operators.CompareString(item2.GetAttribute("logisticsCheck"), "1", false) == 0))
						{
							string text5 = Strings.UCase(Strings.Replace(item2.GetAttribute("partStatus"), " ", "", 1, -1, (CompareMethod)0));
							if (Operators.CompareString(text5, "FLAGGEDFORDELETE", false) != 0)
							{
								sOutputString = "HW GLOSSARY|OK";
								flag = true;
								break;
							}
							sOutputString = "HW GLOSSARY|DELETED";
							flag = false;
						}
						else
						{
							flag = false;
							sOutputString = "HW GLOSSARY|Found but logisticsComment:" + item2.GetAttribute("logisticsComment") + "|POSent:" + (Operators.CompareString(item2.GetAttribute("logisticsCheck"), "1", false) == 0);
						}
					}
				}
				else if (xmlDocument.SelectNodes("/data/part[@replaceBy!='' and @number=\"" + sPartNumber + "\"]").Count <= 0)
				{
					foreach (XmlElement item3 in xmlNodeList2)
					{
						string text6 = Strings.UCase(Strings.Replace(item3.GetAttribute("partStatus"), " ", "", 1, -1, (CompareMethod)0));
						if (Operators.CompareString(text6, "FLAGGEDFORDELETE", false) == 0)
						{
							sOutputString = "HW GLOSSARY|DELETED";
							flag = false;
							continue;
						}
						sOutputString = "HW GLOSSARY|OK";
						flag = true;
						break;
					}
				}
				else
				{
					sOutputString = "HW ALTERNATE|" + xmlDocument.SelectNodes("/data/part[@replaceBy!='' and @number=\"" + sPartNumber + "\"]").Item(1).Attributes.Item(Conversions.ToInteger("replaceBy")).Value;
					flag = false;
				}
				result = flag;
				oResultCollection = clsCollection;
				if (EnableToolUsageLog)
				{
					LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2, sPartNumber);
				}
			}
			return result;
		}
	}

	public ClsCollection GetDocumentsLinkedToEnoviaPart(string PartNumber, string Rev, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		ClsCollection result = new ClsCollection();
		string text = Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/getDocumentLinksToParts/%PartNumber%/%Rev%", "%SERVERNAME%", sServerName(), 1, -1, (CompareMethod)0);
		text = Strings.Replace(Strings.Replace(text, "%PartNumber%", PartNumber, 1, -1, (CompareMethod)0), "%Rev%", Rev, 1, -1, (CompareMethod)0);
		string text2 = sGetWebResult(text, "", GetUncachedResult);
		if ((Operators.CompareString(text2, "", false) == 0) | (Operators.CompareString(text2, webServiceErrorMessage, false) == 0))
		{
			return result;
		}
		return PopulateClsColfromXMLdata(text2);
	}

	public string GetEDRNBox21(string sPartNumber, string sRev, bool bHistory, bool GetUncachedResult = false)
	{
		_GeneralException = "";
		new ClsCollection();
		string text = Conversions.ToString(Interaction.IIf(bHistory, (object)"TRUE", (object)"FALSE"));
		string text2 = Strings.Replace(Strings.Replace(Strings.Replace(Strings.Replace("%SERVERNAME%/api/echecker/reports/enovia/fillEDRNBox21/?partNumber=%PARTNUMBER%&rev=%REVISION%&history=%HISTORY%", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0), "%PARTNUMBER%", sPartNumber, 1, -1, (CompareMethod)0), "%REVISION%", sRev, 1, -1, (CompareMethod)0), "%HISTORY%", text, 1, -1, (CompareMethod)0);
		string result = sGetWebResult(text2, "", GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text2, sPartNumber, sRev);
		}
		return result;
	}

	public string GetUserPassword(string sUserName, [Optional][DefaultParameterValue(false)] ref bool bexitbyuser, bool bConnectionError = false)
	{
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Unknown result type (might be due to invalid IL or missing references)
		_GeneralException = "";
		if (_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd == null)
		{
			_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd = new Hashtable();
		}
		string result = "";
		if (Operators.CompareString(sUserName, "", false) == 0)
		{
			sUserName = Environment.UserName;
		}
		if (!_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd.Contains(sUserName))
		{
			UsrLoginForm usrLoginForm = new UsrLoginForm(sUserName, FrmLeft, FrmTop);
			while (true)
			{
				usrLoginForm.Password = "";
				usrLoginForm.PasswordTextBox.Text = "";
				if (ParentWindowHWND == 0)
				{
					((Form)usrLoginForm).ShowDialog();
				}
				else
				{
					((Form)usrLoginForm).ShowDialog((IWin32Window)(object)NativeWindow.FromHandle((IntPtr)ParentWindowHWND));
				}
				if (usrLoginForm.QuitByUser)
				{
					bexitbyuser = true;
					break;
				}
				string strURL = Strings.Replace("%SERVERNAME%/api/users/me", "%SERVERNAME%", sServerName(UseDevtDBForGet), 1, -1, (CompareMethod)0);
				string text = "";
				text = sGetWebResult(strURL, Base64Encode(sUserName + ":" + usrLoginForm.Password), GetUncachedResult: true);
				if (Operators.CompareString(text, webServiceErrorMessage, false) == 0)
				{
					bConnectionError = true;
					break;
				}
				if (Strings.InStr(text, "\"badge\":", (CompareMethod)0) > 0)
				{
					_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd.Add(sUserName, Base64Encode(sUserName + ":" + usrLoginForm.Password));
					result = Conversions.ToString(_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd[sUserName]);
					break;
				}
				if (Operators.CompareString(text, "", false) == 0)
				{
					usrLoginForm.PasswordLabel.Text = "Invalid Password!";
				}
			}
		}
		else
		{
			result = Conversions.ToString(_0024STATIC_0024GetUserPassword_0024203EE1022_0024CachedPwd[sUserName]);
		}
		return result;
	}

	public ClsCollection PopulateClsColfromXMLdata(string sXMLdata)
	{
		XmlDocument xmlDocument = new XmlDocument();
		ClsCollection oCol = new ClsCollection();
		try
		{
			xmlDocument.LoadXml(sXMLdata);
			_PopulateClsColfromXMLdata(ref oCol, xmlDocument.SelectSingleNode("/*").ChildNodes);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return oCol;
	}

	private void _PopulateClsColfromXMLdata(ref ClsCollection oCol, XmlNodeList oNodeList)
	{
		int num = 1;
		checked
		{
			foreach (XmlNode oNode in oNodeList)
			{
				string name = oNode.Name;
				int num2 = 0;
				foreach (XmlNode oNode2 in oNodeList)
				{
					if (Operators.CompareString(oNode2.Name, name, false) == 0)
					{
						num2++;
					}
				}
				if (num2 > 1)
				{
					name = oNode.Name + "_" + num;
					num++;
				}
				else
				{
					name = oNode.Name;
				}
				oCol.Add(name, new ClsCollection());
				foreach (XmlAttribute attribute in oNode.Attributes)
				{
					object[] array;
					XmlAttribute xmlAttribute2;
					bool[] array2;
					NewLateBinding.LateCall(oCol.get_GetItem((object)name), (Type)null, "Add", array = new object[2]
					{
						attribute.Name,
						(xmlAttribute2 = attribute).Value
					}, (string[])null, (Type[])null, array2 = new bool[2] { false, true }, true);
					if (array2[1])
					{
						xmlAttribute2.Value = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(string));
					}
				}
				ClsCollection oCol2 = (ClsCollection)oCol.get_GetItem((object)name);
				_PopulateClsColfromXMLdata(ref oCol2, oNode.ChildNodes);
			}
		}
	}

	private string sServerName(bool UseDevtDB = false)
	{
		GetCurrentDNSDomainName();
		string text = "space.aero.bombardier.net";
		text = ((!((Operators.CompareString(GetCurrentDNSDomainName(), MySettingsProperty.Settings.sOnsiteDomain, false) == 0) & !GetMachineOS().ToUpper().Contains(" XP "))) ? "space.aero.bombardier.net" : "ca.aero.bombardier.net");
		string text2 = (UseDevtDB ? (("bdi-dev." + text) ?? "") : ("bdi." + text));
		return "https://" + text2;
	}

	private string Base64Encode(object sText)
	{
		Encoding aSCII = Encoding.ASCII;
		object[] obj = new object[1] { sText };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object obj3 = NewLateBinding.LateGet((object)aSCII, (Type)null, "GetBytes", obj, (string[])null, (Type[])null, obj2);
		if (array2[0])
		{
			sText = RuntimeHelpers.GetObjectValue(array[0]);
		}
		return Convert.ToBase64String((byte[])obj3);
	}

	public static bool ValidateRemoteCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
	{
		return true;
	}

	public string GetWebResult(string strURL, string EncryptedUserCredentials = "", bool GetUncachedResult = false)
	{
		string result = sGetWebResult(strURL, EncryptedUserCredentials, GetUncachedResult);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, strURL);
		}
		return result;
	}

	private string sGetWebResult(string strURL, string EncryptedUserCredentials = "", bool GetUncachedResult = false)
	{
		_GeneralException = "";
		_ResultFromCache = false;
		_WebClientReturnStatus = "";
		_WebResponseExceptionStatus = "";
		double timer = DateAndTime.Timer;
		string text = "";
		if (!CheckAuthorization())
		{
			return text;
		}
		if (!GetUncachedResult && oCache.Contains(strURL))
		{
			text = Conversions.ToString(oCache[strURL]);
			_ResultFromCache = true;
			return text;
		}
		ServicePointManager.ServerCertificateValidationCallback = ValidateRemoteCertificate;
		int num = 0;
		checked
		{
			do
			{
				Stream stream = null;
				EnhancedWebClient enhancedWebClient = null;
				StreamReader streamReader = null;
				FileStream fileStream = null;
				try
				{
					enhancedWebClient = new EnhancedWebClient();
					enhancedWebClient.TimeOutInSeconds = TimeOutInSeconds;
					enhancedWebClient.Headers.Add(HttpRequestHeader.UserAgent, sHeaderID);
					if (Operators.CompareString(EncryptedUserCredentials, "", false) != 0)
					{
						enhancedWebClient.Headers.Add(HttpRequestHeader.Authorization, "Basic " + EncryptedUserCredentials);
					}
					stream = enhancedWebClient.OpenRead(strURL);
					WebHeaderCollection responseHeaders = enhancedWebClient.ResponseHeaders;
					string text2 = "";
					if (responseHeaders.Count > 0)
					{
						int num2 = responseHeaders.Count - 1;
						for (int i = 0; i <= num2; i++)
						{
							if (Operators.CompareString(responseHeaders.GetKey(i), "Content-disposition", false) == 0 && LikeOperator.LikeString(responseHeaders.Get(i), "attachment;*", (CompareMethod)0))
							{
								text2 = Strings.Replace(Strings.Split(responseHeaders.Get(i), "filename=", -1, (CompareMethod)0)[Information.UBound((Array)Strings.Split(responseHeaders.Get(i), "filename=", -1, (CompareMethod)0), 1)], "\"", "", 1, -1, (CompareMethod)0);
							}
						}
					}
					if (Operators.CompareString(text2, "", false) != 0)
					{
						fileStream = new FileStream(Path.Combine(DefaultDownloadDirectoryPath, text2), FileMode.Create);
						stream.CopyTo(fileStream);
						text = text2;
					}
					else if (Operators.CompareString(text2, "", false) == 0)
					{
						streamReader = new StreamReader(stream);
						text = streamReader.ReadToEnd();
						if (!oCache.Contains(strURL))
						{
							oCache.Add(strURL, text);
						}
						else
						{
							oCache[strURL] = text;
						}
					}
					_WebClientReturnStatus = "";
					_WebResponseExceptionStatus = "";
					_GeneralException = "";
				}
				catch (WebException ex)
				{
					ProjectData.SetProjectError((Exception)ex);
					WebException ex2 = ex;
					_WebClientReturnStatus = ex2.Status.ToString() + "\r\n" + ex2.StackTrace;
					if (ex2.Status == WebExceptionStatus.ProtocolError)
					{
						HttpWebResponse httpWebResponse = (HttpWebResponse)ex2.Response;
						_WebResponseExceptionStatus = httpWebResponse.StatusDescription + "\r\n" + ex2.StackTrace;
						if (httpWebResponse.StatusCode == HttpStatusCode.Unauthorized)
						{
							text = "";
							ProjectData.ClearProjectError();
							break;
						}
					}
					ProjectData.ClearProjectError();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					_GeneralException = ex4.Message + "\r\n" + ex4.StackTrace;
					ProjectData.ClearProjectError();
				}
				finally
				{
					streamReader?.Close();
					stream?.Close();
					fileStream?.Close();
					fileStream = null;
					streamReader = null;
					enhancedWebClient = null;
					ServicePointManager.ServerCertificateValidationCallback = null;
				}
				num++;
				if ((num == MaximumReadConnectionAttempts) & ((Operators.CompareString(_WebResponseExceptionStatus, "", false) != 0) | (Operators.CompareString(_WebClientReturnStatus, "", false) != 0) | (Operators.CompareString(_GeneralException, "", false) != 0)))
				{
					text = webServiceErrorMessage;
				}
			}
			while (!((num >= MaximumReadConnectionAttempts) | (Operators.CompareString(text, "", false) != 0)));
			_WebtransactionTime = Conversions.ToString(Math.Round(DateAndTime.Timer - timer, 3, MidpointRounding.AwayFromZero));
			return text;
		}
	}

	public void SendXMLToServer(string sDataToUpload, string strURL, string HTTPMethod = "POST", string sEncodedUsrAndPwd = "")
	{
		_SendXMLToServer(sDataToUpload, strURL, HTTPMethod, sEncodedUsrAndPwd);
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, strURL, "", "", null, sDataToUpload);
		}
	}

	private void _SendXMLToServer(string sDataToUpload, string strURL, string HTTPMethod = "POST", string sEncodedUsrAndPwd = "")
	{
		_GeneralException = "";
		_WebtransactionTime = "";
		_ResultFromCache = false;
		if (!CheckAuthorization())
		{
			return;
		}
		double timer = DateAndTime.Timer;
		EnhancedWebClient enhancedWebClient = null;
		bool flag = true;
		int num = 0;
		do
		{
			try
			{
				ServicePointManager.ServerCertificateValidationCallback = ValidateRemoteCertificate;
				num = checked(num + 1);
				enhancedWebClient = new EnhancedWebClient();
				enhancedWebClient.TimeOutInSeconds = TimeOutInSeconds;
				enhancedWebClient.Headers.Add(HttpRequestHeader.UserAgent, sHeaderID);
				enhancedWebClient.Headers.Add(HttpRequestHeader.ContentType, "application/xml");
				if (Operators.CompareString(sEncodedUsrAndPwd, "", false) != 0)
				{
					enhancedWebClient.Headers.Add(HttpRequestHeader.Authorization, "Basic " + sEncodedUsrAndPwd);
				}
				Encoding.UTF8.GetString(enhancedWebClient.UploadData(strURL, HTTPMethod, Encoding.UTF8.GetBytes(sDataToUpload)));
				_WebClientReturnStatus = "";
				_WebResponseExceptionStatus = "";
				_GeneralException = "";
				flag = false;
			}
			catch (WebException ex)
			{
				ProjectData.SetProjectError((Exception)ex);
				WebException ex2 = ex;
				HttpWebResponse httpWebResponse = (HttpWebResponse)ex2.Response;
				_WebClientReturnStatus = ex2.Status.ToString() + "\r\n" + ex2.StackTrace.Trim();
				_WebResponseExceptionStatus = httpWebResponse.StatusDescription + "\r\n" + ex2.StackTrace;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				_GeneralException = ex4.Message + "\r\n" + ex4.StackTrace.Trim();
				ProjectData.ClearProjectError();
			}
			finally
			{
				enhancedWebClient = null;
				ServicePointManager.ServerCertificateValidationCallback = null;
			}
		}
		while (!(MaximumWriteConnectionAttempts >= num || !flag));
		_WebtransactionTime = Conversions.ToString(Math.Round(timer - DateAndTime.Timer, 3, MidpointRounding.AwayFromZero));
	}

	public string AddAttachementToDPS(string sFileNameWithPath, string sBDIDocID, string sEncodedUsrAndPwd = "")
	{
		_GeneralException = "";
		string text = Strings.Replace(Strings.Replace("%SERVERNAME%/bdicommon/attachments/210/%DocumentID%/", "%SERVERNAME%", sServerName(UseDevtDB), 1, -1, (CompareMethod)0), "%DocumentID%", sBDIDocID, 1, -1, (CompareMethod)0);
		string text2 = SendFileToServer(sFileNameWithPath, text, "POST", sEncodedUsrAndPwd);
		string result = text2;
		if (EnableToolUsageLog)
		{
			LogWebserviceUsage(MethodBase.GetCurrentMethod().Name, text, "", "", null, text2);
		}
		return result;
	}

	public string SendFileToServer(string sFileNameWithPath, string strURL, string HTTPMethod = "POST", string sEncodedUsrAndPwd = "")
	{
		_GeneralException = "";
		_GeneralException = "";
		_WebtransactionTime = "";
		_ResultFromCache = false;
		string result = "";
		if (!CheckAuthorization())
		{
			return "";
		}
		double timer = DateAndTime.Timer;
		bool flag = true;
		int num = 0;
		string text = "File Name";
		do
		{
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(strURL);
			MemoryStream memoryStream = new MemoryStream();
			FileStream fileStream = new FileStream(sFileNameWithPath, FileMode.Open);
			checked
			{
				try
				{
					ServicePointManager.ServerCertificateValidationCallback = ValidateRemoteCertificate;
					num++;
					if (Operators.CompareString(sEncodedUsrAndPwd, "", false) != 0)
					{
						httpWebRequest.Headers.Add(HttpRequestHeader.Authorization, "Basic " + sEncodedUsrAndPwd);
					}
					httpWebRequest.UserAgent = sHeaderID;
					httpWebRequest.Timeout = TimeOutInSeconds * 1000;
					httpWebRequest.Method = HTTPMethod;
					string text2 = Guid.NewGuid().ToString().Replace("-", "");
					httpWebRequest.ContentType = "multipart/form-data; boundary=" + text2;
					string text3 = "\r\n";
					string s = "--" + text2 + text3 + "Content-Disposition: form-data; name=\"" + text + "\"; filename=\"" + Path.GetFileName(sFileNameWithPath) + "\"" + text3 + "Content-Type: \tapplication/octet-stream" + text3 + text3;
					memoryStream.Write(Encoding.UTF8.GetBytes(s), 0, Encoding.UTF8.GetBytes(s).Length);
					byte[] array = new byte[4097];
					for (int num2 = fileStream.Read(array, 0, array.Length); num2 > 0; num2 = fileStream.Read(array, 0, array.Length))
					{
						memoryStream.Write(array, 0, num2);
					}
					string s2 = text3 + "--" + text2 + "--" + text3;
					memoryStream.Write(Encoding.UTF8.GetBytes(s2), 0, Encoding.UTF8.GetBytes(s2).Length);
					httpWebRequest.ContentLength = memoryStream.Length;
					Stream requestStream = httpWebRequest.GetRequestStream();
					memoryStream.WriteTo(requestStream);
					((IDisposable)requestStream).Dispose();
					requestStream.Close();
					result = new StreamReader(httpWebRequest.GetResponse().GetResponseStream()).ReadToEnd();
					_WebClientReturnStatus = "";
					_WebResponseExceptionStatus = "";
					_GeneralException = "";
					flag = false;
				}
				catch (WebException ex)
				{
					ProjectData.SetProjectError((Exception)ex);
					WebException ex2 = ex;
					HttpWebResponse httpWebResponse = (HttpWebResponse)ex2.Response;
					_WebClientReturnStatus = ex2.Status.ToString() + "\r\n" + ex2.StackTrace.Trim();
					_WebResponseExceptionStatus = httpWebResponse.StatusDescription + "\r\n" + ex2.StackTrace;
					ProjectData.ClearProjectError();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					_GeneralException = ex4.Message + "\r\n" + ex4.StackTrace.Trim();
					ProjectData.ClearProjectError();
				}
				finally
				{
					fileStream.Close();
					memoryStream.Close();
					ServicePointManager.ServerCertificateValidationCallback = null;
				}
			}
		}
		while (!(MaximumWriteConnectionAttempts >= num || !flag));
		return result;
	}

	private bool CheckAuthorization()
	{
		bool result = false;
		if (File.Exists("I:\\V5_KBE_Tools\\Production\\05_KBE_CATScript\\03_ENOVIA_Connection_Tools\\2-PROD\\AuthorizedApps.xml"))
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load("I:\\V5_KBE_Tools\\Production\\05_KBE_CATScript\\03_ENOVIA_Connection_Tools\\2-PROD\\AuthorizedApps.xml");
			XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/AuthorizedApps/App");
			foreach (XmlElement item in xmlNodeList)
			{
				try
				{
					if (Operators.CompareString(item.GetAttribute("AppName"), UsedbyApplication, false) == 0)
					{
						result = true;
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
		return result;
	}

	public void Close()
	{
		Finalize();
		BckGdWrk = new BackgroundWorker();
		BckGdWrk.RunWorkerAsync(ProcessID);
	}

	private void BckGdWrk_DoWork(object sender, DoWorkEventArgs e)
	{
		int num = Conversions.ToInteger(e.Argument);
		NewLateBinding.LateCall(RuntimeHelpers.GetObjectValue(Interaction.CreateObject("WScript.Shell", "")), (Type)null, "Run", new object[3]
		{
			"taskkill /f /pid " + Conversions.ToString(num),
			0,
			true
		}, (string[])null, (Type[])null, (bool[])null, true);
	}

	~WebServiceAccessTool()
	{
		LogPendingWebserviceUsageInCache();
		base.Finalize();
	}
}
