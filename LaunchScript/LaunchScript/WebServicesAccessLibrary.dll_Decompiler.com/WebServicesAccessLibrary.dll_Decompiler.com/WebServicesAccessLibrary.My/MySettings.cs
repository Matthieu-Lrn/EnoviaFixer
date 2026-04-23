using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

namespace WebServicesAccessLibrary.My;

[CompilerGenerated]
[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.8.0.0")]
[EditorBrowsable(EditorBrowsableState.Advanced)]
internal sealed class MySettings : ApplicationSettingsBase
{
	private static MySettings defaultInstance = (MySettings)(object)SettingsBase.Synchronized((SettingsBase)(object)new MySettings());

	public static MySettings Default => defaultInstance;

	[ApplicationScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("WebServiceAccessTool_TEST")]
	public string sHeaderID => Conversions.ToString(((ApplicationSettingsBase)this)["sHeaderID"]);

	[ApplicationScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("AERO.AERO.BOMBARDIER.NET")]
	public string sOnsiteDomain => Conversions.ToString(((ApplicationSettingsBase)this)["sOnsiteDomain"]);

	[ApplicationScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("I:\\cecc\\env\\Location.Info.txt")]
	public string sLocInfoFile => Conversions.ToString(((ApplicationSettingsBase)this)["sLocInfoFile"]);

	[ApplicationScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("/|\\|#")]
	public string sFTVCDLForbiddenChars => Conversions.ToString(((ApplicationSettingsBase)this)["sFTVCDLForbiddenChars"]);
}
