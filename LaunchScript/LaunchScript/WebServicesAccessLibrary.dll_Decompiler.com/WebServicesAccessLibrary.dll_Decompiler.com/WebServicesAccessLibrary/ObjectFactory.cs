using System.Runtime.InteropServices;

namespace WebServicesAccessLibrary;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class ObjectFactory
{
	public ClsCollection GetClsCollection()
	{
		return new ClsCollection();
	}

	public WebServiceAccessTool GetWebServiceAccessTool()
	{
		return new WebServiceAccessTool();
	}
}
