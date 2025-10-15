using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace CivilesPro.RevitTools.App
{
  [Transaction(TransactionMode.Manual)]
  public class CivilesProApp : IExternalApplication
  {
    public static UIControlledApplication? UIApp { get; private set; }

    public Result OnStartup(UIControlledApplication application)
    {
      UIApp = application;
      RibbonManager.Build(application);
      return Result.Succeeded;
    }

    public Result OnShutdown(UIControlledApplication application)
    {
      return Result.Succeeded;
    }
  }
}
