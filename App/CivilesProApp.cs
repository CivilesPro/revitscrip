using Autodesk.Revit.UI;

namespace CivilesPro.RevitTools.App
{
  public class CivilesProApp : IExternalApplication
  {
    public static UIControlledApplication? UIApp { get; private set; }

    public Result OnStartup(UIControlledApplication application)
    {
      UIApp = application;
      RibbonManager.Create(application);
      return Result.Succeeded;
    }

    public Result OnShutdown(UIControlledApplication application)
    {
      return Result.Succeeded;
    }
  }
}
