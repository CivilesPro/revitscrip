using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;

namespace CivilesPro.RevitTools.App
{
  internal static class RibbonManager
  {
    private const string TabName = "CivilesPro";
    private const string PanelName = "Automatizaciones";

    public static void Create(UIControlledApplication app)
    {
      try
      {
        try { app.CreateRibbonTab(TabName); } catch { /* ya existe */ }
        var panel = EnsurePanel(app, TabName, PanelName);

        // Assembly + path del addin
        string asmPath = Assembly.GetExecutingAssembly().Location;
        var pd = new PushButtonData(
          "btnImportarNiveles",
          "Importar\nNiveles",
          asmPath,
          "CivilesPro.RevitTools.Commands.ImportarNivelesCommand");

        var btn = panel.AddItem(pd) as PushButton;
        if (btn != null)
        {
          btn.ToolTip = "Importa niveles desde un CSV y crea vistas.";
          btn.LongDescription = "Demostración inicial de CivilesPro.";
        }
      }
      catch (Exception ex)
      {
        Core.Logger.Error("Error creando Ribbon", ex);
        TaskDialog.Show("CivilesPro", "No se pudo crear el Ribbon.\n" + ex.Message);
      }
    }

    private static RibbonPanel EnsurePanel(UIControlledApplication app, string tab, string name)
    {
      foreach (var p in app.GetRibbonPanels(tab))
        if (p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
          return p;

      return app.CreateRibbonPanel(tab, name);
    }
  }
}
