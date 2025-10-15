using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;

namespace CivilesPro.RevitTools.Commands
{
  [Transaction(TransactionMode.Manual)]
  public class ImportarNiveles : IExternalCommand
  {
    public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
    {
      var uiapp = data.Application;
      var doc = uiapp.ActiveUIDocument.Document;

      // Demo: solo muestra diálogo y crea 1 nivel de prueba si no existe
      TaskDialog.Show("CivilesPro", "Módulo Importar Niveles funcionando correctamente.");

      using (var tr = new Transaction(doc, "CivilesPro - Crear Nivel demo"))
      {
        tr.Start();
        // Nivel de ejemplo a 0.00 (si no existe)
        var level = CommandUtils.FindLevelByName(doc, "NIVEL_00");
        if (level == null)
        {
          level = Level.Create(doc, 0.0);
          level.Name = "NIVEL_00";
        }
        tr.Commit();
      }

      return Result.Succeeded;
    }
  }
}
