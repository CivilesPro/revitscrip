using Autodesk.Revit.DB;
using System.Linq;

namespace CivilesPro.RevitTools.Commands
{
  internal static class CommandUtils
  {
    public static Level? FindLevelByName(Document doc, string name)
    {
      return new FilteredElementCollector(doc)
        .OfClass(typeof(Level))
        .Cast<Level>()
        .FirstOrDefault(l => l.Name == name);
    }
  }
}
