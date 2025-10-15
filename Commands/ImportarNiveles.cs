using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CivilesPro.RevitTools.Core;
using CivilesPro.RevitTools.FileReaders;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace CivilesPro.RevitTools.Commands
{
  [Transaction(TransactionMode.Manual)]
  public class ImportarNiveles : IExternalCommand
  {
    private const double DuplicateElevationToleranceMillimeters = 1.0; // tolerancia 1 mm

    public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
    {
      var uiDoc = data.Application.ActiveUIDocument;
      var doc = uiDoc?.Document;
      if (doc == null)
      {
        message = "No hay un documento activo.";
        return Result.Failed;
      }

      var filePath = PromptForCsvFile();
      if (filePath is null)
      {
        return Result.Cancelled;
      }

      var unitTypeId = PromptForUnits();
      if (unitTypeId is null)
      {
        return Result.Cancelled;
      }

      IReadOnlyList<RawLevelRow> rows;
      try
      {
        rows = CsvLevelReader.ReadRows(filePath);
      }
      catch (Exception ex)
      {
        Logger.Error("No se pudo leer el archivo de niveles.", ex);
        message = ex.Message;
        return Result.Failed;
      }

      var tolerance = UnitUtils.ConvertToInternalUnits(DuplicateElevationToleranceMillimeters, UnitTypeId.Millimeters);
      var preparedRows = new List<PreparedRow>();
      var warnings = new List<string>();
      var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

      foreach (var row in rows)
      {
        if (string.IsNullOrWhiteSpace(row.Name))
        {
          warnings.Add($"Fila {row.RowNumber}: el nombre está vacío.");
          continue;
        }

        if (!seenNames.Add(row.Name))
        {
          warnings.Add($"Fila {row.RowNumber}: el nombre '{row.Name}' está duplicado en el archivo.");
          continue;
        }

        if (!TryParseDouble(row.ElevationText, out var sourceElevation))
        {
          warnings.Add($"Fila {row.RowNumber}: elevación inválida '{row.ElevationText}'.");
          continue;
        }

        double elevationInternal;
        try
        {
          elevationInternal = UnitUtils.ConvertToInternalUnits(sourceElevation, unitTypeId.Value);
        }
        catch (Exception ex)
        {
          warnings.Add($"Fila {row.RowNumber}: no se pudo convertir la elevación '{row.ElevationText}'. {ex.Message}");
          continue;
        }

        preparedRows.Add(new PreparedRow(row.RowNumber, row.Name, elevationInternal));
      }

      if (preparedRows.Count == 0)
      {
        ShowSummary("No se encontraron filas válidas en el archivo.", warnings, Array.Empty<string>());
        return Result.Cancelled;
      }

      preparedRows.Sort((a, b) => a.ElevationInternal.CompareTo(b.ElevationInternal));

      var existingLevels = new FilteredElementCollector(doc)
        .OfClass(typeof(Level))
        .Cast<Level>()
        .ToDictionary(l => l.Name, StringComparer.OrdinalIgnoreCase);

      var createdLevels = new List<Level>();
      var updatedLevels = new List<Level>();
      var unchangedLevels = new List<Level>();

      var duplicateElevationCheck = existingLevels.Values
        .Select(l => (Level: l, Elevation: l.Elevation))
        .ToList();

      using var group = new TransactionGroup(doc, "Importar niveles CivilesPro");
      group.Start();

      try
      {
        using (var tr = new Transaction(doc, "Crear o actualizar niveles"))
        {
          tr.Start();

          foreach (var row in preparedRows)
          {
            if (existingLevels.TryGetValue(row.Name, out var existingLevel))
            {
              var diff = Math.Abs(existingLevel.Elevation - row.ElevationInternal);
              if (diff <= tolerance)
              {
                unchangedLevels.Add(existingLevel);
                continue;
              }

              var elevParam = existingLevel.get_Parameter(BuiltInParameter.LEVEL_ELEV);
              if (elevParam is null || elevParam.IsReadOnly)
              {
                warnings.Add($"Nivel '{existingLevel.Name}': no se pudo actualizar la elevación (parámetro bloqueado).");
                unchangedLevels.Add(existingLevel);
                continue;
              }

              elevParam.Set(row.ElevationInternal);
              updatedLevels.Add(existingLevel);
              duplicateElevationCheck.Add((existingLevel, row.ElevationInternal));
              continue;
            }

            if (duplicateElevationCheck.Any(tuple => Math.Abs(tuple.Elevation - row.ElevationInternal) <= tolerance))
            {
              warnings.Add($"Fila {row.SourceRow}: existe otro nivel con una elevación similar, se omitió '{row.Name}'.");
              continue;
            }

            Level newLevel;
            try
            {
              newLevel = Level.Create(doc, row.ElevationInternal);
            }
            catch (Exception ex)
            {
              warnings.Add($"Fila {row.SourceRow}: no se pudo crear el nivel '{row.Name}'. {ex.Message}");
              continue;
            }

            try
            {
              newLevel.Name = row.Name;
            }
            catch (Exception ex)
            {
              warnings.Add($"Nivel creado en fila {row.SourceRow}: no se pudo asignar el nombre '{row.Name}'. {ex.Message}");
              continue;
            }

            existingLevels[newLevel.Name] = newLevel;
            createdLevels.Add(newLevel);
            duplicateElevationCheck.Add((newLevel, row.ElevationInternal));
          }

          tr.Commit();
        }

        var createdViews = CreateFloorPlans(doc, createdLevels, warnings);

        group.Assimilate();

        var summary = BuildSummary(filePath, unitTypeId.Value, createdLevels, updatedLevels, unchangedLevels, createdViews, warnings);
        TaskDialog.Show("CivilesPro - Importar Niveles", summary);
      }
      catch (Exception ex)
      {
        Logger.Error("Ocurrió un error al importar los niveles.", ex);
        message = ex.Message;
        return Result.Failed;
      }

      return Result.Succeeded;
    }

    private static string? PromptForCsvFile()
    {
      var dialog = new OpenFileDialog
      {
        Title = "Selecciona el archivo de niveles",
        Filter = "Archivo CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*",
        CheckFileExists = true,
        CheckPathExists = true,
        Multiselect = false
      };

      return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static ForgeTypeId? PromptForUnits()
    {
      var dialog = new TaskDialog("Unidades de elevación")
      {
        MainInstruction = "¿En qué unidades están las elevaciones del archivo?",
        MainContent = "Selecciona la unidad correspondiente antes de importar.",
        AllowCancellation = true,
        CommonButtons = TaskDialogCommonButtons.Cancel
      };

      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Metros (m)");
      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Milímetros (mm)");
      dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Pies (ft)");

      var result = dialog.Show();

      return result switch
      {
        TaskDialogResult.CommandLink1 => UnitTypeId.Meters,
        TaskDialogResult.CommandLink2 => UnitTypeId.Millimeters,
        TaskDialogResult.CommandLink3 => UnitTypeId.Feet,
        _ => null
      };
    }

    private static bool TryParseDouble(string value, out double result)
    {
      var styles = NumberStyles.Float | NumberStyles.AllowThousands;
      value = value.Trim();
      return double.TryParse(value, styles, CultureInfo.InvariantCulture, out result)
        || double.TryParse(value, styles, CultureInfo.CurrentCulture, out result)
        || double.TryParse(value.Replace(',', '.'), styles, CultureInfo.InvariantCulture, out result);
    }

    private static IList<string> CreateFloorPlans(Document doc, IReadOnlyList<Level> levels, IList<string> warnings)
    {
      var createdViews = new List<string>();
      if (levels.Count == 0)
      {
        return createdViews;
      }

      var viewType = new FilteredElementCollector(doc)
        .OfClass(typeof(ViewFamilyType))
        .Cast<ViewFamilyType>()
        .FirstOrDefault(v => v.ViewFamily == ViewFamily.FloorPlan);

      if (viewType == null)
      {
        warnings.Add("No se encontró un tipo de vista de planta. No se crearán vistas.");
        return createdViews;
      }

      var existingNames = new HashSet<string>(new FilteredElementCollector(doc)
        .OfClass(typeof(View))
        .Cast<View>()
        .Select(v => v.Name));

      using (var tr = new Transaction(doc, "Crear vistas de planta"))
      {
        tr.Start();

        foreach (var level in levels)
        {
          ViewPlan? view = null;
          try
          {
            view = ViewPlan.Create(doc, viewType.Id, level.Id);
          }
          catch (Exception ex)
          {
            warnings.Add($"No se pudo crear la vista para el nivel '{level.Name}'. {ex.Message}");
            continue;
          }

          var viewName = GenerateViewName(level.Name, existingNames);
          view.Name = viewName;
          view.Scale = 100;
          if (view.CropBoxActive == false)
          {
            view.CropBoxActive = true;
          }
          view.CropBoxVisible = false;

          existingNames.Add(viewName);
          createdViews.Add(viewName);
        }

        tr.Commit();
      }

      return createdViews;
    }

    private static string GenerateViewName(string levelName, ISet<string> existingNames)
    {
      var baseName = $"Planta - {levelName}";
      if (!existingNames.Contains(baseName))
      {
        return baseName;
      }

      var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
      var candidate = $"{baseName} ({timestamp})";
      var index = 1;
      while (existingNames.Contains(candidate))
      {
        candidate = $"{baseName} ({timestamp}_{index++})";
      }

      return candidate;
    }

    private static string BuildSummary(
      string filePath,
      ForgeTypeId unitTypeId,
      IReadOnlyCollection<Level> createdLevels,
      IReadOnlyCollection<Level> updatedLevels,
      IReadOnlyCollection<Level> unchangedLevels,
      IReadOnlyCollection<string> createdViews,
      IReadOnlyCollection<string> warnings)
    {
      var sb = new StringBuilder();
      sb.AppendLine($"Archivo: {Path.GetFileName(filePath)}");
      sb.AppendLine($"Unidades de entrada: {DescribeUnit(unitTypeId)}");
      sb.AppendLine();
      sb.AppendLine($"Niveles creados: {createdLevels.Count}");
      if (createdLevels.Count > 0)
      {
        foreach (var level in createdLevels)
        {
          sb.AppendLine($"  • {level.Name}");
        }
      }

      sb.AppendLine($"Niveles actualizados: {updatedLevels.Count}");
      if (updatedLevels.Count > 0)
      {
        foreach (var level in updatedLevels)
        {
          sb.AppendLine($"  • {level.Name}");
        }
      }

      sb.AppendLine($"Niveles sin cambios: {unchangedLevels.Count}");

      sb.AppendLine($"Vistas de planta creadas: {createdViews.Count}");
      if (createdViews.Count > 0)
      {
        foreach (var viewName in createdViews)
        {
          sb.AppendLine($"  • {viewName}");
        }
      }

      if (warnings.Count > 0)
      {
        sb.AppendLine();
        sb.AppendLine("Observaciones:");
        foreach (var warning in warnings)
        {
          sb.AppendLine($"  • {warning}");
        }
      }

      return sb.ToString();
    }

    private static void ShowSummary(string mainMessage, IEnumerable<string> warnings, IEnumerable<string> details)
    {
      var sb = new StringBuilder();
      sb.AppendLine(mainMessage);
      foreach (var warning in warnings)
      {
        sb.AppendLine($" - {warning}");
      }

      foreach (var item in details)
      {
        sb.AppendLine($" - {item}");
      }

      TaskDialog.Show("CivilesPro - Importar Niveles", sb.ToString());
    }

    private static string DescribeUnit(ForgeTypeId unitTypeId)
    {
      if (unitTypeId == UnitTypeId.Meters) return "Metros";
      if (unitTypeId == UnitTypeId.Millimeters) return "Milímetros";
      if (unitTypeId == UnitTypeId.Feet) return "Pies";
      return unitTypeId.TypeId;
    }

    private sealed record PreparedRow(int SourceRow, string Name, double ElevationInternal);
  }
}
