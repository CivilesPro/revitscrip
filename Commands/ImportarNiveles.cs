using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CivilesPro.RevitTools.Core;
using CivilesPro.RevitTools.FileReaders;
using Microsoft.Win32;

namespace CivilesPro.RevitTools.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ImportarNivelesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            var uiapp = data.Application;
            var doc = uiapp.ActiveUIDocument?.Document;
            if (doc == null)
            {
                TaskDialog.Show("Importar niveles", "Abra un documento antes de ejecutar el comando.");
                return Result.Cancelled;
            }

            // 1️⃣ Seleccionar Excel
            string path = PickExcel();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return Result.Cancelled;

            // 2️⃣ Leer datos
            List<LevelData> levels;
            try
            {
                levels = ExcelLevelReader.Read(path);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Importar niveles", "Error leyendo el Excel:\n" + ex.Message);
                return Result.Failed;
            }

            if (levels.Count == 0)
            {
                TaskDialog.Show("Importar niveles", "No se encontraron filas válidas (A2: nombre, B2: elevación).");
                return Result.Cancelled;
            }

            // 3️⃣ Ordenar por elevación
            var ordered = levels.OrderBy(l => l.ElevationMeters).ToList();

            // 4️⃣ Formatear nombres
            int digits = Math.Max(2, ordered.Count.ToString().Length);
            for (int i = 0; i < ordered.Count; i++)
            {
                var idx = (i + 1).ToString().PadLeft(digits, '0');
                string signed = ordered[i].ElevationMeters >= 0
                    ? $"+{ordered[i].ElevationMeters:0.00}m"
                    : $"{ordered[i].ElevationMeters:0.00}m";

                string raw = (ordered[i].RawName ?? string.Empty).Trim();
                string safeRaw = Regex.Replace(raw, @"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 _\.\"""“”'’\-]", string.Empty);
                ordered[i].FormattedName = $"{idx} - \"{safeRaw}\" {signed}";
            }

            // 5️⃣ Crear / renombrar
            const double FT_PER_M = 3.280839895;
            const double tolFt = 0.001;
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            using (var t = new Transaction(doc, "Importar niveles desde Excel"))
            {
                t.Start();
                foreach (var item in ordered)
                {
                    double targetFt = item.ElevationMeters * FT_PER_M;

                    var level = existing.FirstOrDefault(l => Math.Abs(l.Elevation - targetFt) < tolFt);
                    if (level == null)
                    {
                        level = Level.Create(doc, targetFt);
                        existing.Add(level);
                    }

                    string finalName = item.FormattedName;
                    string baseName = finalName;
                    int n = 1;
                    while (existing.Any(l => l.Name.Equals(finalName, StringComparison.OrdinalIgnoreCase) && l.Id != level.Id))
                    {
                        finalName = $"{baseName} ({n++})";
                    }

                    level.Name = finalName;
                }
                t.Commit();
            }

            TaskDialog.Show("Importar niveles", $"Se procesaron {ordered.Count} niveles desde:\n{path}");
            return Result.Succeeded;
        }

        private static string PickExcel()
        {
            var dlg = new OpenFileDialog
            {
                Title = "Selecciona el Excel de niveles",
                Filter = "Archivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Todos los archivos (*.*)|*.*",
                CheckFileExists = true
            };
            return dlg.ShowDialog() == true ? dlg.FileName : string.Empty;
        }
    }
}
