using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Microsoft.Win32; // Para OpenFileDialog

namespace CivilesPro.RevitTools.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ImportarNivelesCommand : IExternalCommand
    {
        // === Opciones por defecto (puedes exponerlas luego en un form) ===
        private const bool CREATE_FLOORPLAN_VIEWS = true;
        private const int DEFAULT_VIEW_SCALE = 100;
        private const double ELEV_TOL_MM = 1.0; // tolerancia para igualar cotas (1 mm)

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;
            if (doc == null)
            {
                message = "No hay documento activo.";
                return Result.Failed;
            }

            // 1) Seleccionar archivo CSV
            var ofd = new OpenFileDialog
            {
                Title = "Seleccionar CSV de Niveles",
                Filter = "CSV (*.csv)|*.csv",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog() != true)
                return Result.Cancelled;

            string path = ofd.FileName;

            // 2) Parsear CSV robusto
             ParsedRows rows;
            try
            {
                rows = ReadCsv(path);
                if (rows.Count == 0)
                {
                    TaskDialog.Show("Importar Niveles", "El CSV no contiene filas válidas (revisa encabezados y datos).");
                    return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                message = "No se pudo leer el CSV: " + ex.Message;
                return Result.Failed;
            }

            // 3) Preparar búsqueda de niveles existentes y ViewFamilyType
            var sb = new StringBuilder();
            sb.AppendLine($"Archivo: {Path.GetFileName(path)}");
            sb.AppendLine($"Filas leídas: {rows.Count}");

            // Diccionario de niveles por nombre
            var nivelesPorNombre = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToDictionary(l => l.Name, l => l);

            // Lista de niveles existentes (para matcheo por elevación)
            var nivelesExistentes = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            // Buscar un ViewFamilyType de FloorPlan
            ViewFamilyType floorPlanType = null;
            try
            {
                floorPlanType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.FloorPlan);
            }
            catch { /* ignore */ }

            if (CREATE_FLOORPLAN_VIEWS && floorPlanType == null)
            {
                TaskDialog.Show("Importar Niveles",
                    "No se encontró un ViewFamilyType de Floor Plan. Se crearán niveles sin vistas.");
            }

            // 4) Calcular tolerancias y preparar conversiones
            double tolFeet = UnitUtils.ConvertToInternalUnits(ELEV_TOL_MM, UnitTypeId.Millimeters);

            // 5) Ordenar por elevación (ya en pies al convertir)
            foreach (var r in rows)
            {
                try
                {
                    r.ElevFeet = ToInternalFeet(r.ElevValue, r.UnitHint, rows.UnitFromHeader);
                }
                catch (Exception ex)
                {
                    r.Error = "Unidad/elevación inválida: " + ex.Message;
                }
            }

            var validRows = rows.Where(r => string.IsNullOrWhiteSpace(r.Error))
                                .OrderBy(r => r.ElevFeet)
                                .ToList();

            if (validRows.Count == 0)
            {
                TaskDialog.Show("Importar Niveles", "No hay filas válidas para procesar (todas con error).");
                return Result.Cancelled;
            }

            // 6) Transacciones
            var tg = new TransactionGroup(doc, "Importar Niveles (CivilesPro)");
            tg.Start();

            int created = 0, updated = 0, skipped = 0, viewCreated = 0;

            try
            {
                using (var t = new Transaction(doc, "Crear/Actualizar Niveles"))
                {
                    t.Start();

                    foreach (var r in validRows)
                    {
                        // Buscar por nombre primero
                        Level lvl = null;
                        if (!string.IsNullOrWhiteSpace(r.Name) && nivelesPorNombre.TryGetValue(r.Name, out var found))
                        {
                            lvl = found;
                        }

                        // Si no por nombre, buscar por elevación con tolerancia
                        if (lvl == null)
                        {
                            lvl = nivelesExistentes.FirstOrDefault(L => Math.Abs(L.Elevation - r.ElevFeet) <= tolFeet);
                        }

                        if (lvl == null)
                        {
                            // Crear nivel nuevo
                            lvl = Level.Create(doc, r.ElevFeet);
                            SafeSetName(doc, lvl, r.Name, out string finalName);
                            if (!nivelesPorNombre.ContainsKey(finalName))
                                nivelesPorNombre.Add(finalName, lvl);

                            nivelesExistentes.Add(lvl);
                            created++;
                            sb.AppendLine($"[CREADO] Nivel \"{finalName}\" @ {FeetToString(r.ElevFeet)}");
                        }
                        else
                        {
                            // Actualizar elevación si difiere más que la tolerancia
                            if (Math.Abs(lvl.Elevation - r.ElevFeet) > tolFeet)
                            {
                                double old = lvl.Elevation;
                                lvl.Elevation = r.ElevFeet;
                                updated++;
                                sb.AppendLine($"[ACTUALIZADO] Nivel \"{lvl.Name}\": {FeetToString(old)} → {FeetToString(r.ElevFeet)}");
                            }
                            else
                            {
                                skipped++;
                                sb.AppendLine($"[OMITIDO] Nivel \"{lvl.Name}\" ya coincide (±{ELEV_TOL_MM} mm).");
                            }
                        }

                        // Crear vista de planta si corresponde
                        if (CREATE_FLOORPLAN_VIEWS && floorPlanType != null && lvl != null)
                        {
                            try
                            {
                                ViewPlan vp = ViewPlan.Create(doc, floorPlanType.Id, lvl.Id);

                                // Escala
                                if (DEFAULT_VIEW_SCALE > 0) vp.Scale = DEFAULT_VIEW_SCALE;

                                // Crop: activo y oculto
                                vp.CropBoxActive = true;
                                vp.CropBoxVisible = false;

                                // Nombre de vista (evita colisión con sufijo timestamp compacto)
                                string desired = lvl.Name;
                                vp.Name = UniqueViewName(doc, desired);
                                viewCreated++;
                                sb.AppendLine($"   └─ [VISTA] {vp.Name} (1:{vp.Scale})");
                            }
                            catch (Exception ex)
                            {
                                sb.AppendLine($"   └─ [VISTA-ERROR] {ex.Message}");
                            }
                        }
                    }

                    t.Commit();
                }

                tg.Assimilate();
            }
            catch (Exception ex)
            {
                tg.RollBack();
                message = "Error creando/actualizando niveles: " + ex.Message;
                return Result.Failed;
            }

            // 7) Reporte final + guardar log
            string logPath = Path.Combine(Path.GetTempPath(),
                $"ImportarNiveles_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            try { File.WriteAllText(logPath, sb.ToString(), Encoding.UTF8); } catch { /* ignore */ }

            TaskDialog.Show("Importar Niveles",
                $"Listo.\n" +
                $"- Niveles creados: {created}\n" +
                $"- Niveles actualizados: {updated}\n" +
                $"- Niveles omitidos: {skipped}\n" +
                (CREATE_FLOORPLAN_VIEWS ? $"- Vistas creadas: {viewCreated}\n" : "") +
                $"\nLog: {logPath}");

            return Result.Succeeded;
        }

        // ==== Helpers ====

        private static string UniqueViewName(Document doc, string baseName)
        {
            string name = baseName;
            int i = 1;
            while (new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Any(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                name = $"{baseName} - {DateTime.Now:yyyyMMddHHmmss}-{i}";
                i++;
            }
            return name;
        }

        private static void SafeSetName(Document doc, Element el, string desired, out string finalName)
        {
            // Si desired está vacío, genera uno genérico
            if (string.IsNullOrWhiteSpace(desired))
                desired = "Nivel";

            string name = desired;
            int i = 1;
            while (ElementNameExists(doc, name))
            {
                name = $"{desired} ({i})";
                i++;
            }
            el.Name = name;
            finalName = name;
        }

        private static bool ElementNameExists(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Any(e => e.Name != null && e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        private static string FeetToString(double feet)
        {
            // Sólo para el log
            double meters = UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Meters);
            return $"{feet:0.###} ft ({meters:0.###} m)";
        }

        private static double ToInternalFeet(double value, UnitHint rowUnit, UnitHint headerUnit)
        {
            // Prioridad: unidad por fila > unidad por cabecera > m si nada se indica
            UnitHint unit = rowUnit != UnitHint.Unknown ? rowUnit :
                            headerUnit != UnitHint.Unknown ? headerUnit :
                            UnitHint.Meters;

            ForgeTypeId id = UnitTypeId.Meters;
            switch (unit)
            {
                case UnitHint.Millimeters: id = UnitTypeId.Millimeters; break;
                case UnitHint.Meters:      id = UnitTypeId.Meters;      break;
                case UnitHint.Feet:        id = UnitTypeId.Feet;        break;
            }

            return UnitUtils.ConvertToInternalUnits(value, id); // a pies internos
        }

        private class RowNivel
        {
            public string Name = string.Empty;
            public double ElevValue;
            public UnitHint UnitHint;
            public double ElevFeet;
            public string Error = string.Empty;
        }

        private enum UnitHint { Unknown, Millimeters, Meters, Feet }

        private class ParsedRows : List<RowNivel>
        {
        public UnitHint UnitFromHeader { get; set; } = UnitHint.Unknown;
        }

        private static ParsedRows ReadCsv(string path)
        {
            var lines = File.ReadAllLines(path, Encoding.UTF8)
                            .Where(l => !string.IsNullOrWhiteSpace(l))
                            .ToList();

            if (lines.Count < 2)
                return new ParsedRows();

            // Encabezado
            var headerCells = SplitCsvLine(lines[0]);
            int idxNombre = IndexOf(headerCells, new[] { "nombre", "name", "nivel" });
            int idxElev = IndexOf(headerCells, new[] { "elevacion", "elevación", "elevation", "cota" });
            int idxUnidad = IndexOf(headerCells, new[] { "unidad", "unit" });

            if (idxNombre < 0 || idxElev < 0)
                throw new Exception("Encabezados requeridos: 'Nombre' y 'Elevacion'.");

            // Detectar unidad en la cabecera de Elevación, p.ej. "Elevacion (m)" o "[mm]"
            var parsed = new ParsedRows();
            parsed.UnitFromHeader = DetectUnit(headerCells[idxElev]);

            // Filas
            for (int i = 1; i < lines.Count; i++)
            {
                var cells = SplitCsvLine(lines[i]);
                if (cells.Count == 0) continue;

                string nombre = SafeGet(cells, idxNombre);
                string elevStr = SafeGet(cells, idxElev);
                string unidadStr = idxUnidad >= 0 ? SafeGet(cells, idxUnidad) : null;

                var row = new RowNivel { Name = nombre?.Trim() };

                if (!double.TryParse(elevStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double elev))
                {
                    // Reintento con cultura local por si usan coma decimal
                    if (!double.TryParse(elevStr, NumberStyles.Float, CultureInfo.CurrentCulture, out elev))
                    {
                        row.Error = $"Fila {i + 1}: Elevación no numérica ({elevStr}).";
                        parsed.Add(row);
                        continue;
                    }
                }
                row.ElevValue = elev;

                row.UnitHint = ParseUnit(unidadStr ?? "");
                parsed.Add(row);
            }

            return parsed;
        }

        private static List<string> SplitCsvLine(string line)
        {
            // Split CSV muy simple (sin comillas escapadas complejas). Suficiente para nuestra plantilla.
            return line.Split(',')
                       .Select(s => s.Trim())
                       .ToList();
        }

        private static int IndexOf(List<string> headers, IEnumerable<string> candidates)
        {
            int idx = -1;
            int i = 0;
            foreach (var h in headers)
            {
                string s = (h ?? "").ToLowerInvariant();
                // Limpia sufijos de unidad en cabecera de elevación
                s = s.Replace("(m)", "").Replace("[m]", "")
                     .Replace("(mm)", "").Replace("[mm]", "")
                     .Replace("(ft)", "").Replace("[ft]", "").Trim();

                if (candidates.Any(c => s.Equals(c, StringComparison.OrdinalIgnoreCase)))
                    return i;
                i++;
            }
            return idx;
        }

        private static string? SafeGet(List<string> cells, int index)
        {
            if (index < 0 || index >= cells.Count) return null;
            return cells[index];
        }

       private static UnitHint DetectUnit(string? headerCell)
        {
            if (string.IsNullOrWhiteSpace(headerCell)) return UnitHint.Unknown;
            string s = headerCell.ToLowerInvariant();
            if (s.Contains("(mm)") || s.Contains("[mm]")) return UnitHint.Millimeters;
            if (s.Contains("(m)")  || s.Contains("[m]"))  return UnitHint.Meters;
            if (s.Contains("(ft)") || s.Contains("[ft]")) return UnitHint.Feet;
            return UnitHint.Unknown;
        }

         private static UnitHint ParseUnit(string? unitCell)
        {
            if (string.IsNullOrWhiteSpace(unitCell)) return UnitHint.Unknown;
            string s = unitCell.Trim().ToLowerInvariant();
            if (s == "mm" || s == "milimetros" || s == "milímetros") return UnitHint.Millimeters;
            if (s == "m"  || s == "metro" || s == "metros") return UnitHint.Meters;
            if (s == "ft" || s == "pie" || s == "pies") return UnitHint.Feet;
            return UnitHint.Unknown;
        }
    }
}
