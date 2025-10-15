using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using CivilesPro.RevitTools.Core;
using Microsoft.Office.Interop.Excel;

namespace CivilesPro.RevitTools.FileReaders
{
    public static class ExcelLevelReader
    {
        /// <summary>
        /// Lee hoja 1. A2: nombre, B2: elevación (metros).
        /// Detiene cuando ambas celdas están vacías.
        /// </summary>
        public static List<LevelData> Read(string path)
        {
            var list = new List<LevelData>();
            Application app = null;
            Workbook wb = null;
            Worksheet ws = null;
            Range used = null;

            try
            {
                app = new Application { Visible = false, DisplayAlerts = false };
                wb = app.Workbooks.Open(path, ReadOnly: true);
                ws = (Worksheet)wb.Sheets[1];
                used = ws.UsedRange;

                int totalRows = used.Rows.Count;
                for (int r = 2; r <= totalRows; r++)
                {
                    var cellName = (Range)used.Cells[r, 1];
                    var cellElev = (Range)used.Cells[r, 2];

                    string name = Convert.ToString(cellName?.Value2)?.Trim();
                    object rawElev = cellElev?.Value2;

                    if (string.IsNullOrWhiteSpace(name) && rawElev == null)
                        break;

                    if (string.IsNullOrWhiteSpace(name) || rawElev == null)
                        continue;

                    double elevM;
                    string? s = Convert.ToString(rawElev)?.Trim();

                    if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out elevM))
                    {
                        if (!double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out elevM))
                        {
                            if (rawElev is double d)
                            {
                                elevM = d;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }

                    list.Add(new LevelData { RawName = name, ElevationMeters = elevM });
                }

                return list;
            }
            finally
            {
                if (used != null) Marshal.ReleaseComObject(used);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
        }
    }
}
