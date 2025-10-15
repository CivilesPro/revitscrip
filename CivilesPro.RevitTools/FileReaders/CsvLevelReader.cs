using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CivilesPro.RevitTools.FileReaders
{
  internal record LevelData(string Nombre, double ElevacionFeet);

  internal static class CsvLevelReader
  {
    // CSV esperado: Nombre;Elevacion_m
    public static IEnumerable<LevelData> Read(string path)
    {
      using var sr = new StreamReader(path);
      string? line;
      var sep = ';';
      while ((line = sr.ReadLine()) != null)
      {
        if (string.IsNullOrWhiteSpace(line)) continue;
        var parts = line.Split(sep);
        if (parts.Length < 2) continue;
        var nombre = parts[0].Trim();
        if (double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var elevM))
        {
          var elevFt = elevM * 3.280839895;
          yield return new LevelData(nombre, elevFt);
        }
      }
    }
  }
}
