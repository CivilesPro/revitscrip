using System;
using System.Collections.Generic;
using System.IO;

namespace CivilesPro.RevitTools.FileReaders
{
  /// <summary>
  /// Representa una fila del archivo CSV de niveles antes de convertir unidades.
  /// </summary>
  internal sealed record RawLevelRow(int RowNumber, string Name, string ElevationText);

  internal static class CsvLevelReader
  {
    /// <summary>
    /// Lee un archivo CSV con formato <c>Nombre;Elevacion</c> u <c>Nombre,Elevacion</c>.
    /// La primera fila no vacía se asume como encabezado y se ignora.
    /// </summary>
    /// <exception cref="FileNotFoundException">Si el archivo no existe.</exception>
    public static IReadOnlyList<RawLevelRow> ReadRows(string path)
    {
      if (!File.Exists(path))
      {
        throw new FileNotFoundException($"No se encontró el archivo: {path}", path);
      }

      var rows = new List<RawLevelRow>();
      using var sr = new StreamReader(path);

      string? line;
      var lineNumber = 0;
      var headerProcessed = false;
      char delimiter = ';';

      while ((line = sr.ReadLine()) != null)
      {
        lineNumber++;
        if (string.IsNullOrWhiteSpace(line))
        {
          continue;
        }

        if (!headerProcessed)
        {
          delimiter = DetectDelimiter(line);
          headerProcessed = true;
          continue; // encabezado
        }

        var parts = line.Split(delimiter);
        var name = parts.Length > 0 ? parts[0].Trim() : string.Empty;
        var elevation = parts.Length > 1 ? parts[1].Trim() : string.Empty;

        rows.Add(new RawLevelRow(lineNumber, name, elevation));
      }

      return rows;
    }

    private static char DetectDelimiter(string headerLine)
    {
      if (headerLine.Contains(';')) return ';';
      if (headerLine.Contains(',')) return ',';
      return ';';
    }
  }
}
