using System;
using Autodesk.Revit.UI;

namespace CivilesPro.RevitTools.Core
{
  internal static class Logger
  {
    public static void Info(string msg) => System.Diagnostics.Debug.WriteLine("[CivilesPro] " + msg);

    public static void Error(string msg, Exception? ex = null)
    {
      System.Diagnostics.Debug.WriteLine("[CivilesPro][ERR] " + msg + " " + ex?.Message);
      try { TaskDialog.Show("CivilesPro", msg + (ex != null ? "\n" + ex.Message : "")); } catch { }
    }
  }
}
