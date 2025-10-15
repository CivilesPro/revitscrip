# CivilesPro.RevitTools

Base de add-in para Revit (Ribbon + comando demo).

## Compilación
- Requiere .NET Framework 4.8 (para Revit 2023-2025) y SDK .NET 8.
- Ajusta las rutas de `RevitAPI.dll` y `RevitAPIUI.dll` en el `.csproj` si usas otra versión de Revit.
- Compila el target **net48**. El post-build copiará la DLL y el `.addin` a:
  - `C:\ProgramData\Autodesk\Revit\Addins\2023`
  - `C:\ProgramData\Autodesk\Revit\Addins\2024`
  - `C:\ProgramData\Autodesk\Revit\Addins\2025`

## Uso
Abre Revit → pestaña **CivilesPro** → panel **Automatizaciones** → botón **Importar Niveles**.
