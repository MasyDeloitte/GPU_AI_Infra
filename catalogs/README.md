# Versioned sizing catalogs

These files are generated from `Forge_Catalog_Reference_v1.xlsx` and preserve the four calculation catalogs needed by the sizing engine:

- `Model_Mapping_Catalog.json`
- `Model_Selection.json`
- `Hardware_Mapping_Catalog.json`
- `GPU_Server_Selection.json`
- `manifest.json`

The exporter preserves formula text instead of replacing formulas with cached values. This makes the catalog export auditable and suitable for migrating calculation logic to a server-side engine later.

Regenerate after changing the workbook:

```powershell
.venv\Scripts\python.exe scripts\export-catalogs.py
```

The application should treat the catalog version as part of every saved project and every generated BOM. Catalog updates should be reviewed and tested as data migrations, not silently overwritten in production.
