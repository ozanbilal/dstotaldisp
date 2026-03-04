# DEEPSOIL WASM UI (Pyodide)

## Run

Serve from the `Deepsoil` directory so `../disp_core.py` is reachable by `web-ui/worker.js`.

```powershell
cd "H:\Drive'ım\Ortak\Bildiri_Makale\Gapping-Nongapping\Deepsoil"
python -m http.server 8010
```

Then open:

`http://localhost:8010/web-ui/`

## Notes

- Browser target: Chrome / Edge.
- Uses client-side Pyodide (no backend service).
- Input mode:
  - folder selection (`webkitdirectory`)
  - direct multi-file selection (`.xlsx`)
- Output sheets per pair:
  - `Strain_Relative`
  - `Legacy_Methods`
  - `Comparison`
