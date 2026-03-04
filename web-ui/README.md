# DEEPSOIL WASM UI (Pyodide)

## Run

Serve from the `Deepsoil` directory so `../disp_core.py` is reachable by `web-ui/worker.js`.

```powershell
cd "<path-to-dstotaldisp>"
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
  - supports both X/Y paired processing and single-file processing
  - Method-2 export toggle (default on)
  - Method-3 export toggle (default on)
  - processing filter controls:
    - `Apply Baseline` (default off)
    - `Apply Filtering` (default off)
    - `Processing Order`
    - `Filter Domain`
    - `Baseline Method`
    - `Filter Config`
    - `Filter Type`
    - `Base Reference` (`Input Motion Proxy` / `Deepest Layer Proxy`)
    - `F Low`, `F High`, `Order`
- Output sheets per pair:
  - `Strain_Relative`
  - `Legacy_Methods`
  - `Comparison`
  - `Depth_Profiles`
  - `Profile_BaseCorrected`
  - `Direction_X_Time`
  - `Direction_Y_Time`
  - `Resultant_Time`
  - `TBDY_Total_X_Time`
  - `TBDY_Total_Y_Time`
  - `TBDY_Total_Resultant_Time`
- Output sheets per single file:
  - `Single_Direction_Summary`
  - `Direction_Time`
  - `Strain_Relative_Time`
  - `TBDY_Total_Time`
  - `InputProxy_Relative_Time`
- Additional output files:
  - `output_method2_<record>.xlsx` (per input file)
  - `output_method3_profiles_all.xlsx` (aggregate X/Y depth profiles)
