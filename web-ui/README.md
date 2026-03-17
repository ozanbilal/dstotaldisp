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
  - direct multi-file selection (`.xlsx`, `.db`, `.db3`)
  - supports both X/Y paired processing and single-file processing
  - `deepsoilout.db3` gibi ayni isimli DB dosyalari folder mode'da parent klasor adiyla ayrilir
  - Method-2 export toggle (default on)
  - Method-3 export toggle (default on)
  - `Use manual pairing` toggle:
    - selected file list icinden X/Y ciftlerini elle kurar
    - manual pair'e girmeyen adaylar single olarak islenir
  - `Use DB3 directly` toggle:
    - reads `VEL_DISP` and `PROFILES` tables directly from `.db/.db3`
    - disables filtering, baseline, integration-compare and base-reference controls
  - `Compare with FFT-Regularized integration` toggle (default off)
  - `Include resultant (RSS) totals in Depth_Profiles` toggle (default off)
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
  - compare toggle acikken:
    - alt low-cut filtering aciksa `F Low`, kapaliysa `0.05 Hz`
- Output sheets per pair:
  - `Strain_Relative`
  - `Legacy_Methods`
  - `Comparison`
  - `Depth_Profiles`
    - includes `Direction_X_*` / `Direction_Y_*` signed `(+max, -min)` columns
    - compare on ise `_ALT (+max, -min)` direction columns
    - resultant (RSS) totals istenirse toggle ile eklenir
  - `Profile_BaseCorrected`
  - `Direction_X_Time`
  - `Direction_Y_Time`
  - `Resultant_Time`
  - `TBDY_Total_X_Time`
  - `TBDY_Total_Y_Time`
  - `TBDY_Total_Resultant_Time`
  - compare toggle aciksa:
    - `Direction_X_Time_ALT`
    - `Direction_Y_Time_ALT`
    - `Resultant_Time_ALT`
    - `TBDY_Total_X_Time_ALT`
    - `TBDY_Total_Y_Time_ALT`
    - `TBDY_Total_Resultant_Time_ALT`
  - DB pair girislerinde:
    - `DB_Profile_Summary`
    - `DB_Depth_Profiles`
    - `DB_Total_X_Time`
    - `DB_Total_Y_Time`
    - `DB_Relative_X_Time`
    - `DB_Relative_Y_Time`
- Output sheets per single file:
  - `Single_Direction_Summary`
  - `Direction_Time`
  - `Strain_Relative_Time`
  - `TBDY_Total_Time`
  - `InputProxy_Relative_Time`
  - compare toggle aciksa:
    - `Direction_Time_ALT`
    - `Strain_Relative_Time_ALT`
    - `TBDY_Total_Time_ALT`
    - `InputProxy_Relative_Time_ALT`
  - DB single girislerinde:
    - `DB_Summary`
    - `DB_Total_Time`
    - `DB_Relative_Time`
- Additional output files:
  - `output_method2_<record>.xlsx` (per input file)
    - compare aciksa ayni dosyada `_ALT` ve `_Delta` sheetleri
  - `output_method3_profiles_all.xlsx` (aggregate X/Y depth profiles)
    - compare aciksa `Method3_Profile_*_ALT` ve `Method3_Delta_*` sheetleri
  - DB direct mode aciksa:
    - `output_method2_db_<record>.xlsx`
    - `output_method3_db_profiles_all.xlsx`
