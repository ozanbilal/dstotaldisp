# DeepSoil Total Displacement Developer Guide

Son guncelleme: 2026-06-10

Bu dokuman, yeni bir session veya agent repoya girdiginde hizli sekilde dogru dosyalara, veri sozlesmesine, testlere ve deploy kosullarina ulasabilsin diye tutulur. Kullaniciya acik asil kilavuz `docs/user-guide.md` dosyasidir; bu dosya gelistirici ve bakim notlaridir.

## Hizli Repo Haritasi

- `disp_core.py`: ortak hesap cekirdegi. CLI ve web worker ayni `process_batch_files(...)` API'sini kullanir.
- `GetDisp4.py`: CLI giris noktasi.
- `web-ui/index.html`: summary-first ana arayuz.
- `web-ui/source_app.js`: UI state, dosya secimi, worker orchestration, result merge ve download akislarinin ana dosyasi.
- `web-ui/worker.js`: Pyodide runtime'i kurar, `disp_core.py` ve `web-ui/py/pyodide_entry.py` dosyalarini fetch eder, dosyalari `/input` altina yazar ve Python'u calistirir.
- `web-ui/py/pyodide_entry.py`: JS worker ile `disp_core.process_batch_files` arasindaki ince Python koprusu.
- `web-ui/viewer.mjs`: detayli source explorer modelini kurar.
- `web-ui/summary_viewer.mjs`: summary-first toplam deplasman sahnesini kurar.
- `web-ui/server.js`: Railway/canli ortam icin Node auth gatekeeper ve static file server.
- `Dockerfile`: canli Node image'i; `web-ui/` ile birlikte repo kokundeki `disp_core.py` dosyasini da image icine kopyalamalidir.
- `docs/user-guide.md`: kullaniciya acik kanonik kilavuz.
- `web-ui/docs.html`: `docs/user-guide.md` dosyasini tarayicida render eder.
- `README.md`: CLI, cikti dosyalari, sheet adlari ve teknik davranis ozeti.
- `AGENTS.md`: zorunlu bakim kurallari.
- `tests/`: Python ve Node regresyon testleri.

## Runtime Modlari

### CLI

`GetDisp4.py`, dosyalari lokal diskten okur ve `disp_core.process_batch_files(...)` fonksiyonuna aktarir. CLI davranisi, cikti dosya adlari, workbook sheet adlari veya veri yorumlari degistiginde `README.md` guncellenmelidir.

### Lokal Web UI

Repo kokunden static server baslat:

```powershell
cd "<path-to-dstotaldisp>"
python -m http.server 8010
```

Sonra ac:

```text
http://localhost:8010/web-ui/
```

Bu modda analiz tamamen browser icinde Pyodide ile calisir. `web-ui/worker.js`, `disp_core.py` icin once `../disp_core.py`, sonra `/disp_core.py` URL'lerini dener. Bu nedenle server repo kokunden baslamalidir.

### Canli / Railway Node Gatekeeper

Canli yolda `Dockerfile` Node image'i uretir:

```dockerfile
COPY web-ui/ ./
COPY disp_core.py ./disp_core.py
CMD ["node", "server.js"]
```

`web-ui/server.js` su sozlesmeyi korumak zorundadir:

- `/health` auth olmadan JSON doner.
- Uretim modunda diger istekler SSO/auth gatekeeper'dan gecer.
- `/disp_core.py` ve `/web-ui/disp_core.py`, HTML degil gercek Python kaynak dosyasini `text/x-python` olarak dondurur.
- Uzantili eksik dosyalar SPA fallback ile `index.html` dondurmemelidir; `404 text/plain` donmelidir.
- `/web-ui` static alias'i korunmalidir, cunku tarayici bu yol altindan calisir.

Bu sozlesme bozulursa Pyodide, HTML'i Python diye import etmeye calisir ve genellikle su tip hata gorulur:

```text
SyntaxError: invalid decimal literal
```

Stack trace icinde `disp_core.py` satirinda `<p class="eyebrow">...` gibi HTML gorunuyorsa sebep neredeyse kesin olarak `/disp_core.py` endpoint'inin HTML dondurmesidir.

## Remote ve Deploy Notu

Bu repoda en az iki remote kullanilabiliyor:

- `origin`: ana repo.
- `geodisp`: canli domain icin kullanilmis deploy remote'u.

Canli siteyle ilgili degisikliklerde remote durumunu komutla dogrula:

```powershell
git remote -v
git log --oneline --decorate -5
git ls-remote origin refs/heads/main
git ls-remote geodisp refs/heads/main
```

Canli deploy'un hangi remote'u izledigini varsayma. Son bilinen durumda `https://disp.geoproje.com.tr/` icin kritik push `geodisp/main` tarafina da gerekiyordu. Canli davranisi etkileyen fixlerde `origin/main` ve aktif deploy remote'u ayni commit'e geldikten sonra isi tamamlanmis say.

## Pyodide Veri Akisi

1. Kullanici dosyalari `web-ui/source_app.js` tarafinda secer.
2. Folder select modunda desteklenmeyen yardimci dosyalar filtrelenir.
3. Ayni isimli `deepsoilout.db3` dosyalari parent klasor adiyla mantiksal olarak yeniden adlandirilir.
4. Worker baslatilir: `new Worker("./worker.js?v=<APP_VERSION>")`.
5. `web-ui/worker.js`, Pyodide paketlerini yukler: `numpy`, `pandas`, `sqlite3`, gerekirse `openpyxl` ve `micropip`.
6. Worker, `disp_core.py` ve `web-ui/py/pyodide_entry.py` kaynaklarini fetch eder.
7. Kaynaklar Pyodide FS icinde `/app/disp_core.py` ve `/app/pyodide_entry.py` olarak yazilir.
8. Secilen input dosyalari `/input` altina yazilir.
9. `pyodide_entry.run_batch_from_fs(...)`, `process_batch_files(...)` fonksiyonunu cagirir.
10. Sonuc JS tarafina output dosyalari, ozetler, loglar, `sourceCatalog` ve `summaryCatalog` ile doner.

## Web Veri Sozlesmesi

`process_batch_files(..., _returnWebResults=True)` sonucu summary-first UI icin su alanlari korumalidir:

- `sourceCatalog`: detayli explorer'in ham source/family/chart/layer modelidir.
- `summaryCatalog`: ana toplam deplasman sahnesinin kayit/variant modelidir.
- `previewCharts`: geriye uyumluluk alani; ana summary-first arayuzun birincil veri kaynagi degildir.
- `outputs` veya output file listeleri: workbook ve ZIP indirme akisini besler.

Kural:

- `summaryCatalog` bozulursa ana sahne bos veya yanlis secimle acilir.
- `sourceCatalog` bozulursa detayli inceleme ve chart drilldown bozulur.
- Yeni kullaniciya gorunen yontem, varyant, chart ailesi veya veri yorumu eklenirse `docs/user-guide.md` ve ilgili test guncellenmelidir.

## Summary ve Source Ayrimi

Summary-first yuzey kullaniciya karar verilecek toplam deplasman profilini gosterir. Detayli explorer ise uretilen tum serileri incelemek icindir.

- `summaryCatalog` icinde her kayit varyant listesi tasir.
- Varyantlar guven sirasina gore siralanir.
- Gecersiz varyantlar cizilmez; nedenleri kullaniciya acik not olarak kalir.
- `sourceCatalog` icinde aileler vardir. Ornek: `db-motion`, `db-layer-series`, zaman serisi aileleri, profile aileleri.

Bir pair isleminde tek kullanici kaydi icin birden fazla source olusabilir. Bu nedenle `Viewer sources` sayaci dosya adedi degil explorer source kaydi sayisidir.

## Excel / Strain Akisi

`.xlsx` girdiler icin normal akis:

1. X/Y pair tespiti ad kaliplarindan yapilir.
2. Eslesmeyen dosyalar single-file olarak islenir.
3. Layer sheet'lerinden strain serileri okunur.
4. `u_rel_base = sum(gamma * h)` katman bazli birikimli goreceli deplasman uretilir.
5. `Input Motion` veya deepest layer ivmesinden base proxy hesaplanabilir.
6. `u_total = u_base_proxy + u_rel_base` varyantlari summary sahnesine yazilir.
7. Workbook ciktilari pair, single, Method-2 ve Method-3 akislariyla uretilir.

Bu akis degisirse `README.md`, `docs/user-guide.md` ve Python testleri birlikte guncellenmelidir.

## DB Direct Akisi

`.db` / `.db3` girdiler icin `useDbDirect` aktifse `process_batch_files(...)`, `_process_db_batch_files(...)` yoluna gider.

Ana fonksiyonlar:

- `_process_db_batch_files(...)`: DB adaylarini ayirir, pair/single islemeyi yonetir, output workbook'lari ve catalog'lari birlestirir.
- `_resolve_db_xy_pairs(...)`: parent folder ve dosya adlarindan fuzzy X/Y esleme yapar.
- `_read_db_disp_bundle(...)`: veritabanindan displacement ve opsiyonel layer serilerini okur.
- `_build_db_single_source_catalog_entry(...)` / `_build_db_pair_source_catalog_entry(...)`: detay explorer kaynaklarini kurar.
- `_build_db_single_summary_entry(...)` / `_build_db_pair_summary_entry(...)`: summary sahnesi icin DB direct varyantlarini kurar.

Desteklenen DB semalari:

- Oncelikli tablo: `VEL_DISP`
  - `LAYERn_DISP_TOTAL`
  - `LAYERn_DISP_RELATIVE`
- Fallback sema:
  - `TIME_HISTORIES.LAYER#_DISP`
  - `PROFILES.MIN_DISP_RELATIVE`
  - `PROFILES.MAX_DISP_RELATIVE`

Opsiyonel detay serileri kolon deseniyle okunur:

- `Layer#_Accel`
- `Layer#_Vel`
- `Layer#_Disp`
- `Layer#_Arias`
- `Layer#_Strain`
- `Layer#_Stress`
- `Layer#_RS`
- `Layer#_FAS`
- `Layer#_FAS_Ratio`

DB direct modunda baseline, filter, integration compare ve base reference kontrolleri uygulanmaz. Bu kontroller UI tarafinda devre disi kalmalidir.

## DB Pairing Notlari

Folder select ile secilen DeepSoil batch ciktilarinda her kayit genelde ayni dosya adini tasir:

```text
Motion_DD2_X_..._H1/deepsoilout.db3
Motion_DD2_Y_..._H2/deepsoilout.db3
```

Browser File API dosya adini sadece `deepsoilout.db3` olarak verebildigi icin UI, parent klasor adini dosya kimligine katar. DB pair resolver da X/Y tokenlarini, RSN/kayit tokenlarini ve noise token listesini kullanarak fuzzy pair kurar.

Bu alanla ilgili buglarda once su senaryoyu dogrula:

```text
C:\DEEPSOIL\Batch Output\Batch_run_55\profile_1
```

Beklenen davranis: parent klasorlerden gelen coklu `deepsoilout.db3` dosyalari pair olarak yakalanir; `VEL_DISP` yoksa `TIME_HISTORIES`/`PROFILES` fallback calisir; `sourceCatalog` DB source'lari ve `summaryCatalog` DB direct summary kayitlarini tasir.

## Cache ve Versiyon Bump

Web tarafinda cache bust icin ayni surum degeri birden fazla yerde bulunabilir. UI/worker davranisi veya Python kaynak fetch'i degistiginde arama yap:

```powershell
rg -n "APP_VERSION|v=20[0-9]{6}" web-ui
```

Son bilinen yerler:

- `web-ui/worker.js`
- `web-ui/source_app.js`
- `web-ui/app.js`
- `web-ui/index.html`

Surum bump atlanirsa browser eski worker veya eski module dosyasini kullanabilir.

## Test Matrisi

Genel Python testleri:

```powershell
python -m pytest
```

Node/module testleri:

```powershell
Get-ChildItem tests -Filter *.mjs | ForEach-Object { node $_.FullName }
```

Hizli sozlesme testleri:

```powershell
python -m pytest tests\test_db_direct_viewer.py tests\test_docs_contract.py tests\test_viewer_shell_contract.py
node --check web-ui\server.js
node --check web-ui\worker.js
```

Canli gatekeeper endpoint'ini lokal taklit etmek icin:

1. Gecici klasore `web-ui/*` ve repo kokundeki `disp_core.py` kopyalanir.
2. `npm install --omit=dev` calistirilir.
3. `NODE_ENV=development PORT=<port> node server.js` baslatilir.
4. `/disp_core.py` endpoint'inin `text/x-python` dondurdugu ve HTML ile baslamadigi kontrol edilir.

## Dokumantasyon Kurali

Degisiklik turune gore guncelle:

- Kullaniciya gorunen UI, yontem, veri yorumu: `docs/user-guide.md`
- CLI, workbook, sheet, cikti dosya adi, veri davranisi: `README.md`
- Runtime, deploy, Pyodide sozlesmesi, agent handoff: `docs/developer-guide.md`
- Bakim kurali veya zorunlu repo sozlesmesi: `AGENTS.md`
- Site docs render baglantisi bozulursa: `web-ui/docs.html`

Dokumantasyon guncellenmeden ozellik isi tamamlanmis sayilmaz.

## Commit ve Push Kontrol Listesi

Commit oncesi:

```powershell
git status --short --branch
git diff --check
python -m pytest
Get-ChildItem tests -Filter *.mjs | ForEach-Object { node $_.FullName }
```

Canli davranis etkileniyorsa:

```powershell
git push origin main
git push geodisp main
git ls-remote origin refs/heads/main
git ls-remote geodisp refs/heads/main
```

Remote isimleri veya deploy kaynagi degismis olabilir; push yapmadan once `git remote -v` ile guncel durumu kontrol et.
