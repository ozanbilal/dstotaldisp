# AGENTS.md

Bu repo icin calisma ve bakim kurallari:

## Kapsam

- Uygulamanin ana kullanici yuzeyi `web-ui/` altindaki summary-first arayuzdur.
- Teknik hesap cekirdegi `disp_core.py` icindedir.
- Kullaniciya acik kanonik dokumantasyon `docs/user-guide.md` dosyasidir.
- `web-ui/docs.html`, bu markdown kaynaginin sitede render edilen yuzudur.
- Gelistirici / agent handoff dokumani `docs/developer-guide.md` dosyasidir.

## Zorunlu dokumantasyon guncelleme kurallari

- Kullaniciya gorunen her ozellik degisiminde `docs/user-guide.md` guncellenmelidir.
- `web-ui/docs.html`, kanonik markdown kaynagini render etmeye devam etmelidir; sayfa baglantilari veya sunum bozulursa ayni iste duzeltilmelidir.
- CLI davranisi, cikti dosya adlari, sheet adlari veya veri davranisi degisirse `README.md` guncellenmelidir.
- Runtime, deploy, Pyodide dosya servis sozlesmesi, remote/push akisi veya agent handoff bilgisi degisirse `docs/developer-guide.md` guncellenmelidir.
- Yeni toplam deplasman yontemi veya varyant eklenirse:
  - guven sirasi
  - gecerlilik kosullari
  - kullaniciya gosterilen kisa yorum
  dokumantasyonda ayni iste guncellenmelidir.
- Dokumantasyon guncellenmeden ozellik isi tamamlanmis sayilmaz.

## Canli deploy ve Pyodide sozlesmesi

- Lokal web UI repo kokunden servis edilmelidir; `web-ui/worker.js`, `disp_core.py` dosyasini `../disp_core.py` veya `/disp_core.py` yolundan fetch eder.
- Canli Node gatekeeper `web-ui/server.js` dosyasidir; `Dockerfile`, `web-ui/` yaninda repo kokundeki `disp_core.py` dosyasini da image icine kopyalamalidir.
- `/disp_core.py` ve `/web-ui/disp_core.py` endpoint'leri HTML degil Python kaynak dosyasi dondurmelidir.
- Uzantili eksik dosyalarda SPA fallback ile `index.html` dondurulmemelidir; aksi halde Pyodide HTML'i Python diye import edip `SyntaxError` verir.
- Canli domain icin remote varsayimi yapma. Son bilinen deploy akisi `geodisp/main` remote'unu da gerektirmistir; push oncesi `git remote -v` ve `git ls-remote` ile dogrula.

## UI ve veri sozlesmesi notlari

- `process_batch_files(..., _returnWebResults=True)` sonucu hem `sourceCatalog` hem `summaryCatalog` donmelidir.
- `summaryCatalog`, kullaniciya ozet sahneyi; `sourceCatalog`, detayli incelemeyi besler.
- `previewCharts` geriye uyumluluk icin korunur, ancak summary-first yuzeyin birincil veri kaynagi degildir.
- DB direct modunda `VEL_DISP` yoksa `TIME_HISTORIES` + `PROFILES` fallback davranisi korunmalidir.
- Folder select ile gelen ayni isimli `deepsoilout.db3` dosyalari parent klasor adiyla ayirt edilmeli ve DB X/Y pair resolver bunu kullanabilmelidir.

## Test beklentisi

- Kullaniciya gorunen davranis degisirse ilgili Python veya Node testleri guncellenmelidir.
- En azindan summary payload, docs varligi ve arayuz baglantilari icin regresyon kontrolu eklenmelidir.
- Deploy/server sozlesmesi degisirse `node --check web-ui\server.js` ve ilgili docs/contract testleri calistirilmelidir.
