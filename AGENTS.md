# AGENTS.md

Bu repo icin calisma ve bakim kurallari:

## Kapsam

- Uygulamanin ana kullanici yuzeyi `web-ui/` altindaki summary-first arayuzdur.
- Teknik hesap cekirdegi `disp_core.py` icindedir.
- Kullaniciya acik kanonik dokumantasyon `docs/user-guide.md` dosyasidir.
- `web-ui/docs.html`, bu markdown kaynaginin sitede render edilen yuzudur.

## Zorunlu dokumantasyon guncelleme kurallari

- Kullaniciya gorunen her ozellik degisiminde `docs/user-guide.md` guncellenmelidir.
- `web-ui/docs.html`, kanonik markdown kaynagini render etmeye devam etmelidir; sayfa baglantilari veya sunum bozulursa ayni iste duzeltilmelidir.
- CLI davranisi, cikti dosya adlari, sheet adlari veya veri davranisi degisirse `README.md` guncellenmelidir.
- Yeni toplam deplasman yontemi veya varyant eklenirse:
  - guven sirasi
  - gecerlilik kosullari
  - kullaniciya gosterilen kisa yorum
  dokumantasyonda ayni iste guncellenmelidir.
- Dokumantasyon guncellenmeden ozellik isi tamamlanmis sayilmaz.

## UI ve veri sozlesmesi notlari

- `process_batch_files(..., _returnWebResults=True)` sonucu hem `sourceCatalog` hem `summaryCatalog` donmelidir.
- `summaryCatalog`, kullaniciya ozet sahneyi; `sourceCatalog`, detayli incelemeyi besler.
- `previewCharts` geriye uyumluluk icin korunur, ancak summary-first yuzeyin birincil veri kaynagi degildir.

## Test beklentisi

- Kullaniciya gorunen davranis degisirse ilgili Python veya Node testleri guncellenmelidir.
- En azindan summary payload, docs varligi ve arayuz baglantilari icin regresyon kontrolu eklenmelidir.
