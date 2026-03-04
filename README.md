# Deepsoil Total Disp Calculator

DEEPSOIL X/Y sonuc dosyalarindan toplam ve goreceli yerdegistirme profillerini ureten arac.
Hem CLI (`GetDisp4.py`) hem de tarayici tabanli WASM arayuz (`web-ui/`) ayni hesap cekirdegini (`disp_core.py`) kullanir.

## Ne Cozer?

Bu proje, su iki dunyayi ayni cikti setinde birlestirir:

- DEEPSOIL tarafi (goreceli): `u_rel(z,t) = u(z,t) - u(base,t)`
- TBDY yorumu (toplam): `u_total(z,t) = u(base,t) + u_rel(z,t)`

Not: Mevcut cekirdek, strain tabanli yolakta dogrudan `*_base_rel_*` (base-relative) uretir.
`u(base)` serisi acikca eklenirse TBDY toplami da kolon olarak uretilir; su an ana workbook'ta odak base-relative karsilastirmadir.

## Giris Beklentisi

- Dosya tipi: `.xlsx`
- Esleme kurali:
  - X dosyasi: adinda `_X_` ve sonda `_H1.xlsx`
  - Y dosyasi: ayni adin `_X_ -> _Y_`, `_H1 -> _H2` donusmus hali
- Varsayilan dislama:
  - `output_*.xlsx`
  - `~$*.xlsx`
  - `*-manip.xlsx` (yalniz `--include-manip` ile dahil edilir)

## Hesap Mantigi

## 1) Strain_Relative (oncelikli yontem)

Katman bazinda `Strain (%)` okunur, `gamma = strain/100` cevrilir.
Kalınlik `h` ile katman goreceli deplasman katkisi:

`du_k(t) = gamma_k(t) * h_k`

Tabana gore goreceli deplasman:

`u_rel_base(z_i,t) = sum_{k=i..N} du_k(t)`

Sonra zarf degerler:

- `X_base_rel_max_m = max_t |u_rel_base_x|`
- `Y_base_rel_max_m = max_t |u_rel_base_y|`
- `Total_base_rel_max_m = max_t sqrt(u_rel_base_x^2 + u_rel_base_y^2)`

Ayrica input-proxy referansi uretilir:

- `Input Motion` ivmesi -> baseline correction -> cift integrasyon -> `u_input_proxy`
- `u_rel_input = u_rel_base - u_input_proxy`

ve onun max kolonlari yazilir.

## 2) Legacy_Methods

Iki kaynaktan ozet cikarilir:

- `Profile` sheet max deplasmanlari:
  - `Profile_X_max_m`, `Profile_Y_max_m`
  - `Profile_RSS_total_m = sqrt(Profile_X_max_m^2 + Profile_Y_max_m^2)`
- Layer ivme serileri:
  - Her layer icin `Acceleration (g)` alinip baseline correction + cift integrasyonla `dx(t), dy(t)`
  - `TimeHist_Resultant_total_m = max_t sqrt(dx^2 + dy^2)`

## 3) Comparison

`Strain_Relative` ve `Legacy_Methods` birlestirilir.
Asagidaki kolonlar olusur:

- Base-vs-profile farklari
- Base-vs-timehistory farklari
- Base-corrected profile kolonlari:
  - `Profile_X_minus_bottom_m`
  - `Profile_Y_minus_bottom_m`
  - `Profile_RSS_minus_bottom_m`
- X/Y hizalama deltasi:
  - `Delta_Xbase_vs_ProfileXminusbottom_m`
  - `Delta_Ybase_vs_ProfileYminusbottom_m`

## Baseline Filtering Var mi?

Evet. `_baseline_correct(...)` icinde ivme serisine 3. derece polinom trend fit edilip cikariliyor.
Ardindan hiz ve deplasman, trapez integrasyon (`_cumtrapz`) ile iki kez entegrasyonla uretiliyor.
Bu akiş `legacy` ve `input-proxy` yollarinda aktif.

## Uretilen Workbook (output_total_*.xlsx)

Her X/Y cifti icin 1 workbook uretilir.
Sheet'ler:

1. `Strain_Relative`
- Strain tabanli base-relative ve input-proxy-relative max kolonlari

2. `Legacy_Methods`
- Profile RSS ve ivmeden cikan legacy zaman-gecmis ozetleri

3. `Comparison`
- Yontemler arasi fark/ratio ve base-corrected kolonlar

4. `Depth_Profiles`
- Derinlige bagli toplu profile karsilastirma tablosu

5. `Profile_BaseCorrected`
- X/Y bazinda taban-duzeltilmis profile hizalama tablosu

6. `Direction_X_Time`
- Tum layerlar icin tek sheette yanyana signed X deplasman-zaman kolonlari

7. `Direction_Y_Time`
- Tum layerlar icin tek sheette yanyana signed Y deplasman-zaman kolonlari

8. `Resultant_Time`
- Tum layerlar icin `sqrt(X^2 + Y^2)` zaman serileri (pozitif)

## Workbook Icindeki Grafikler

- `Depth_Profiles`: derinlige bagli 4 profilin tek grafikte karsilastirmasi
- `Profile_BaseCorrected`: X ve Y icin base-relative vs base-corrected profile grafik ciftleri
- `Direction_X_Time`, `Direction_Y_Time`, `Resultant_Time`: layer bazli toplu zaman-serisi grafikleri
- Eksenler ve tick etiketleri acik olarak gosterilecek sekilde konfigure edilir.

## Opsiyonel Rapor Dosyalari

CLI `--with-report` ile workbook yanina su dosyalar eklenir:

- `*_alignment_report.md`
- `*_base_corrected_profiles.png`
- `*_alignment_deltas.png`

Bu rapor, strain base-relative ve DEEPSOIL profile (base-corrected) uyumunu sayisal ve gorsel ozetler.

## CLI Kullanim

Proje klasorunde:

```powershell
python GetDisp4.py --input-dir .
```

Opsiyonlar:

- `--output-dir <path>`: cikti klasoru
- `--include-manip`: `*-manip.xlsx` dosyalarini da eslemeye kat
- `--fail-fast`: ilk hatada dur
- `--with-report`: alignment markdown + png raporlarini da uret

Ornek:

```powershell
python GetDisp4.py --input-dir . --output-dir . --with-report
```

## Web UI (Pyodide / tam client-side)

```powershell
cd "H:\Drive'im\Ortak\Bildiri_Makale\Gapping-Nongapping\Deepsoil"
python -m http.server 8010
```

Ac:

`http://localhost:8010/web-ui/`

UI ozellikleri:

- Folder select (`webkitdirectory`) + direkt dosya secimi
- Secilen dosya listesi ve pair sayaci
- Batch run
- Tekil cikti indir + toplu ZIP indir

## Klasor Yapisi

- `disp_core.py`: ortak hesap cekirdegi
- `GetDisp4.py`: CLI giris noktasi
- `report_alignment.py`: rapor ve plot uretici
- `web-ui/`: Pyodide arayuz

## Bilinen Yorum Notu

`Profile` tablosundaki max deplasmanlar, taban ofseti nedeniyle dogrudan strain base-relative ile birebir oturmayabilir.
Bu nedenle `Comparison` icinde base-corrected kolonlar ayrica verilir.

## Lisans / Surumleme

Bu repo su an proje-ici arastirma amacli duzende tutuluyor.
Ihtiyac olursa `LICENSE` ve semantik versiyonlama (`v0.x`) eklenebilir.
