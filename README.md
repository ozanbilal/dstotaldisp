# Deepsoil Total Disp Calculator

DEEPSOIL sonuc dosyalarindan toplam ve goreceli yerdegistirme profillerini ureten arac.
Hem CLI (`GetDisp4.py`) hem de tarayici tabanli WASM arayuz (`web-ui/`) ayni hesap cekirdegini (`disp_core.py`) kullanir.

## Ne Cozer?

Bu proje, su iki dunyayi ayni cikti setinde birlestirir:

- DEEPSOIL tarafi (goreceli): `u_rel(z,t) = u(z,t) - u(base,t)`
- TBDY yorumu (toplam): `u_total(z,t) = u(base,t) + u_rel(z,t)`

Not: Mevcut cekirdek, strain tabanli yolakta dogrudan `*_base_rel_*` (base-relative) uretir.
`u(base)` proxy serisi (`Input Motion` entegrasyonu) ile `u_total = u_base + u_rel` explicit olarak da uretilir.

## Giris Beklentisi

- Dosya tipi: `.xlsx`
- Pair modu esleme kurali:
  - X dosyasi: adinda `_X_` ve sonda `_H1.xlsx`
  - Y dosyasi: ayni adin `_X_ -> _Y_`, `_H1 -> _H2` donusmus hali
- Single-file modu:
  - Eslestirilemeyen her uygun `.xlsx` dosya tek basina islenir.
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

### `X_base_rel_max_m` nasil hesaplanir?

Her katman `i` icin su adimlarla uretilir:

1. `Layer i` sheet'inden X yonu strain serisi okunur (`Strain (%)` -> `gamma_i(t)`).
2. Tum katmanlar icin ortak zaman dizisi kurulur:
   - ortak pencere: katmanlarin kesisen zaman araligi
   - `dt`: katman zaman adimlarinin medyanlarindan en kucugu
   - her katman strain'i bu ortak zamana lineer enterpole edilir
3. Katman katkisi hesaplanir:
   - `du_i(t) = gamma_i(t) * h_i`
4. Taban-referansli kümülatif deplasman bulunur:
   - `u_rel_base_x(i,t) = sum_{k=i..N} du_k(t)` (en alttan yukariya birikimli toplam)
5. Son kolon degeri:
   - `X_base_rel_max_m(i) = max_t |u_rel_base_x(i,t)|`

Yani `X_base_rel_max_m`, her katman icin strain tabanli goreceli deplasman zaman serisinin mutlak maksimum zarfi.

Sonra zarf degerler:

- `X_base_rel_max_m = max_t |u_rel_base_x|`
- `Y_base_rel_max_m = max_t |u_rel_base_y|`
- `Total_base_rel_max_m = max_t sqrt(u_rel_base_x^2 + u_rel_base_y^2)`
- `X_tbdy_total_max_m = max_t |u_base_ref_x + u_rel_base_x|`
- `Y_tbdy_total_max_m = max_t |u_base_ref_y + u_rel_base_y|`
- `Total_tbdy_total_max_m = max_t sqrt((u_base_ref_x + u_rel_base_x)^2 + (u_base_ref_y + u_rel_base_y)^2)`

Ek olarak DEEPSOIL `Profile` taban ofsetine hizali bir tahmin seti yazilir:

- `X_profile_offset_total_est_m = (X_base_rel_max_m - X_base_rel_max_m(bottom)) + Profile_X_max_m(bottom)`
- `Y_profile_offset_total_est_m = (Y_base_rel_max_m - Y_base_rel_max_m(bottom)) + Profile_Y_max_m(bottom)`
- `Total_profile_offset_total_est_m = sqrt(X_profile_offset_total_est_m^2 + Y_profile_offset_total_est_m^2)`

Bu kolonlarin alt satiri DEEPSOIL profile taban degerine esit olur (ornegin ~`0.0145 m`).

`u_base_ref` secimi opsiyoneldir:

- `input` (varsayilan): `Input Motion` ivmesinin entegrasyonundan gelen proxy
- `deepest_layer`: en alt layer ivmesinin entegrasyonundan gelen proxy

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
- TBDY toplam karsilastirmasi:
  - `Delta_tbdy_vs_profile_m`
  - `Ratio_tbdy_to_profile`
- Profile-bottom-ofset tahmin karsilastirmasi (varsa):
  - `Delta_profileoffset_vs_profile_m`
  - `Ratio_profileoffset_to_profile`

## Baseline Filtering Var mi?

Evet.

- Varsayilan davranis: baseline ve filtering **kapali** (ham ivme entegrasyonu).
- Geriye uyumlu modda (legacy high-pass alanlari verilirse): 3. derece baseline + yumusak FFT high-pass (`cutoff=0.03 Hz`, `transition=0.02 Hz`) kullanilir.
- Yeni gelismis modda: `Processing Order`, `Baseline Method`, `Filter Domain`, `Filter Config`, `Filter Type`, `F Low`, `F High`, `Order` parametreleriyle akis ayarlanir.
- Integrasyon her durumda `_cumtrapz` ile yapilir.
- Opsiyonel karsilastirma modunda (`integrationCompareEnabled=true`) ayni on-islenmis ivmeden ikinci bir alternatif uretilir:
  - `fft_regularized` (frekans alaninda `U(f) = -A(f) * H_hp(f) / (2*pi*f)^2`)
  - alt low-cut kurali: filtering aciksa `F Low`, kapaliysa `0.05 Hz`
  - primary korunur, alt ve `ALT - primary` farklari ek kolon/sheet olarak yazilir.

## Uretilen Workbook (Pair Modu, output_total_*.xlsx)

Her X/Y cifti icin 1 workbook uretilir.
Sheet'ler:

1. `Strain_Relative`
- Strain tabanli base-relative, TBDY total (`u_base + u_rel`) ve input-proxy-relative max kolonlari

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

9. `TBDY_Total_X_Time`
- Tum layerlar icin `X_total(t) = X_base_proxy(t) + X_relative(t)` serileri

10. `TBDY_Total_Y_Time`
- Tum layerlar icin `Y_total(t) = Y_base_proxy(t) + Y_relative(t)` serileri

11. `TBDY_Total_Resultant_Time`
- Tum layerlar icin `sqrt(X_total^2 + Y_total^2)` serileri

Compare aciksa ek sheet'ler:

- `Direction_X_Time_ALT`
- `Direction_Y_Time_ALT`
- `Resultant_Time_ALT`
- `TBDY_Total_X_Time_ALT`
- `TBDY_Total_Y_Time_ALT`
- `TBDY_Total_Resultant_Time_ALT`

## Uretilen Workbook (Single-File Modu, output_single_*.xlsx)

Eger dosya X/Y ciftine eslesmiyorsa tek basina islenir ve su sheet'ler uretilir:

1. `Single_Direction_Summary`
- Katman bazli `Base_rel_max_m`, `TBDY_total_max_m`, `Input_proxy_rel_max_m`, `Profile_max_m`, `TimeHist_maxabs_m`

2. `Direction_Time`
- O dosyanin yonundeki (X veya Y) tum layer signed deplasman-zaman serileri

3. `Strain_Relative_Time`
- Strain tabanli base-relative signed zaman serileri

4. `TBDY_Total_Time`
- Explicit `u_base + u_rel` signed zaman serileri

5. `InputProxy_Relative_Time`
- `u_rel - u_input_proxy` signed zaman serileri

Compare aciksa ek sheet'ler:

- `Direction_Time_ALT`
- `Strain_Relative_Time_ALT`
- `TBDY_Total_Time_ALT`
- `InputProxy_Relative_Time_ALT`

## Method-2 Ek Ciktisi (Dosya Bazli, output_method2_*.xlsx)

Her uygun giris `.xlsx` dosyasi icin ayri uretilir.

- X dosyalari: `Method2_TBDY_X_Time`
- Y dosyalari: `Method2_TBDY_Y_Time`
- Ilk sutun: `Time_s`
- Diger sutunlar: katman bazli TBDY toplam deplasman zaman serileri (`u_base + u_rel`)

Compare aciksa ayni dosyada ek sheet:

- `Method2_TBDY_X_Time_ALT` / `Method2_TBDY_Y_Time_ALT`
- `Method2_TBDY_X_Delta` / `Method2_TBDY_Y_Delta` (`ALT - primary`)

## Method-3 Ek Ciktisi (Toplu, output_method3_profiles_all.xlsx)

Tek bir toplu dosyada tum kayitlarin katman bazli maksimum profil degerleri verilir.

- `Method3_Profile_X`: `Depth_m` + her X kayit icin `max(|u(t)|)` kolonu
- `Method3_Profile_Y`: `Depth_m` + her Y kayit icin `max(|u(t)|)` kolonu
- Derinlikler dis birlestirme (`outer`) ile hizalanir.

Compare aciksa ek sheet:

- `Method3_Profile_X_ALT`
- `Method3_Profile_Y_ALT`
- `Method3_Delta_X` (`ALT - primary`)
- `Method3_Delta_Y` (`ALT - primary`)

Method-3, dosya bazinda su seriden uretilir:

- `u_tbdy_total(t) = u_rel_base(t) + u_base_ref(t)`
- `u_rel_base`: strain tabanli `sum(gamma*h)` birikimi
- `u_base_ref`: secime gore `Input Motion` veya `deepest layer` ivmesinin cift integrasyonundan
- Method-3 profili: her derinlikte `max_t |u_tbdy_total(z,t)|`

Sonra tum X dosyalari tek tabloda (`Method3_Profile_X`), tum Y dosyalari tek tabloda (`Method3_Profile_Y`) birlestirilir.

## Direction vs TBDY Total Farki (Grafikler Neden Farkli?)

`Direction_X_Time` / `Direction_Y_Time` ile `TBDY_Total_X_Time` / `TBDY_Total_Y_Time` ayni buyukluk degildir:

- `Direction_*_Time`:
  - her layerin kendi `Acceleration (g)` serisi cift integre edilir
  - yani "layer acceleration -> displacement" yolagi
- `TBDY_Total_*_Time`:
  - strain'den gelen goreceli deplasman (`u_rel_base = sum(gamma*h)`)
  - buna `Input Motion`dan gelen taban proxy deplasmani eklenir (`u_input_proxy`)
  - yani `u_total = u_rel_base + u_input_proxy`

Bu nedenle genlik/faz ve katmanlar arasi ayrisma dogal olarak farkli gorunur.
Ozellikle `TBDY_Total_*` grafiklerinde cizgilerin birbirine yakin olmasi normaldir; tum katmanlara ortak `u_input_proxy` bileseni eklenir.

Onemli: karsilastirmayi ayni yon icin yapin (`Direction_X_Time` vs `TBDY_Total_X_Time` veya `Direction_Y_Time` vs `TBDY_Total_Y_Time`).
X ile Y grafigini capraz karsilastirmak (or. `TBDY Total X` vs `Direction Y`) dogrudan uyum vermez.

## Workbook Icindeki Grafikler

- `Depth_Profiles`: derinlige bagli profil karsilastirmasi
  - `Direction_X_*` / `Direction_Y_*` signed `(+max, -min)` ve varsa `_ALT (+max, -min)` kolonlari
  - resultants toplami (`sqrt(x^2+y^2)`) CLI/UI opsiyonu ile acilip kapanabilir
- `Profile_BaseCorrected`: X ve Y icin base-relative vs base-corrected profile grafik ciftleri
- `Direction_X_Time`, `Direction_Y_Time`, `Resultant_Time`: layer bazli toplu zaman-serisi grafikleri
- `TBDY_Total_X_Time`, `TBDY_Total_Y_Time`, `TBDY_Total_Resultant_Time`: explicit TBDY total zaman-serisi grafikleri
- `Direction_Time`, `Strain_Relative_Time`, `TBDY_Total_Time`, `InputProxy_Relative_Time`: single-file grafik seti
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
- `--no-method23`: Method-2 ve Method-3 ek ciktilarini devre disi birak
- `--no-method2`: sadece Method-2 ciktilarini devre disi birak
- `--no-method3`: sadece Method-3 ciktisini devre disi birak
- `--baseline-on`: baseline duzeltmeyi ac
- `--filter-on`: filtering'i ac
- `--base-reference {input,deepest_layer}`: TBDY total icin base deplasman referansi secimi
- `--integration-compare`: FFT-regularized alt entegrasyon karsilastirmasini ac
- `--hide-resultant-profiles`: `Depth_Profiles` sheet/chart icinde resultant/toplam serileri gizle
- `--alt-integration-method {fft_regularized}`: alt entegrasyon yontemi (su an tek secenek)

Ornek:

```powershell
python GetDisp4.py --input-dir . --output-dir . --with-report
```

## Web UI (Pyodide / tam client-side)

```powershell
cd "<path-to-dstotaldisp>"
python -m http.server 8010
```

Ac:

`http://localhost:8010/web-ui/`

UI ozellikleri:

- Folder select (`webkitdirectory`) + direkt dosya secimi
- Secilen dosya listesi ve pair sayaci
- `Method-2 outputs` ve `Method-3 outputs` ayri ayri secilebilir
- `Compare with FFT-Regularized integration` (varsayilan kapali)
- `Include resultant (RSS) totals in Depth_Profiles` (varsayilan kapali)
- Processing paneli:
  - `Apply Baseline` (default kapali)
  - `Apply Filtering` (default kapali)
  - `Processing Order`
  - `Filter Domain`
  - `Baseline Method`
  - `Filter Config`
  - `Filter Type`
  - `Base Reference` (`Input Motion Proxy` / `Deepest Layer Proxy`)
  - `F Low`, `F High`, `Order`
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
