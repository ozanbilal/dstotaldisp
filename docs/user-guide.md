# DeepSoil Toplam Deplasman Kullanici Kilavuzu

Son guncelleme: 2026-04-10

## Arac ne cozer

Bu arac, DeepSoil ciktilarindan toplam deplasman profillerini tek arayuzde toplar, gecerli hesap yollarini ayni grafikte ust uste gosterir ve uretilen workbook/ZIP ciktilarinin indirilmesini saglar.

Ana hedef:

- goreceli deplasman ureten DeepSoil exportlarini toplam deplasman yorumuna tasimak
- dosyada hangi veri varsa ona gore gecerli yontemleri secmek
- kullaniciyi gereksiz sheet/kolon detayina bogmadan sonucu gostermek
- gerekirse detayli source explorer ile ham egri seviyesine inmek

## Desteklenen girisler

- `.xlsx` DeepSoil export workbook
- `.db` / `.db3` DeepSoil `deepsoilout` veritabani

Bugun aktif kaynak sistemi DeepSoil'dir. RSSeismic icin adapter siniri hazirdir ancak bu surumde dogrudan parser bagli degildir.

## Hangi veriyle ne uretilir

- Sadece `.xlsx` ve tum layer sheet'leri varsa:
  - strain + input proxy toplam profil
  - strain + deepest layer proxy toplam profil
  - profile offset yaklasik toplam profil
  - time-history dolayli maksimum profil
  - profile referans profili
- `.xlsx` var ama layer seti eksikse:
  - strain tabanli ve time-history tabanli profiller gizlenir
  - elde kalaniyla `ubase + urel`/profile offset yaklasik profil gosterilir
  - arayuz bu durumu "sinirli veri" olarak isaretler
- `.db/.db3` verisi varsa ve toplam deplasman tablolarina erisilebiliyorsa:
  - DB direct total birincil yontem olur
  - toplam deplasman profili dogrudan veritabanindan okunur

## Pair, single ve DB direct akislar

- Pair:
  - X ve Y kayitlari eslesen girdiler birlikte islenir
  - ana ozet grafikte resultant veya cift-yonlu toplam profil one cikar
- Single:
  - eslesmeyen tek yon kayitlari tekil profil olarak islenir
  - varyant etiketleri yon bilgisini korur
- DB direct:
  - strain veya ivme entegrasyonu yerine veritabanindaki displacement kolonlari okunur
  - filter, baseline ve base reference kontrolleri devre disi kalir

## Arayuz akisi

- Ana yuzey summary-first tasarlanmistir:
  - once `Toplam deplasman profili`
  - hemen altinda `Inputs & artifacts`
  - en sonda varsayilan olarak kapali `Detayli inceleme` drawer
- Mobil yerlesimde run paneli summary bolumunun hemen altina gelir.
- `Detay` dugmesi veya `Detayli inceleme` baglantisi explorer drawer'ini acar ve ilgili source kaydina atlar.
- `Gorunen varyant` secimi ile ayni kayit icindeki gecerli overlay vurgusu degistirilebilir.
- `Spectrum max period` yalniz o anda acik olan spectrum chart'ini sinirlar; diger chart ve ekranlara tasinmaz.
- Paylasilabilir URL yalniz temel sahne durumunu tasir:
  - `shell`
  - `summary`
  - `source`
  - `family`
  - `chart`
  - `layer`
- `Viewer sources` sayaci dosya adedini degil, detay explorer'a giren source catalog kayitlarini sayar. Bir pair isleminde ayni iki dosya icin genelde `X`, `Y` ve ortak `pair` olmak uzere birden fazla viewer source olusur.

## Yontemler

Arayuz gecerli varyantlari sabit bir guven sirasi ile siralar:

1. `DB direct total`
2. `strain + input proxy total`
3. `strain + deepest-layer proxy total`
4. `profile-offset / approx total`
5. `time-history indirect total`
6. `profile raw reference`

Yorumlama kurali:

- en yuksek guvenli gecerli varyant ana egri olarak vurgulanir
- diger gecerli varyantlar ayni grafikte destekleyici overlay olarak kalir
- gecersiz varyantlar grafikte cizilmez, nedenleri altta listelenir

## Senaryolar

### Eksik layer bulunan `.xlsx`

- Tum katmanlar yoksa her sey hesaplanamaz.
- Bu durumda strain ve dolayli time-history profilleri gizlenebilir.
- Yine de profile offset veya `ubase + urel` tabanli yaklasik profil gosterilebilir.

### Yalniz `.xlsx`

- En yaygin akis budur.
- Dosyadaki katman kapsamina gore dogrudan, yaklasik ve dolayli yontemler birlikte overlay edilir.

### `.db/.db3` mevcut

- Analizde toplam deplasman grafikleri veritabanina yazildiysa en guvenilir kaynak burasidir.
- Arayuz DB direct varyantini birincil secim olarak isaretler.

### Manual pairing

- Otomatik X/Y eslemesi yeterli degilse legacy panelinden elle esleme yapilabilir.
- Pair olustuktan sonra ozet kaydi buna gore yenilenir.

## Grafik nasil okunur

- x ekseni deplasmani, y ekseni derinligi gosterir
- derinlik ekseni ters cevrilir; ustte zemin yuze yi, altta daha derin katmanlar bulunur
- kalin ve baskin cizgi birincil yontemdir
- daha soluk veya kesikli cizgiler destekleyici varyantlardir
- "sinirli veri" etiketi, mevcut dosyanin tum yontemleri desteklemedigini gosterir

## Indirme ve ciktilar

Indirme kapsaminda mevcut workbook ve ZIP ciktilari korunur:

- pair output workbook'lari
- single output workbook'lari
- method-2 / method-3 workbook'lari
- DB direct workbook'lari
- toplu ZIP indirmesi

Bu surumde yeni CSV veya PNG export eklenmemistir.

## Sinirlar ve dikkat notlari

- `.xlsx` icinde eksik layer varsa tam strain tabanli yorum yapilamaz
- DB direct yalniz veritabaninda ilgili displacement kolonlari varsa calisir
- profile referansi ile strain tabanli toplam profil ayni kavram degildir; ayni grafikte karsilastirma amaclidir
- base reference secimi legacy davranisi etkiler ancak summary arayuzu veri varsa birden fazla proxy varyantini birlikte gosterebilir

## Sorun giderme

- Grafik bos gorunuyorsa:
  - dosya seciminin dogru oldugunu
  - desteklenen uzanti kullanildigini
  - ilgili varyantin gecerli veri uretebildigini kontrol et
- Yalniz tek bir yontem gorunuyorsa:
  - giriste yeterli layer veya DB toplam deplasman verisi olmayabilir
- Pair beklenirken single cikiyorsa:
  - dosya adlarini veya manual pairing ayarini kontrol et
- DB direct secenegi pasif gorunuyorsa:
  - secilen girdilerin `.db`/`.db3` oldugunu dogrula

## Daha teknik bilgi

- UI dokumantasyonu: `web-ui/docs.html`
- Teknik repo aciklamasi: `README.md`
- Ham workbook ve hesap cekirdegi ayrintilari: `disp_core.py`
