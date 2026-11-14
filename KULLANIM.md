# İş Merkezi

İş takibi, fatura görüntüsü, ortak klasör ve isteğe bağlı e-posta tek pencerede.

## Açılış
Masaüstündeki **İş Merkezi** uygulamasına çift tıkla. Pencere doğrudan açılır.

## Ekran
Soldaki sütunlar bekleyen, devam eden ve tamamlanan işleri gösterir. Kartın üzerine basınca kayıt açılır. Sağda son faturalar ve fatura görüntüsü durur. Görsele basınca büyük açılır; Esc ile kapanır.

Faturalar bölümünde liste solda, görüntü sağdadır. Tümü, Ödenmedi, Klasör ve E-posta süzgeçleri vardır.

## Ortak klasör ve Excel
Varsayılan klasör: `~/Desktop/Faturalar`  
Varsayılan liste: `Fatura_Takip_Liste.xlsx`

Klasöre bırakılan PDF ve görüntüler uygulamaya eklenir. Listede olmayan dosyalar Excel'e de satır olarak yazılır. Tutar dosyadan güvenilir okunamazsa satır **Kontrol Gerekli** kalır; Excel'de zaten yazılı tutarlar uygulamaya gelir.

İlk Excel yazımından önce aynı klasörde `Fatura_Takip_Liste.onceki.xlsx` yedeği oluşur. Uygulamadan silinen kayıt Excel satırını silmez. Klasördeki dosya duruyorsa ve değiştirilmediyse yeniden eklenmez.

## E-posta
Kapalı başlar. Ayarlar'dan sunucu, kullanıcı ve parolayı kaydedince açılır. Bağlantı her zaman şifreli IMAP'dir; posta kutusu salt okunur açılır. Son 21 gündeki PDF ve görüntü ekleri alınır, uygulamaya ve Excel'e yazılır. Parola veritabanında durmaz, macOS Anahtar Zinciri'nde tutulur.

## Güvenlik
- Yalnızca PDF ve görüntü kabul edilir. İçerik, uzantıyla uyuşmuyorsa alınmaz.
- Ekler çalıştırılmaz.
- İzlenen klasör ev dizininin içinde ayrı bir klasör olmalıdır. Masaüstünün tamamı seçilemez.
- Excel dosyası bu klasörün içinde kalır.
- Uygulama verisi yalnız bu Mac'te, `~/Library/Application Support/IsMerkezi/` altındadır.

## Diğer kayıtlar
Görev, hatırlatma, firma, gelir-gider ve proje kayıtları duruyor. Tarih biçimi `YYYY-AA-GG`, saat `HH:MM`. Hatırlatma bildirimi uygulama açıkken gelir.
