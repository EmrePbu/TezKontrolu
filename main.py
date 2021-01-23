import methods as myMethods

# bu kısım çalışmıyor
#        word dosyasının kesinlikle ./documents klasöründe olması gerekiyor
#        fileName = "5150rnek_Tez_O_Orhan_YL_Parametresiz"


# Word dosyasında var olan bütün resimleri belirtilen dosya konumuna json dosyası olarak alır.
# ./buffer/word/media/ dosya yolunu kontrol edebilirsiniz.
myMethods.GetAllImages()

# Word dosyasından kullanılan bütün yazı fontlarını belitrilen dosya konumuna json dosyası olarak alır.
# ./buffer/word/fontTable.xml.json/ dosya yolunu kontrol edebilirsiniz.
myMethods.GetAllFonts()

# Word dosyasının içeriğinin bulunduğu kısım body kısmıdır bunuda belirtilen dosya konumuna json dosyası olarak alır.
# ./buffer/word/document.xml.json/ dosya yolunu kontrol edebilirsiniz.
myMethods.GetBody()

# Word dosyasının sayfa sayısını verir.
pages = myMethods.GetPageNumber()
print("Toplam sayfa sayısı: ", pages)

# Word dosyasındaki bütün sayfaların kenar boşluklarını kontrol eder.
# Hangi sayfaların Kenar boşluğu uygun değilse onu yazar. uygun ise True yazar.
myMethods.GetPagesMargin()
