Sub Sınav()
'
' Sınav Makro, Zamanusta tarafından 14.12.2017 tarihinde oluşturuldu.
'
'
Randomize Timer
n = 55 'Öğrenci sayısı.


'Kaç soru sormayı planlıyorsanız aşağıda o kadar string array yaratmanız gerekiyor. Her bir array, o soru için sorabileceğiniz tüm alternatif soruları içermelidir.
arsoru1 = Array("Soru1-1.alternatif", "Soru1-2.alternatif", "Soru1-3.alternatif", "Soru1-4.alternatif")
arsoru1 = Array("Soru2-1.alternatif", "Soru2-2.alternatif", "Soru2-3.alternatif", "Soru2-4.alternatif")
arsoru1 = Array("Soru3-1.alternatif", "Soru3-2.alternatif")
arsoru1 = Array("Soru4-1.alternatif", "Soru4-2.alternatif", "Soru4-3.alternatif")

For i = 1 To n
Selection.Font.Bold = True
head = "20XX-20XX Y Dersi Sınavı  -  Kağıt No:  #" + Str(i) + "               Öğrenci Adı ve Soyadı:"
cizgi = "----------------------------------------------------------------------------------------------------------------------------"
Selection.TypeText Text:=head
Selection.InsertBreak Type:=wdLineBreak
Selection.TypeText Text:=cizgi
Selection.InsertBreak Type:=wdLineBreak
Selection.InsertBreak Type:=wdLineBreak
Selection.Font.Bold = False

'SORULAR İÇİN RANDOM SAYILAR
s1 = Int(Rnd * 4)
s2 = Int(Rnd * 4)
s3 = Int(Rnd * 2)
s4 = Int(Rnd * 3)

'SORULARIN OLUŞTURULMASI

soru1 = "1) " + arsoru1(s1) + " (20 puan)"
soru2 = "2) " + arsoru2(s2) + " (30 puan)"
soru3 = "3) " + arsoru3(s3) + " (20 puan)"
soru4 = "4) " + arsoru3(s4) + " (20 puan)"

'YAZDIRIYORUZ
Selection.TypeText Text:=soru1
Selection.InsertBreak Type:=wdLineBreak
Selection.InsertBreak Type:=wdLineBreak

Selection.TypeText Text:=soru2
Selection.InsertBreak Type:=wdLineBreak
Selection.InsertBreak Type:=wdLineBreak

Selection.TypeText Text:=soru3
Selection.InsertBreak Type:=wdLineBreak
Selection.InsertBreak Type:=wdLineBreak

Selection.TypeText Text:=soru4
Selection.InsertBreak Type:=wdLineBreak
Selection.InsertBreak Type:=wdLineBreak

Selection.TypeText Text:="Başarılar... Süre 40 dakika."
Selection.InsertBreak Type:=wdLineBreak
Selection.TypeText Text:=cizgi



Selection.InsertBreak Type:=wdPageBreak
Next



End Sub
