<%
'dim mevsim
'redim mevsim(3)
'mevsim(0)="Ýlkbahar"
'mevsim(1)="Yaz"
'mevsim(2)="Sonbahar"
'mevsim(3)="Kýþ"
'mevsim=array("Ýlkbahar","Yaz","Sonbahar","Kýþ")

'yenidizi=filter(mevsim, "bahar")
metin = "Ben seni sen de beni tanýyoruz."
metin1 = "Ýlkbahar|Yaz|Sonbahar|Kýþ"
metin2 = "10.00$15.25$30.55$31.67"

yenidizi=split(metin2,"$",2)

for each deger in yenidizi
%>*<%response.Write(deger) & "*<br>"
next

mevsim=array("Ýlkbahar","Yaz","Sonbahar","Kýþ")
for each deger in mevsim
%>*<%response.Write(deger) & "*<br>"
next
metin1 = "Ýlkbahar$Yaz$Sonbahar$Kýþ"
yenidizi=split(metin1,"$")
yenimetin=join(yenidizi,"#")
'response.write(yenimetin)
response.write( replace("Ýlkbahar$Yaz$Sonbahar$Kýþ","$","*"))

%>
