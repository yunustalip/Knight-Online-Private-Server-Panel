<%
'dim mevsim
'redim mevsim(3)
'mevsim(0)="�lkbahar"
'mevsim(1)="Yaz"
'mevsim(2)="Sonbahar"
'mevsim(3)="K��"
'mevsim=array("�lkbahar","Yaz","Sonbahar","K��")

'yenidizi=filter(mevsim, "bahar")
metin = "Ben seni sen de beni tan�yoruz."
metin1 = "�lkbahar|Yaz|Sonbahar|K��"
metin2 = "10.00$15.25$30.55$31.67"

yenidizi=split(metin2,"$",2)

for each deger in yenidizi
%>*<%response.Write(deger) & "*<br>"
next

mevsim=array("�lkbahar","Yaz","Sonbahar","K��")
for each deger in mevsim
%>*<%response.Write(deger) & "*<br>"
next
metin1 = "�lkbahar$Yaz$Sonbahar$K��"
yenidizi=split(metin1,"$")
yenimetin=join(yenidizi,"#")
'response.write(yenimetin)
response.write( replace("�lkbahar$Yaz$Sonbahar$K��","$","*"))

%>
