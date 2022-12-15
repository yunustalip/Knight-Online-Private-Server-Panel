<% 
Set sozluknesnesi = CreateObject("Scripting.Dictionary")
sozluknesnesi.add "TR", "Türkiye Cumhuriyeti"
sozluknesnesi.add "USA", "Amerika Birleþik Devletleri"
sozluknesnesi.add "ENG", "Ýngiltere"

'response.write sozluknesnesi.item("TR") & "<br>"
'response.write sozluknesnesi.item("ENG") & "<br>"
'response.write sozluknesnesi.item("USA") & "<br>"

sozluknesnesi.item("ENG")="Britanya"
sozluknesnesi.key("ENG")="ÝNG"
sozluknesnesi.item("JPN")="Japonya"

for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next

response.write "sözlükte " & sozluknesnesi.count & " adet sözcük var<br><br>"

for sayac=0 to sozluknesnesi.count-1
anahtarlar = sozluknesnesi.keys
degerler = sozluknesnesi.items
response.write anahtarlar(sayac)  & "-" & degerler(sayac) & "<br>"
next
%>

