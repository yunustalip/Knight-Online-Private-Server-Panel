<% 
Set sozluknesnesi = CreateObject("Scripting.Dictionary")
sozluknesnesi.add "TR", "T�rkiye Cumhuriyeti"
sozluknesnesi.add "USA", "Amerika Birle�ik Devletleri"
sozluknesnesi.add "ENG", "�ngiltere"

'response.write sozluknesnesi.item("TR") & "<br>"
'response.write sozluknesnesi.item("ENG") & "<br>"
'response.write sozluknesnesi.item("USA") & "<br>"

sozluknesnesi.item("ENG")="Britanya"
sozluknesnesi.key("ENG")="�NG"
sozluknesnesi.item("JPN")="Japonya"

for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next

response.write "s�zl�kte " & sozluknesnesi.count & " adet s�zc�k var<br><br>"

for sayac=0 to sozluknesnesi.count-1
anahtarlar = sozluknesnesi.keys
degerler = sozluknesnesi.items
response.write anahtarlar(sayac)  & "-" & degerler(sayac) & "<br>"
next
%>

