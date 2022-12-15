<% 
Set sozluknesnesi = CreateObject("Scripting.Dictionary")
sozluknesnesi.comparemode = 1 ' 1: Büyük-küçük harf ayýrmaz
sozluknesnesi.comparemode = 0 ' 0: Büyük-küçük harf ayýrýr
sozluknesnesi.add "TR", "Türkiye Cumhuriyeti"
sozluknesnesi.add "USA", "Amerika Birleþik Devletleri"
sozluknesnesi.add "ENG", "Ýngiltere"
sozluknesnesi.add "eng", "Ýngiltere"

for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next
response.write "<br>"

sozluknesnesi.remove("USA")
for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next
response.write "<br>"

sozluknesnesi.removeall
response.Write(sozluknesnesi.count)
for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next
response.write "<br>"

%>

