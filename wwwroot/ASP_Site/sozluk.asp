<% 
Set sozluknesnesi = CreateObject("Scripting.Dictionary")
sozluknesnesi.comparemode = 1 ' 1: B�y�k-k���k harf ay�rmaz
sozluknesnesi.comparemode = 0 ' 0: B�y�k-k���k harf ay�r�r
sozluknesnesi.add "TR", "T�rkiye Cumhuriyeti"
sozluknesnesi.add "USA", "Amerika Birle�ik Devletleri"
sozluknesnesi.add "ENG", "�ngiltere"
sozluknesnesi.add "eng", "�ngiltere"

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

