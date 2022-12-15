<%
Set FSO = CreateObject("Scripting.FileSystemObject")
Set sozluknesnesi = CreateObject("Scripting.Dictionary")
dosyaadi=server.mappath("Turkcesi_varken.txt")
set dosyanesne = fso.opentextfile(dosyaadi,1)

while not dosyanesne.atendofstream
okunankelimeler= dosyanesne.readline
kelimeler = split(okunankelimeler, "|")
sozluknesnesi.add kelimeler(0), kelimeler(1)
wend

for each anahtar in sozluknesnesi.keys
response.write anahtar  & "-" & sozluknesnesi.item(anahtar) & "<br>"
next
response.write "<br>"


dosyanesne.close
set sozluknesnesi=nothing
set fso=nothing
%> 