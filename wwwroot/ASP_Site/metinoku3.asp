<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("sinan.txt")
set dosyanesne = fso.opentextfile(dosyaadi,1)
dosyanesne.skipline
dosyanesne.skipline
while not dosyanesne.atendofstream
response.write dosyanesne.readline

wend

dosyanesne.close
set fso=nothing
%>
