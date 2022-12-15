<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("sinan.txt")
set dosyanesne = fso.opentextfile(dosyaadi,1)
'okunanbilgi=dosyanesne.readall
okunanbilgi=dosyanesne.read(10)
response.Write(okunanbilgi)
okunanbilgi=dosyanesne.readline
okunanbilgi=dosyanesne.readline
okunanbilgi=dosyanesne.readline
response.Write(okunanbilgi)

dosyanesne.close
set fso=nothing
%>
