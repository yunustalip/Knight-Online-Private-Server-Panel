<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("sinan.txt")
set dosyanesne = fso.opentextfile(dosyaadi,1)
dosyanesne.skipline
okunanbilgi=dosyanesne.read(10)
response.write "�u an " & dosyanesne.line & ".sat�r " & dosyanesne.column & ". s�tundas�n�z <br>"
response.write "<br>" & dosyanesne.read(10) & "<br>"

dosyanesne.close
set fso=nothing
%>
