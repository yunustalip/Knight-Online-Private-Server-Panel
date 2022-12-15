<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("sinan.txt")
set dosyanesne = fso.opentextfile(dosyaadi,1)
'okunanbilgi=dosyanesne.readall

call satiratlat(4)
okunanbilgi=dosyanesne.readline
response.Write(okunanbilgi) & "<br>"
okunanbilgi=dosyanesne.readline
response.Write(okunanbilgi)
'okunanbilgi=dosyanesne.readline
'okunanbilgi=dosyanesne.readline
'okunanbilgi=dosyanesne.readline
'response.Write(okunanbilgi)


sub satiratlat(sayi)
for i = 1 to sayi
     dosyanesne.skipline
next
end sub

dosyanesne.close
set fso=nothing
%>
