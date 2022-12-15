<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=("d:\windows\system32\inetsrv\data\HitCnt.cnt")
set dosyanesne = fso.opentextfile(dosyaadi,1)
okunanbilgi=dosyanesne.readall

response.Write(okunanbilgi) & "<br>"
okunanbilgi=dosyanesne.readline

dosyanesne.close
set fso=nothing
%>
