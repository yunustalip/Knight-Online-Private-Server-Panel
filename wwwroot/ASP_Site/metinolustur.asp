<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("metin.txt")
'response.write (dosyaadi)
' d:\windows\system32
set metindosyasi = FSO.CreateTextFile(dosyaadi)
metindosyasi.writeline "ilk yazýlan satýrýmýz"




metindosyasi.close
set FSO=Nothing
%>
