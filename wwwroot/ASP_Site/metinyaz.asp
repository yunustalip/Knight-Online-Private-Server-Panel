<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("metin.txt")
set dosyanesne = fso.opentextfile(dosyaadi,8)
dosyanesne.writeblanklines(3)
dosyanesne.write "yazýlan 2. bilgi"
dosyanesne.write "yazýlan 3. bilgi"
dosyanesne.write "yazýlan 4. bilgi"
dosyanesne.writeblanklines(1)
dosyanesne.writeline "yeni satýr 1"
dosyanesne.writeline "yeni satýr 2"
dosyanesne.writeline "yeni satýr 3"

dosyanesne.close
set fso=nothing
%>
