<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("metin.txt")
set dosyanesne = fso.opentextfile(dosyaadi,8)
dosyanesne.writeblanklines(3)
dosyanesne.write "yaz�lan 2. bilgi"
dosyanesne.write "yaz�lan 3. bilgi"
dosyanesne.write "yaz�lan 4. bilgi"
dosyanesne.writeblanklines(1)
dosyanesne.writeline "yeni sat�r 1"
dosyanesne.writeline "yeni sat�r 2"
dosyanesne.writeline "yeni sat�r 3"

dosyanesne.close
set fso=nothing
%>
