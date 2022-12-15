<%
Set FSO = CreateObject("Scripting.FileSystemObject")
dosyaadi=server.mappath("/asp_site/xml/Bizim.XML")
set dosyanesne = fso.createtextfile(dosyaadi)


Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/asp_site/data/egitim.mdb")
VTNesne.Open ConnectString
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select * from setler;"
KayitSeti.Open SQLKomut, VTNesne , 3 


dosyanesne.writeline "<?xml version=""1.0"" encoding=""iso-8859-9""?>"
dosyanesne.writeline "<rss version=""2.0"">"
dosyanesne.writeline "<datadizin>"

do until kayitseti.EOF 
dosyanesne.writeline "   <SET>"
dosyanesne.writeline "      <no>" & KayitSeti.Fields("no") & "</no>"
dosyanesne.writeline "      <kategori>" & KayitSeti.Fields("kategori") & "</kategori>"
dosyanesne.writeline "      <set_adi>" & KayitSeti.Fields("set_adi") & "</set_adi>"
dosyanesne.writeline "      <fiyat>" & KayitSeti.Fields("Fiyat") & "</fiyat>"
dosyanesne.writeline "   </SET>"

KayitSeti.MoveNext
loop

dosyanesne.writeline "</datadizin>"
dosyanesne.writeline "</rss>"
%><%

Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

dosyanesne.close
set fso=nothing
%>

