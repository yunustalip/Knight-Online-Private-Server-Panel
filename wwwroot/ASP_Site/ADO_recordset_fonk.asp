<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritaban� kapal�<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select Set_adi, Fiyat from setler;"
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
KayitSeti.Open SQLKomut, VTNesne , 3 
%><%

' Kayitseti.MoveFirst - ilk kayda konumlan�r
' Kayitseti.MoveNext - bir sonraki kayda konumlan�r
' Kayitseti.Moveprevious - bir �nceki kayda konumlan�r
' Kayitseti.MoveLast - en son kayda konumlan�r
' Kayitseti.Move  - istenen kayda konumlan�r
%>
<%=KayitSeti.Fields(0).name%>  -  <%=KayitSeti.Fields(1).name%> <br />
<%do until kayitseti.EOF %>
<%=KayitSeti.Fields("Set_adi") & "-" & KayitSeti.Fields("Fiyat")%> <br />
<%
KayitSeti.MoveNext
loop
response.write "-------------------<br>"%>
<%kayitseti.movefirst%>
<%kayitseti.move 10%>
<%=KayitSeti.Fields(0).name%>  -  <%=KayitSeti.Fields(1).name%> <br />
<%do until kayitseti.EOF %>
<%=KayitSeti.Fields("Set_adi") & "-" & KayitSeti.Fields("Fiyat")%> <br />
<%
KayitSeti.MoveNext
loop
response.write "-------------------<br>"
%>
<% kayitseti.movelast%>
<%=KayitSeti.Fields(0).name%>  -  <%=KayitSeti.Fields(1).name%> <br />
<%do until kayitseti.BOF %>
<%=KayitSeti.Fields("Set_adi") & "-" & KayitSeti.Fields("Fiyat")%> <br />
<%
KayitSeti.MovePrevious
loop


Kayitseti.close
VTNesne.close
Set VTNesne=Nothing
%>

