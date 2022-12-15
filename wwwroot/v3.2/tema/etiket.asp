			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Etiket: <%=etiket%></font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3">
<%
if not etiket="" then
Set Objetiket = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select * from etiket where etiket='"&etiket&"'"
Objetiket.open rSQL,data,1,3
if not Objetiket.eof then
adet = Objetiket.recordcount
sayfa = Request.QueryString("sayfa")
if isnumeric(sayfa)=false then sayfa=1
if sayfa="" then sayfa=1
Objetiket.pagesize = StrAramaSayi
sayfa_sayisi = Objetiket.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
Objetiket.absolutepage = sayfa
for i=1 to Objetiket.pagesize
if Objetiket.eof then exit for

Set zd_msg = Server.CreateObject("adodb.RecordSet")
SQL="SELECT * From blog WHERE id="&Objetiket("blog_id")&"" 
zd_msg.open SQL,data,3,3

tarih=zd_msg("tarih")
parcala=""&tarih&""

parcala = split(parcala,"/" )
ay=Left(MonthName(parcala(0)),3)
gun=parcala(1)
yil=Right(parcala(2),2)
%>
<div align="center">
<table border="0" width="530" id="table4" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td height="24">
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse" height="49">
	<tr>
		<td width="44" height="49" rowspan="3">
		<table border="0" width="100%" id="table2" cellpadding="0" style="border-collapse: collapse" height="100%" background="tema/images/calendar.gif">
			<tr>
				<td height="20"><center><font class="takvimay"><%=ay%>`<%=yil%></font></center></td>
			</tr>
			<tr>
				<td>
				<center><font class="takvimgun"><%=gun%></font></center></td>
			</tr>
		</table>
		</td>
		<td width="486"><font class="blog"><a href="<%=SEOLink(zd_msg("id"))%>"><%=zd_msg("konu")%></a></font></td>
	</tr>
	<tr>
		<td width="486" height="1" background="tema/images/nokta.gif"></td>
	</tr>
	<tr>
	   <td>
<%
set yorum = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yorum_say from yorum where blog_id= "&zd_msg("id")&" and onay=0"
yorum.open SQL,data,1,3
%>
<table border="0" width="486" id="table1" cellpadding="0">
	<tr>
		<td width="20"><img src="tema/images/mini-category.gif"></td>
<%
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from Kategori where id = "&zd_msg("kat_id") &""
ktg.open SQL,data,1,3
%>
		<td width="150"><a href="kategori.asp?id=<%=ktg("id")%>"><%=ktg("ad")%></a></td>
<% ktg.close : set ktg = nothing %>
		<td width="196"></td>
		<td width="20"><img src="tema/images/mini-comment.gif"></td>
		<td width="100">
		<p align="center"><font class="orta">Yorumlar(<%=yorum("yorum_say")%>)</font></td>
	</tr>
</table>
<%
yorum.close
set yorum = Nothing
%>
	   </td>
</table>
		</td>
	</tr>
	</tr>
	<tr>
		<td valign="top"><font class="orta"><%=YaziKirp(zd_msg("mesaj"),SEOLink(zd_msg("id")))%></font></td>
	</tr>
</table>
</div>
<br>
<%
zd_msg.close
set zd_msg=nothing
Objetiket.movenext
next
%>
<div align="center">
	<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="center"><font class="orta">Toplam <%=adet%> Blog, <%=sayfa_sayisi%> Sayfada Gösterilmektedir.</font></td>
			</tr>
		<tr>
			<td width="100%" align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?etiket="&etiket&"&sayfa=1"" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?etiket="&etiket&"&sayfa=" & a & """ title=""Önceki"">«</a></b> "
End If
If sayfa + 5 > sayfa_sayisi Then
b = sayfa_sayisi 
Elseif sayfa_sayisi < 5 Then
sayfa_sayisi = 1
Else
b = sayfa + 5
End If
If sayfa < 5 Then
c = 1
Else
c = sayfa - 5
End If
if c < 1 then 
c = 1
end if
For j = c To b
If j = CInt(sayfa) Then
Response.Write "<font class=""orta"">[<b>" & j & "</b>] </font>"
Else
Response.Write "<b><a href=""?etiket="&etiket&"&sayfa=" & j & """>" & j & "</a><b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?etiket="&etiket&"&sayfa=" & a & """ title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""?etiket="&etiket&"&sayfa=" & sayfa_sayisi & """ title=""Son Sayfa"">»»</a></b>"
End If
%>
			</td>
		</tr>
	</table>
</div>
<%
else
response.write "<center><font class=""orta""><b>Kayýt Bulunamadý</b></font></center>"
end if
%>
<% Objetiket.close : set Objetiket=nothing %>
<% Else %>
<center><font class="orta"><b>Etiket Seçilmedi!</b></font></center>
<% End if %>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>