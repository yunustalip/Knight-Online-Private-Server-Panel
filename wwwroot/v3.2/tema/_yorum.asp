<%
response.expires=0
Response.Charset = "iso-8859-9"
id=Request.QueryString("id")
if not id="" then
	if isnumeric(id)=false then
		response.redirect "index.asp"
	end if
Set rs = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select * FROM yorum where onay=0 and blog_id="&id&" ORDER BY tarih desc"
rs.open rSQL,data,3,3
adet = rs.recordcount
if not rs.eof then

sayfa = Request.QueryString("sayfa")
    if isnumeric(sayfa)=false then
        Response.redirect "index.asp"
    Else
if sayfa="" then sayfa=1
rs.pagesize = 10
sayfa_sayisi = rs.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
rs.absolutepage = sayfa
mode = 2
for i=1 to rs.pagesize
if rs.eof then
exit for
end if
	if mode=1 then
	stil="yorum_t"
	else
	stil="yorum_t2"
	end if
	
Yazan=rs("ekleyen")
Tarih=FormatDateTime(rs("tarih"),1)&"&nbsp;&"&FormatDateTime(rs("tarih"), 4)
Yorum=MesajFormatla(rs("yorum"))
%>
<div align="center">
<table width="100%" cellpadding="0" class="<%=stil%>">
	<tr>
		<td><font class="orta"><span style="float:left"><b>Ekleyen:</b> <%=Yazan%></span><span style="float:right"><%=Tarih%></span></font></td>
	</tr>
	<tr>
		<td><font class="orta"><%=Yorum%></font></td>
	</tr>
</table>
</div>
<br style="font-size:4px">
<%
rs.movenext
%>
<%
	if mode=2 then
	mode=1
	else
	mode=2
	end if
%>
<% next %> 
	<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="right"><font class="orta">Toplam <%=adet%> Yorum, <%=sayfa_sayisi%> Sayfa</font></td>
			</tr>
		<tr>
			<td align="right" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""javascript:void(0)"" onclick=""open_url('_yorum.asp?id="&id&"&sayfa=1','my_site_content');"" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""javascript:void(0)"" onclick=""open_url('_yorum.asp?id="&id&"&sayfa=" & a & "','my_site_content');"" title=""Önceki"">«</a></b> "
End If
If sayfa + 5 > sayfa_sayisi Then
b = sayfa_sayisi 
Elseif sayfa_sayisi < 5 Then
sayfa_sayisi = 1
Else
b = sayfa + 5
End If
If sayfa_sayisi < 5 Then
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
Response.Write "<b><a href=""javascript:void(0)"" onclick=""open_url('_yorum.asp?id="&id&"&sayfa=" & j & "','my_site_content');"">" & j & "</a></b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""javascript:void(0)"" onclick=""open_url('_yorum.asp?id="&id&"&sayfa=" & a & "','my_site_content');"" title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""javascript:void(0)"" onclick=""open_url('_yorum.asp?id="&id&"&sayfa=" & sayfa_sayisi & "','my_site_content');"" title=""Son Sayfa"">»»</a></b>"
End If
rs.close : set rs = nothing
%>
			</td>
		</tr>
	</table>
<% End if %>
<% Else %>
<%
Set bg = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select id,yorumdurum FROM blog where id="&id&""
bg.open rSQL,data,3,3
if not bg.eof then
if bg("yorumdurum")="1" then
Else
%>
<br>
<center><font class="orta"><b>Henüz Yorum Yapýlmadý</b></font></center>
<%
End if
End if
bg.close : set bg=nothing
%>
<% End if %>
<% End if %>