			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik"><%=baslik%></font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3">

<%
kategori=Request.QueryString("kategori")

	Set kat = Server.CreateObjecT("ADODB.recordSet")
	SQL = "Select * FROM galeri_kat ORDER BY id asc"
	kat.open SQL,data,3,3

set saydir = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as resim_saydir from galeri"
saydir.open SQL,data,1,3
resimsayi=saydir("resim_saydir")
saydir.close : set saydir = nothing

if kat.eof then
response.write "<center><b><font class=""orta"">Kayýt Bulunamadý</font></b></center>"
else
	toplam=kat.recordcount
	sirala=3
%>
<center>
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
<tr>
<%
for i = 1 to toplam
if kat.eof or kat.bof then exit for
set say = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as resim_say from galeri where kat_id= "&kat("id")&""
say.open SQL,data,1,3
%>
<td width="33%"> 
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10%"><img src="tema/images/mini-category.gif"></td>
<% if not say("resim_say")=0 then %>
		<td width="90%" align="left"><font class="orta"><a href="?kategori=<%=kat("id")%>"><%=kat("isim")%></a> (<%=say("resim_say")%>)</font></td>
<% Else %>
		<td width="90%" align="left"><font class="orta"><%=kat("isim")%> (<%=say("resim_say")%>)</font></td>
<% End if %>
	</tr>
</table>
</td>
<% 
If i mod sirala = 0 Then 
Response.Write "</tr><tr>" 
End If 
%>
<%
say.close : set say=nothing
kat.movenext
next
%>
<% kat.close : set kat=nothing %>
					</td>
				</tr>
			</table>
<center><b><font class="orta">Toplam <%=toplam%> Kategori, <%=resimsayi%> Resim.</font></b></center>
<%
if isnumeric(kategori)=false or kategori="" then
set galeri = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri order by tarih DESC"
galeri.open SQL,data,1,3
if not galeri.eof then
%>
<br>
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td align="center" class="tool" height="25"><font class="orta">Son 10 Resim</font></td>
	</tr>
	<tr>
		<td>
              <MARQUEE ID="kay" DIRECTION="left" WIDTH="100%" SCROLLAMOUNT="3" SCROLLDELAY="50" STYLE="border:0px none;" onmouseover="kay.scrollAmount=0" onmouseout="kay.scrollAmount=3">
<% For p = 1 To 10
if galeri.eof Then exit For %><a ONCLICK="window.open('slide.asp?kat=<%=galeri("kat_id")%>&id=<%=galeri("id")%>','slide','top=20,left=20,width=578,height=474,toolbar=no,scrollbars=no');" href="#slide" title="<%=galeri("isim")%>">&nbsp;<img src="<%=galeri("url_kucuk")%>" width="100" height="100" border="0"></a><% galeri.movenext : next %>
              </MARQUEE>
		</td>
	</tr>
</table>
<% End if
galeri.close : set galeri = nothing 
end if
%>
<%
end if
if isnumeric(kategori)=true and kategori<>"" then
response.write "<br />"
Set zd_msg = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select * FROM galeri where kat_id="&kategori&" ORDER BY id desc"
zd_msg.open rSQL,data,3,3
if not zd_msg.eof then
adet = zd_msg.recordcount

sayfa = Request.QueryString("sayfa")
    if isnumeric(sayfa)=false then
        Response.redirect "index.asp"
    Else
If sayfa="" Then sayfa=1
zd_msg.pagesize = 20
sayfa_sayisi = zd_msg.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
zd_msg.absolutepage = sayfa
yanyana = 4
%>
<div align="center">
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
<tr>
<%
for i = 1 to zd_msg.pagesize
if zd_msg.eof or zd_msg.bof then exit for
%>
<td valign="center"> 
<div align="center">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse" align="center">
	<tr>
		<td align="center"><a ONCLICK="window.open('slide.asp?kat=<%=zd_msg("kat_id")%>&id=<%=zd_msg("id")%>','slide','top=20,left=20,width=578,height=474,toolbar=no,scrollbars=no');" href="#slide">
		<img src="<%=zd_msg("url_kucuk")%>" width="125" height="125" onerror="imajlar/spacer.gif" alt="<%=zd_msg("aciklama")%>" border="0"></a></td>
	</tr>
	<tr>
		<td height="6"><center><font class="orta" style="font-size:11px"><%=Left(zd_msg("isim"),20)%></center></font></td>
	</tr>
</table>
</div>
</td>
<% 
If i mod yanyana = 0 Then 
Response.Write "</tr><tr>" 
End If 
%>
<%
zd_msg.movenext
next
%>
</tr>
</table>
</div>
<input type="submit" value="Slide Olarak Göster" class="dugme" ONCLICK="window.open('slide.asp?kat=<%=kategori%>','slide','top=20,left=20,width=578,height=474,toolbar=no,scrollbars=no');">
<div align="center">
	<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="center"><font class="orta">Toplam <%=adet%> Resim, <%=sayfa_sayisi%> Sayfada Gösterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?kategori="&kategori&"&sayfa=1"" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?kategori="&kategori&"&sayfa=" & a & """ title=""Önceki"">«</a></b> "
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
Response.Write "<b><a href=""?kategori="&kategori&"&sayfa=" & j & """>" & j & "</a></b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?kategori="&kategori&"&sayfa=" & a & """ title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""?kategori="&kategori&"&sayfa=" & sayfa_sayisi & """ title=""Son Sayfa"">»»</a></b>"
End If
%>

			</td>
		</tr>
	</table>
</div>
<% End if %>

<%
Else
Response.Write "<B><center><font class=""orta"">Kayýt Bulunamadý....</font></center></B>"
End if
%>
<% End if %>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>