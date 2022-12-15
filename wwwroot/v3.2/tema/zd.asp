<% if (Request.QueryString("zd"))="yaz" then %>
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.yazan.value == "") {
   alert("Adýnýzý ve Soyadýnýzý Yazýnýz.");
   return false; }
if (form.yer.value == "") {
   alert("Lütfen Bulunduðunuz Yeri Yazýnýz.");
   return false; }
if (form.mail.value == "") {
   alert("Lütfen E-mail Yazýnýz.");
   return false; }
if (form.mesaj.value == "") {
   alert("Lütfen Mesajýnýzý Yazýnýz.");
   return false; }
return true;
}
</SCRIPT>
<script type="text/javascript">
function ismaxlength(obj){
var mlength=obj.getAttribute? parseInt(obj.getAttribute("maxlength")) : ""
if (obj.getAttribute && obj.value.length>mlength)
obj.value=obj.value.substring(0,mlength)
}
</script>
<SCRIPT language=JavaScript>
	function AddForm(form)
			{
				document.formcevap.mesaj.value = document.formcevap.mesaj.value + form
				document.formcevap.mesaj.focus();
			}
</script>
			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Ziyaretçi Defterine Yaz</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td valign="top" width="550" colspan="3" class="o_tab">
<table border="0" width="98%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td align="center">
		<font class="blog"><a href="zd.asp">Defteri Oku</a></font></td>
	</tr>
</table>
<%
aid=Filtre(Request.QueryString("aid"))
if not aid="" then
if isnumeric(aid)=false then
response.redirect"index.asp"
else
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from zd where id="&aid&""
rs.open SQL,data,1,3
if not rs.eof then
alinti=rs("id")
%>
<font class="orta"><b>&nbsp;&nbsp;&nbsp;Alýntý Yapýlacak Mesaj </b>(mesajýnýzýn baþýna eklenecek):<br></font>
<div align="center">
<table width="530" id="table1" cellpadding="0" class="yorum_t">
	<tr>
		<td height="22" style="border-style: solid; border-width: 0"><font class="orta"><div style="float:left;"><b><%=rs("yazan")%></b> Tarafýndan <%=FormatDateTime(rs("tarih"),1)%>&nbsp;<%=FormatDateTime(rs("tarih"), 4)%> Tarihinde Yazýldý.</div></font></td>
	</tr>
	<tr>
		<td valign="top" style="border-style: solid; border-width: 0"><font class="orta"><b>Yer:</b> <%=rs("yer")%><br>
		<b>Mesaj:</b> <br><%=MesajFormatla(rs("mesaj"))%></font></td>
	</tr>
</table>
<br style="font-size:3px">
</div>
<%
rs.close : set rs=nothing
end if
end if
end if
if not isnumeric(Strmesajuzunluk) then Strmesajuzunluk="400"
%>
<div align="center">
<table border="0" width="530" id="table1" cellspacing="0" cellpadding="0">
<form action="islem.asp?islem=yaz&s=<%=session.sessionID%>&aid=<%=alinti%>" method="post" onSubmit="return validate(this)" name="formcevap">
	<tr>
		<td width="985" align="center" colspan="3">
		<p><b><font class="orta">Bütün Alanlarýn 
		Doldurulmasý Zorunludur.</font></b></td>
		</tr>
	<tr>
		<td width="85" align="right"><font class="orta">*Ad - Soyad :</font></td>
		<td width="1"></td>
		<td width="444"><input name="yazan" class="alan" size="37" value="<%=Request.Cookies("isim")%>"></td>
	</tr>
	<tr>
		<td width="85" align="right"><font class="orta">*Nerden :</font></td>
		<td width="1"></td>
		<td width="444"><input name="yer" class="alan" size="37" value="<%=Request.Cookies("yer")%>"></td>
	</tr>
	<tr>
		<td width="85" align="right"><font class="orta">*Mail Adresi :</font></td>
		<td width="1"></td>
		<td width="444"><input name="mail" class="alan" size="37" value="<%=Request.Cookies("mail")%>"><font class="orta">(Gizli 
		Kalacaktýr)</font></td>
	</tr>
	<tr>
		<td width="85" align="right" valign="top"><font class="orta">*Mesaj (Max.<%=Strmesajuzunluk%> 
		Krktr) :</font></td>
		<td width="1"></td>
		<td width="444">
		<textarea rows="11" cols="55" name="mesaj" class="alan"<%if session("admin")=false then%> onKeyUp="return ismaxlength(this)" maxlength="<%=Strmesajuzunluk%>"<%End if%>></textarea></td>
	</tr>
	<tr>
		<td width="85" align="right"></td>
		<td width="1"></td>
		<td width="444">
<A href="javascript:AddForm(':)')"><img src="tema/images/smileys/smile.gif" border="0"></a> 
<A href="javascript:AddForm(':(')"><img src='tema/images/smileys/frown.gif' border=0></a> 
<A href="javascript:AddForm(':D')"><img src='tema/images/smileys/biggrin.gif' border=0></a> 
<A href="javascript:AddForm(':o:')"><img src='tema/images/smileys/redface.gif' border=0></a> 
<A href="javascript:AddForm(';)')"><img src='tema/images/smileys/wink.gif' border=0></a> 
<A href="javascript:AddForm(':p')"><img src='tema/images/smileys/tongue.gif' border=0></a> 
<A href="javascript:AddForm(':cool:')"><img src='tema/images/smileys/cool.gif' border=0></a> 
<A href="javascript:AddForm(':rolleyes:')"><img src='tema/images/smileys/rolleyes.gif' border=0></a> 
<A href="javascript:AddForm(':mad:')"><img src='tema/images/smileys/mad.gif' border=0></a> 
<A href="javascript:AddForm(':eek:')"><img src='tema/images/smileys/eek.gif' border=0></a> 
<A href="javascript:AddForm(':confused:')"><img src='tema/images/smileys/confused.gif' border=0>
		</td>
	</tr>
	<tr>
		<td width="85"><div align="center"></div></td>
		<td width="1"></td>
		<td width="444"><input type="submit" name="konu" class="dugme" value="Gönder"></td>
		</tr>
</form>
</table>
					</div>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>
<% Else %>
			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Ziyaretçi Defteri</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td valign="top" width="550" colspan="3" class="o_tab">

<table border="0" width="98%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td align="center">
		<font class="blog"><a href="?zd=yaz">Deftere Yaz</a></font><br>
		</td>
	</tr>
</table>
</div>
<% Set zd_msg = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select * FROM zd where onay=0 ORDER BY tarih desc"
zd_msg.open rSQL,data,3,3
adet = zd_msg.recordcount
if not zd_msg.eof then

sayfa = Request.QueryString("sayfa")
    if isnumeric(sayfa)=false then
        Response.redirect "index.asp"
    Else
if sayfa="" then sayfa=1
zd_msg.pagesize = 10
sayfa_sayisi = zd_msg.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
zd_msg.absolutepage = sayfa
mode = 2
for i=1 to zd_msg.pagesize
if zd_msg.eof then
exit for
end if
	if mode=1 then
	stil="yorum_t"
	else
	stil="yorum_t2"
	end if
%>
<div align="center">
<table width="530" id="table1" cellpadding="0" class="<%=stil%>">
	<tr>
		<td height="22" style="border-style: solid; border-width: 0"><font class="orta"><div style="float:left;"><b><%=zd_msg("yazan")%></b> Tarafýndan <%=FormatDateTime(zd_msg("tarih"),1)%>&nbsp;<%=FormatDateTime(zd_msg("tarih"), 4)%> Tarihinde Yazýldý.</div>
		<div style="float:right;"><b><a href="?zd=yaz&aid=<%=zd_msg("id")%>">Alýntý Yap</a></b></div>
		</font></td>
	</tr>
	<tr>
		<td valign="top" style="border-style: solid; border-width: 0"><font class="orta"><b>Yer:</b> <%=zd_msg("yer")%><br>
		<b>Mesaj:</b> <br><%=MesajFormatla(zd_msg("mesaj"))%></font></td>
	</tr>
</table>
<br style="font-size:3px">
</div>
<%
zd_msg.movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
%>
<% next %> 
<div align="center">
	<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="center"><font class="orta">Toplam <%=adet%> Mesaj, <%=sayfa_sayisi%> Sayfada Gösterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?sayfa=1"" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?sayfa=" & a & """ title=""Önceki"">«</a></b> "
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
Response.Write "<b><a href=""?sayfa=" & j & """>" & j & "</a><b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?sayfa=" & a & """ title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""?sayfa=" & sayfa_sayisi & """ title=""Son Sayfa"">»»</a></b>"
End If
zd_msg.close : set zd_msg = nothing
%>
			</td>
		</tr>
	</table>
</div>
<% End if %>
<% Else %>
<font class="orta">Kayýt Bulunamadý</font>
<% End if %>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>
<% End if %>