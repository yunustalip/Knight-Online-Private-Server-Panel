			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">Hakkýmda</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3" align="center">
<table border="0" width="530" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td>
				<font class="orta">
				<%=hakkimda%>
				</font>
		</td>
	</tr>
<%
if StrHakkimdaYorum="1" then
%>
	<tr>
		<td>

<a name="yorumlar"></a>
<div id="my_site_content">
</div>
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.Ekleyen.value == "") {
   alert("Adýnýzý ve Soyadýnýzý Yazýnýz.");
   return false; }
if (form.yorum.value == "") {
   alert("Lütfen Yorumunuzu Yazýnýz.");
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
				document.formcevap.yorum.value = document.formcevap.yorum.value + form
				document.formcevap.yorum.focus();
			}
</script>
<a name="yorum"></a>
<form method="post" action="islem.asp?islem=yorumekle&id=0&s=<%=session.sessionID%>" onSubmit="return validate(this)" name="formcevap">
<div align="center">
<table border="0" id="table1" cellspacing="1" cellpadding="0" style="border-collapse: collapse" width="530">
	<tr>
		<td width="107">
		&nbsp;</td>
		<td width="440">
		<font class="orta"><b># Yorum Yaz #</b></font></td>
	</tr>
	<tr>
		<td width="107">
		<p align="right"><font class="orta">Ýsim :</font></td>
		<td width="440">
		<input type="text" class="alan" name="Ekleyen" size="26"></td>
	</tr>
	<tr>
		<td width="107" valign="top">
		<p align="right"><font class="orta">Yorum :<%if session("admin")=false then%><br>(Max. 400 Karakter)<%End if%></font></td>
		<td width="440">
		<textarea rows="6" cols="39" class="alan" name="yorum"<%if session("admin")=false then%> onKeyUp="return ismaxlength(this)" maxlength="400"<%End if%>></textarea></td>
	</tr>
	<tr>
		<td width="110"></td>
		<td width="420">
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
		<td width="110">
		&nbsp;</td>
		<td width="420">
		<input type="submit" class="dugme" value="Gönder"></td>
	</tr>
</table>
</div>
</form>


		</td>
	</tr>
<% End if %>
</table>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>