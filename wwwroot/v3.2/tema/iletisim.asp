			<table border="0" width="550" id="table1" cellpadding="0" style="border-collapse: collapse">
				<tr>
					<td height="26" width="15">
					<img border="0" src="tema/images/blok_1.gif" width="15" height="26"></td>
					<td height="26" background="tema/images/blok_2.gif" width="525">
					<p align="center">
					<font class="baslik">�leti�im</font></td>
					<td height="26" width="10">
					<img border="0" src="tema/images/blok_3.gif" width="10" height="26"></td>
				</tr>
				<tr>
					<td class="o_tab" valign="top" width="550" colspan="3">
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.isim.value == "") {
   alert("Ad�n�z� ve Soyad�n�z� Yaz�n�z.");
   return false; }
if (form.yer.value == "") {
   alert("L�tfen Bulundu�unuz Yeri Yaz�n�z.");
   return false; }
if (form.konu.value == "") {
   alert("L�tfen Konu Ba�l���n� Yaz�n�z.");
   return false; }
if (form.mail.value == "") {
   alert("L�tfen Mail Adresinizi Yaz�n�z.");
   return false; }
if (form.mesaj.value == "") {
   alert("L�tfen Mesaj�n�z� Yaz�n�z.");
   return false; }
return true;
}
</SCRIPT>
<form action="islem.asp?islem=ilet&s=<%=session.sessionID%>" method="post" onSubmit="return validate(this)">

<div align="center">

<table border="0" width="530" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td width="985" align="right" colspan="3">
		<p align="center"><font class="orta">Ba��nda (*) Bulunan B�t�n Alanlar�n Doldurulmas� Zorunludur.</font></td>
		</tr>
	<tr>
		<td width="90" align="right"><font class="orta">Ad - Soyad* :</font></td>
		<td width="6"></td>
		<td width="454"><input name="isim" class="alan" size="47" value="<%=Request.Cookies("isim")%>"></td>
	</tr>
	<tr>
		<td width="90" align="right"><font class="orta">Nerden?* :</font></td>
		<td width="6"></td>
		<td width="454"><input name="yer" class="alan" size="47" value="<%=Request.Cookies("yer")%>"></td>
	</tr>
	<tr>
		<td width="90" align="right"><font class="orta">Konu* :</font></td>
		<td width="6"></td>
		<td width="454"><input name="konu" class="alan" size="47"></td>
	</tr>
	<tr>
		<td width="90" align="right"><font class="orta">Mail Adresi* :</font></td>
		<td width="6">&nbsp;</td>
		<td width="454"><input name="mail" class="alan" size="47" value="<%=Request.Cookies("mail")%>"></td>
	</tr>
	<tr>
		<td width="90" align="right"><font class="orta">Web Adresi :</font></td>
		<td width="6">&nbsp;</td>
		<td width="454">
		<input name="url" class="alan" size="47" value="http://"></td>
	</tr>
	<tr>
		<td width="90" align="right"><font class="orta">�leti* :</font></td>
		<td width="6"></td>
		<td width="454"><textarea rows="11" cols="54" class="alan" name="mesaj"></textarea></td>
	</tr>
	<tr>
		<td width="90"></td>
		<td width="6"></td>
		<td width="454">
		<input type="submit" name="konu1" class="dugme" value="G�nder"></td>
		</tr>
</table>
</div>
</form>
					</td>
				</tr>
				<tr>
					<td colspan="3" width="550"><img src="tema/images/orta_alt.gif"></td>
				</tr>
			</table>