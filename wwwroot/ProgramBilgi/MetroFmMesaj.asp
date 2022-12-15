<%response.charset="utf-8"%><style>
.Serit5 {
	background-image:url(http://www.metrofm.com.tr/images/Serit5.gif);
}
.SagMenuBg {
	background-image:url(http://www.metrofm.com.tr/images/SagMenuBg.gif);
	background-position:bottom;
	height:121px;
}
.style5 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
	color: #ffffff;
	height:41px;
	font-weight:bold;
	padding-left:10px;
}
.style5 a {
	font-size: 14px;
	color: #ffffff;
}
.style5 a:hover {
	font-size: 14px;
	color: #ffffff;
}
.input1 {
	background-image:url(http://www.metrofm.com.tr/images/InputBg.gif);
	height:30px;
	line-height:27px;
	border:0;
	width:144px;
	font-size:12px;
}
.input2 {
	background-image:url(http://www.metrofm.com.tr/images/InputBg1.gif);
	height:30px;
	line-height:27px;
	border:0;
	width:217px;
	font-size:12px;
}

</style>
<tr>
            <td class="style5 Serit5">CANLI YAYINA MESAJ GÖNDERİN</td>
          </tr>
<tr>
            <td class="SagMenuBg">
<table cellspacing="0" cellpadding="0" border="0" align="center" width="95%">
	<tbody><tr>
    	<td height="10" colspan="2"></td>
    </tr>
<script language="JavaScript" type="text/javascript">
	<!--
	function Kontrol()
	{
		if ((document.myform.AdSoyad.value==" Adınız Soyadınız") || (document.myform.AdSoyad.value<="          "))  {
			alert("Adınızı yazınız.");
			document.myform.AdSoyad.focus();
			return false;
		}
		if ((document.myform.AdSoyad.value=="") || (document.myform.AdSoyad.value<="          "))  {
			alert("Adınızı yazınız.");
			document.myform.AdSoyad.focus();
			return false;
		}
		if ((document.myform.Ozet.value==" Mesajınız"))  {
			alert("Mesajınızı yazınız.");
			document.myform.Ozet.focus();
			return false;
		}
		if ((document.myform.Ozet.value==""))  {
			alert("Mesajınızı yazınız.");
			document.myform.Ozet.focus();
			return false;
		}
	}
	//-->
</script>

<form action="http://www.metrofm.com.tr/mesajg.asp" method="post" onsubmit="return Kontrol();" name="myform" id="myform2">
        <tr>
          <td height="38" style="padding-right: 5px;">
	<input type="text" onfocus="if (this.value==' Adınız Soyadınız'){this.value=''}" onblur="if (this.value=='') { this.value=' Adınız Soyadınız'; }" class="input1" value=" Adınız Soyadınız" name="AdSoyad">
	<input type="text" onfocus="if (this.value==' E-posta Adresiniz'){this.value=''}" onblur="if (this.value=='') { this.value=' E-posta Adresiniz'; }" class="input1" value=" E-posta Adresiniz" name="EMail"></td>
        </tr>
        <tr>
          <td height="38" colspan="2">
	<div style="float: left;">
	<input type="text" onfocus="if (this.value==' Mesajınız'){this.value=''}" onblur="if (this.value=='') { this.value=' Mesajınız'; }" class="input2" value=" Mesajınız" name="Ozet"></div>
	<div style="float: left; padding-top: 1px; padding-left: 13px;">
	<input type="image" value="Submit" id="button" name="button" src="http://www.metrofm.com.tr/images/Gonder.gif" onclick="return Kontrol();"></div></td>
          </form></tr>

      </tbody></table>
            </td>
          </tr>