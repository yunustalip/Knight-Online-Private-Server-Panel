<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%response.expires=0
Dim MenuAyar,ksira,humansayi,karussayi,toplamchar,REFERER_URL,REFERER_DOMAIN,s,humansay,karussay,x
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Register'")
If MenuAyar("PSt")=1 Then

If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Register")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("Script_Name")
If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Register")
End If
%>
		<script language="JavaScript1.1">
function testPassword(passwd)
{
var description = new Array();
description[0] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=30 bgcolor=#ff0000></td><td height=15 width=120 bgcolor=#dddddd></td></tr></table></td><td class=bold>Çok Zayýf</td></tr></table>";
description[1] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=60 bgcolor=#bb0000></td><td height=15 width=90 bgcolor=#dddddd></td></tr></table></td><td class=bold>Zayýf</td></tr></table>";
description[2] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=90 bgcolor=#ff9900></td><td height=15 width=60 bgcolor=#dddddd></td></tr></table></td><td class=bold>Orta</td></tr></table>";
description[3] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=120 bgcolor=#00bb00></td><td height=15 width=30 bgcolor=#dddddd></td></tr></table></td><td class=bold>Güçlü</td></tr></table>";
description[4] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=150 bgcolor=#00ee00></td></tr></table></td><td class=bold>Çok Güçlü</td></tr></table>";
description[5] = "<table border=0 cellpadding=0 cellspacing=0 align=right><tr><td><table cellpadding=0 cellspacing=2><tr><td height=15 width=150 bgcolor=#dddddd></td></tr></table></td><td class=bold>Bekleniyor.</td></tr></table>";

		var intScore   = 0
		var strVerdict = 0
		
		// PASSWORD LENGTH
		if (passwd.length==0 || !passwd.length)                         // length 0
		{
			intScore = -1
		}
		else if (passwd.length>0 && passwd.length<5) // length between 1 and 4
		{
			intScore = (intScore+3)
		}
		else if (passwd.length>4 && passwd.length<8) // length between 5 and 7
		{
			intScore = (intScore+6)
		}
		else if (passwd.length>7 && passwd.length<12)// length between 8 and 15
		{
			intScore = (intScore+12)
		}
		else if (passwd.length>11)                    // length 16 or more
		{
			intScore = (intScore+18)
		}
		
		
		// LETTERS (Not exactly implemented as dictacted above because of my limited understanding of Regex)
		if (passwd.match(/[a-z]/))                              // [verified] at least one lower case letter
		{
			intScore = (intScore+1)
		}
		
		if (passwd.match(/[A-Z]/))                              // [verified] at least one upper case letter
		{
			intScore = (intScore+5)
		}
		
		// NUMBERS
		if (passwd.match(/\d+/))                                 // [verified] at least one number
		{
			intScore = (intScore+5)
		}
		
		if (passwd.match(/(.*[0-9].*[0-9].*[0-9])/))             // [verified] at least three numbers
		{
			intScore = (intScore+5)
		}
		
		
		// SPECIAL CHAR
		if (passwd.match(/.[!,@,#,$,%,^,&,*,?,_,~]/))            // [verified] at least one special character
		{
			intScore = (intScore+5)
		}
		
																 // [verified] at least two special characters
		if (passwd.match(/(.*[!,@,#,$,%,^,&,*,?,_,~].*[!,@,#,$,%,^,&,*,?,_,~])/))
		{
			intScore = (intScore+5)
		}
	
		
		// COMBOS
		if (passwd.match(/([a-z].*[A-Z])|([A-Z].*[a-z])/))        // [verified] both upper and lower case
		{
			intScore = (intScore+2)
		}

		if (passwd.match(/(\d.*\D)|(\D.*\d)/))                    // [FAILED] both letters and numbers, almost works because an additional character is required
		{
			intScore = (intScore+2)
		}
 
																  // [verified] letters, numbers, and special characters
		if (passwd.match(/([a-zA-Z0-9].*[!,@,#,$,%,^,&,*,?,_,~])|([!,@,#,$,%,^,&,*,?,_,~].*[a-zA-Z0-9])/))
		{
			intScore = (intScore+2)
		}
	
	
		if(intScore == -1)
		{
		   strVerdict = description[5];
		}
		else if(intScore > -1 && intScore < 16)
		{
		   strVerdict = description[0];
		}
		else if (intScore > 15 && intScore < 25)
		{
		   strVerdict = description[1];
		}
		else if (intScore > 24 && intScore < 35)
		{
		   strVerdict = description[2];
		}
		else if (intScore > 34 && intScore < 45)
		{
		   strVerdict = description[3];
		}
		else
		{
		   strVerdict = description[4];
		}
	
	document.getElementById("passcheck").innerHTML= (strVerdict);
	
}
// End-->
</script>
<script type="text/javascript" language="JavaScript">
function reloadImage()
{
   document.images["simage"].src =  'securityImage.asp?rand='+Math.random() * 1000000
}
</script><br>

<script language="javascript">
function formyolla(){
$.ajax({
   type: 'POST',
   url: 'registercontrol.asp',
   data: $('#kayit').serialize() ,
   start:  $('#registerreply').html('&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   success: function(ajaxCevap) {
      $('#registerreply').html(ajaxCevap);
   }
});
}
function icerikal(){
$.ajax({
   type: 'GET',
   url: 'kontrol.asp',
   data: $('#usernam').serialize() ,
   start:  $('#sonuc').html('&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   success: function(ajaxCevap) {
      $('#sonuc').html(ajaxCevap);
   }
});
}
</script>
<script type="text/javascript" src="js/jquery.validate.js"></script>
    <script type="text/javascript">

$(document).ready(function() {

    $('#kayit').validate({
    rules: {
  username: {
  required: true
  },
  pwd: {
  required: true,
  minlength: 4
  },
  pwd2: {
  equalTo: "#pwdx"
  },
  
  email: {
  required: true,
  email: true
  },
  gizlicevap: {
  required: true
  },
  formHC: {
  required: true
  }

    },

	messages: {

   username: {
   required: ' Boþ Býrakýlamaz!'
   },
   pwd: {
   required: ' Boþ Býrakýlamaz!',
   minlength: ' En az 4 karakter giriniz.'
   },
   pwd2: {
   equalTo: ' Þifreler Uyuþmuyor !'
   },
   email: {
   required: ' Boþ Býrakýlamaz!',
   email: ' Lütfen geçerli bir e-mail adresi giriniz'
   },
   gizlicevap: {
   required: ' Boþ Býrakýlamaz!'
   },
   formHC: {
	required: ' Boþ Býrakýlamaz!'
   }

      }

  });

      });

      </script>

<style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}

</style><img src="imgs/register.gif" /><br /><br /><br>
<%set humansayi=Conne.Execute("select count(nation) nation from userdata where nation=2")
	humansayi=humansayi("Nation")
	set karussayi=Conne.Execute("select count(nation) nation from userdata where nation=1")
	karussayi=karussayi("Nation")
	toplamchar=humansayi+karussayi%>
        <table align="center">
        <tr>
          <td width="178" colspan="2"  align="center"><b>Oyuncu Ýstatistikleri</b></td>
        </tr>
        <tr><td>Toplam Karus oyuncu :</td>
        <td width="100"><%=karussayi%></td>
        </tr>
        <tr><td>Toplam Human oyuncu :</td>
        <td><%=humansayi%></td>
        </tr>
        <tr><td colspan="2">
        
 <table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr><td align="center">Human: <%=round(100/toplamchar*humansayi,1)&"%"%></td><td align="center">Karus: <%=round(100/toplamchar*karussayi,1)&"%"%></td>
        <tr><td colspan="2">
	<img src="imgs/solbar.gif" align="middle" >
	<img src="imgs/humano.gif" align="middle" style="position:relative;left:-8px;z-index:2">
	<img src="imgs/humanbar.gif" width="<%
	humansay=round(60/toplamchar*humansayi)
	karussay=round(60/toplamchar*karussayi)
	if karussay<=0 Then
	humansay=humansay-1
	karussay=1
	End If
	Response.Write humansay%>%" height="16" align="middle" style="position:relative;left:-15px;z-index:0">
	<img src="imgs/warX.gif" align="middle" style="position:relative;left:-21px;z-index:2">
	<img src="imgs/karusbar.gif" width="<%
	Response.Write karussay%>%" height="16" align="middle" style="position:relative;left:-27px;z-index:1">
	<img src="imgs/karuso.gif" align="middle" style="position:relative;left:-33px;z-index:3">
	<img src="imgs/sagbar.gif" align="middle" style="position:relative;left:-39px;z-index:2"></td>
          </tr>
        </table>
        <% if humansayi<karussayi Then
		Response.Write "Eþitliði dengelemek için <font color=""red""><b>Human</b></font> char açmanýz önerilir."
		elseif karussayi<humansayi Then
		Response.Write "Eþitliði dengelemeniz için <font color=""blue""><b>Karus</b></font> char açmanýz önerilir."
		else
		End If%>
        </td>
        </tr>
        </table>
<style>
.idc-text{
border-style:dashed;
border-color:#89640B;
font-size:11px;
color:#89640B;
font-weight:bold;
border-width:thin;
background-color:#FFF
}
	.rselect{
	background-color:#DCD1BA;
	color:#89640B; 
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	}

</style>
<form action="javascript:formyolla();" method="post" id="kayit" name="kayit">
<table width="560" border="0" align="left" style="position:relative;top:10px;left:45px" >
  <tr>
    <td width="200" align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Kullanýcý Adýnýz</span></td>
    <td width="360" align="left" valign="top">
     <input name="username" class="idc-text" type="text" id="usernam" autocomplete="off" maxlength="21" size="20" class="idc-text" onblur="icerikal();" >
     <span id="sonuc"></span></td>
  </tr>
  <tr>
    <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Þifreniz</span></span></td>
    <td align="left" style="font-weight:bold;color:red" valign="top"><input name="pwd" type="password" autocomplete="off" id="pwdx" maxlength="13" size="20" class="idc-text"   onkeyup="testPassword(document.forms.kayit.pwd.value);"><span id="passcheck"></span></td>
  </tr>
  <tr >
    <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Þifreniz (Onay) </span></span></td>
    <td align="left" style="font-weight:bold;color:red" valign="top"><input name="pwd2" type="password" autocomplete="off" id="pwd2" maxlength="13" size="20" class="idc-text"  ></td>
  </tr>

    <tr >
    <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>E- Mail adresiniz </span></span></td>
    <td align="left" style="font-weight:bold;color:red" valign="top"><input type="text" id="email" name="email" maxlength="50" size="30" class="idc-text"  ></td>
  </tr>
  <tr >
  <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Gizli Soru</span></td>
  <td align="left" style="font-weight:bold;color:red;" valign="top">
  <select name="gizlisoru" class="rselect">
	<%
	for x=1 to 8
	Response.Write "<option value="""&x&""">"&gizlis(""&x&"")&"</option>"
	next
	%>
  </select></td>
  </tr>
  <tr >
  <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Gizli Cevap</span></td>
  <td align="left" style="font-weight:bold;color:red" valign="top"><input type="text" name="gizlicevap"  id="gizlicevap" autocomplete="off" maxlength="100" size="20" class="idc-text"  ></td>
  </tr>
  <tr >
    <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Güvenlik Kodu </span></td>
    <td align="left" style="font-weight:bold;color:red" valign="top"><label for="field4"><img id="simage" src="securityImage.asp" style="position:relative;top:-5px"></label> <a href="#" onclick="reloadImage();return false;" style="position:relative;top:-13px"> ( Kodu Yenile )</a></td>
  </tr>
  <tr>
    <td align="center" style="color: rgb(247, 231, 33);"><img src="imgs/mnu.gif"><span style="position:relative;top:-15px;font-weight:bold"><br>Güvenlik Kodunu Girin </span></td>
    <td align="left" style="font-weight:bold;color:red" valign="top"><input name="formHC" id="formHC" size="8" maxlength="4" class="idc-text"/><div id="guvk"></div></td>
  </tr>
  <tr>
  <td></td>
  <td align="center"><input type="submit" class="styleform" id="kayitbutton" name="kayitbutton" value="KAYIT OL"/></td>
  </tr>
  <tr>
    <td colspan="2" align="center" id="registerreply" name="registerreply">&nbsp;</td>
    </tr>
</table>
</form>
		<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing%>