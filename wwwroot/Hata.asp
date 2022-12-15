<html >

<head>
<style type="text/css">
<!--
body,td,th,input {
	color: #00CC00;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
	font-weight:bold;
	background-color: #000000;
}
body {
	background-color: #000000;
}
-->
</style>
<script language="javascript">
function formkontrol(){
if (document.getElementById('username').value==''){
document.getElementById('loginb').disabled=true;
return false;
}


}
</script>
</head>

<body>
<p align="center">Power Up storeyi gezebilmek için.<br />
  Lütfen kullanýcý giriþi yapýnýz.
</p>
<div align="center">
<table width="258" border="0">
<tr>
   <td height="21" align="center" background="imgs/menubg.gif" colspan="2">
   <font color="#FFFFFF">Kullanýcý Giriþi</font></td>
              </tr>
      <tr><td>
	<form action="loginok2.asp" method="post" id="loginp" name="loginp" >
	<p>Kullanýcý Adý : </p>
		</td><td>
	<p>
	<input  name="username" id="username" type="text" maxlength="21" size="20"  /> </td>
	</tr><tr><td>
	<p>Þifre :</td><td>
	<p><input name="pwd" id="pwd" type="password" maxlength="13" size="20" /></tr>
	
	<tr ><td colspan="2">
	<p align="center">
	<input type="submit" value="Giriþ" name="loginb" id="loginb" class="styleform" onClick="logingiris()" />
	</form></td>
       
          </tr>
        
	</td>
	</tr></table>	
</div>
&nbsp;