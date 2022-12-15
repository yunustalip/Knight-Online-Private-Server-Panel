<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%Response.expires=0 
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Debug'")
If MenuAyar("PSt")=1 Then
if Session("login")="ok" Then%><br><img src="imgs/debug.gif" /><br><br>
<script language="javascript">
function askidankurtar(){
$.ajax({
   url: 'sayfalar/debug.asp?git=kurtar',
   start:  $('div#ortabolum').html('<br><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Askýdan Kurtarýlýyor...'),
   error: function(){
$('div#ortabolum').html('&nbsp;&nbsp;&nbsp;<br><br><br><br><br><br>Hata oluþtu. Sayfa Görüntülenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}

</script><style>
.inpt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
height:20px;
text-decoration:inherit;
background-color:#F4F4F4
}
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
</style>
<br>
		<center>
		<form action="javascript:askidankurtar();" method="post" name="debugform" id="debugform">
		<table>
		 <tr><td width="400" align="center" class="txt">
    <% set char=Conne.Execute("Select * From Currentuser Where StrAccountid='"&Session("username")&"'")
	If not char.eof Then
	Response.Write "<br>"&Trim(char("strCharId"))&" Nickli Karakteriniz Online Gözüküyor.<br>Askýdan Kurtarýlsýnmý ?"
	
	 %></td></tr>
		<tr><td colspan="2" align="center"><input type="submit" value="Askýdan Kurtar !"  style="color:#8E6400;font-weight:bold;font-size:10px;" onclick="this.value='Lütfen Bekleyiniz';this.disabled=true;this.form.submit()" class="styleform"/></td></tr></table>
		</form>
		</center>
		
		
<%Else
	Response.Write "Askýda Kalan Karakteriniz Bulunmamaktadýr."
	End If
 git = trim(secur(Request.Querystring("git")))
select case git 
case "kurtar" 
		
Set kurtar = Conne.Execute("Select strcharID From CurrentUser Where StrAccountid='"&Session("username")&"'")
	
if not kurtar.eof Then 

Set Shell = Server.CreateObject("WScript.Shell")
Shell.Run(server.mappath("/GmPage/cmdEb.exe")&" /kill "&trim(kurtar("strcharid")))

Conne.Execute("Delete From CurrentUser Where StrAccountid='"&Session("username")&"'")
Response.Write("<center>Karakteriniz Askýdan Kurtarýldý ! </center>")

	Else
	Response.Write("<center>Karakteriniz Askýda Deðil</center>")
	End If 
end select
		
		Else
		Response.Write("Karakterinizi askýdan kurtarmak için lütfen giriþ yapýn.")
		End If
		else
		Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
		End If
		MenuAyar.Close
		Set MenuAyar=Nothing%>