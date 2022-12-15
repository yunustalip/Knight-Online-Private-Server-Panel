<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%Response.Charset = "iso-8859-9"
Response.expires=0 
Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='SubmitTicket'")
If MenuAyar("PSt")=1 Then
	if Session("login")="ok" Then %>
	<script language="javascript">
function ticketyolla(){
$.ajax({
   type: 'POST',
   url: 'sayfalar/submitticket.asp?durum=ok',
   data: $('form#ticketform').serialize() ,
   start:  $('div#ortabolum').html('<br><br><br><br>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>'),
   error: function(){
$('div#ortabolum').html('&nbsp;&nbsp;&nbsp;<br>Hata oluþtu. Sayfa Görüntülenemiyor...');
},
   success: function(ajaxCevap) {
      $('div#ortabolum').html(ajaxCevap);
   }
});
}
</script>
<style>
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
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
.inpt2{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:12px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
text-decoration:inherit;
background-color:#F4F4F4
}
</style>
    <div align="center" class="style4" style="font-size:16px; padding-top:10px"><img src="imgs/ticket.gif" align="absmiddle" />
    </div>
	<br><br>
	<ul>
		<li> 
		Sitedeki hatalar veya oyun içi hile bug bildirim için lütfen ticket atýnýz.<br />
		</li>
		<li>
		Hile bildirimlerinde resimleri herhangi bir siteye upload edip linki yollayýn.
		</li>
	</ul>
	</div>   
<form action="javascript:ticketyolla();" method="post" id="ticketform" name="ticketform">
<table width="350" height="200" border="0" align="center">
  <tr>
    <td width="113" class="txt">Oyun nickiniz:</td>
    <td width="227"><select name="charname" class="inpt" style="padding-top:2px">
	<%set chars=Conne.Execute("select straccountid,strcharid1,strcharid2,strcharid3 from account_char where straccountid='"&Session("username")&"'")
	if len(trim(chars("strcharid1")))>0 Then
	Response.Write "<option value="""&chars("strcharid1")&""" style=""height:15px"">"&chars("strcharid1")&"</option>"
	End If
	if len(trim(chars("strcharid2")))>0 Then
	Response.Write "<option value="""&chars("strcharid2")&""" style=""height:15px"">"&chars("strcharid2")&"</option>"
	End If
	if len(trim(chars("strcharid3")))>0 Then
	Response.Write "<option value="""&chars("strcharid3")&""" style=""height:15px"">"&chars("strcharid3")&"</option>"
	End If
	%>
	</select></td>
  </tr>
  <tr>
    <td class="txt">E-Mail Adresiniz:</td>
    <td><input type="text" name="email"  class="inpt" size="25"></td>
  </tr>
  <tr>
    <td class="txt">Konu:</td>
    <td><input type="text" name="subject"  class="inpt" autocomplete="false"></td>
  </tr>
  <tr>
    <td height="95" valign="top" class="txt">Mesaj:</td>
    <td valign="top"><textarea cols="35" rows="8" name="message" class="inpt2" style="font-size:11px"></textarea></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="right"><input type="submit" value="" style=" background:url(../imgs/next.gif) right; background-repeat:no-repeat; width:95px; height:25px " ></td>
  </tr>
</table>
    </form>

<% durum = trim(secur(Request.Querystring("durum"))) 
if durum="" Then
durum="yok"
End If 
select case durum 
case "yok" 
case "ok" 
charname = Trim(secur(request.form("charname")))
subject = Trim(secur(request.form("subject")))
message = Trim(secur(request.form("message")))
email = Trim(secur(request.form("email")))

If charname="" Or subject="" Or email="" Or message="" Then
Response.Write("<script>alert('Boþ Býraktýðýnýz Alanlar Var.  Boþ Býraktýðýnýz Yerleri Doldurun.')</script>")
Response.End
Else 

If EmailKontrol(email)="False" Then
Response.Write("<script>alert('Size Ulaþabilmemiz Için Lütfen Geçerli Bir E-Mail Adresi Giriniz.')</script>")
Response.End
End If

Set gtict = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tickets where charid='"&Session("username")&"' And durum=1 And date like '%"&day(date())&"%' And date like '%"&month(date())&"%' And date like '%"&year(date())&"%'"
gtict.Open sql,conne,1,3

If Not gtict.Eof Then 
Response.Write("<script>alert('Bugün Içinde Art Arda Mesaj Gönderemezsiniz. Yeni Mesaj Gönderebilmeniz Için Yöneticinin Mesajýnýzý Okumuþ Olmasý Gerekir.\nYarýn Tekrar Ticket Atabilirsiniz.')</script>")
Response.End
Else

If trim(charname)=trim(chars("strcharid1")) or trim(charname)=trim(chars("strcharid2")) or trim(charname)=trim(chars("strcharid3")) Then

gtict.Addnew
gtict("charid")=Session("username")
gtict("name")=charname
gtict("email")=email
gtict("subject")=subject
gtict("message")=message
gtict("date")=now
gtict("durum")=1
gtict.update
Response.Write("<script>alert('Mesajýnýz Tarafýmýza Ýletilmiþtir. En kýsa zamanda size dönülecektir. Ýlginiz için Teþekkür Ederiz !')</script>")
Else
Response.Write "Ýçerik Yöneticisi Devrede!!! Bu karakter size ait deðildir !"
End If
End If 
	End If 
	case else 
	end select 
	else 
	Response.Write("Ticket Gönderebilmeniz için üye giriþi yapmanýz Gerekmektedir.")
	End If 	
	else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing	%>