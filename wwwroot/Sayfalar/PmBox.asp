<!--#include file="../_inc/conn.asp"-->
<%Response.expires=0
Session.CodePage=65001
Response.Charset = "utf-8"
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PmBox'")
If MenuAyar("PSt")=1 Then %>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<head>
</head><br>
<img src="imgs/pmbox.gif"><br>
<script language="javascript">
function pmgonder(formname){
$.ajax({
   url: 'Sayfalar/pmbox.asp?pm=gonder',
   data: $('#'+formname).serialize() ,
   type: 'POST',
   success: function(ajaxCevap) {
$('#ortabolum').html(ajaxCevap);
$('#kullogin').fadeOut("fast"),
$('#kullogin').load('login.asp'),
$('#kullogin').fadeIn("normal");
   }
});

}
function pmsil(formname){
$.ajax({
   type: 'post',
   url: 'Sayfalar/pmbox.asp?pm=sil',
   data: $('#'+formname).serialize() ,
   success: function(ajaxCevap) {
$('#ortabolum').html(ajaxCevap);
$('#kullogin').fadeOut("fast"),
$('#kullogin').load('login.asp'),
$('#kullogin').fadeIn("fast");
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
<%
if Session("login")="ok" Then

function secur(data) 
Data = Replace( data , "'" , "", 1, -1,1)
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1)
data = Replace (data ,"Chr(39)","",1,-1,1)
secur=data 
end function
Dim Account,char1,char2,char3,pm,pmbox,pmboxtop
Set Account=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"'")
char1=trim(Account("StrCharid1"))
char2=trim(Account("StrCharid2"))
char3=trim(Account("StrCharid3"))

Set pmbox=Conne.Execute("select top 5 * from pmbox where alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"'")
Set pmboxtop=Conne.Execute("select count(alici) as toplam from pmbox where alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"'")
pm=secur(Request.Querystring("pm"))
if pm="" Then%>
<table width="535" style="position:relative;top:30px">
<tr>
<td width="473">
<fieldset>
<legend>Pm Inbox (<%=pmboxtop("toplam")%> / 5)</legend>
<form action="javascript:pmsil('pmsil')" method="post" id="pmsil" name="pmsil">
<table width="507" border="0" >
<tr>
    <td width="114"><b>Gönderen</b></td>
    <td width="140"><b>Konu</b></td>
    <td width="137"><b>Tarih</b></td>
    <td width="54"><b>Durum</b></td>
    <td width="40"><b>Sil</b></td>
    </tr>
<%if not pmbox.eof Then
do while not pmbox.eof%>
<tr>
<td><%=pmbox("gonderen")%></td>
<td><a href="#" onClick="pageload('Sayfalar/pmbox.asp?pm=oku&id=<%=pmbox("id")%>','1');return false" class="link1"><%=pmbox("konu")%></a></td>
<td><%=pmbox("tarih")%></td>
<td><%if pmbox("durum")=0 Then
Response.Write "<img src='../imgs/Mailnew.gif' alt=""Yeni Mesaj (Okunmamış)"" title=""Yeni Mesaj (Okunmamış)"">"
else
Response.Write "<img src='../imgs/Mailverified.gif' alt=""Okundu"" title=""Okundu"">"
End If%></td>
<td><input type="checkbox" name="pmdel" value="<%=pmbox("id")%>"/></td>
</tr>
<%
pmbox.movenext
loop
Response.Write " <tr><td colspan=""5"" align=""right""><input type=""submit"" value=""Sil""></td></tr>"
else
%>
<tr><td><font class="style4">Mesajınız bulunmamaktadır.</font></td></tr>
<%End If%>
</table></form><br />
<br />
<a href="#" onClick="pageload('Sayfalar/pmbox.asp?pm=new','1');return false" class="link1"><img src="../imgs/Mail_add.gif" border="0" /><br />
Yeni Mesaj<br /></a>
</fieldset>
</td>
</tr></table>
<%elseif pm="oku" Then
id=Request.Querystring("id")
if isnumeric(id)=false Then
Response.End
End If
dim pmoku,sql
Set pmoku = Server.CreateObject("ADODB.Recordset")
sql = "select * from pmbox where id="&id&" and alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"'"
pmoku.open sql,conne,1,3
if not pmoku.eof Then
pmoku("durum")=1
pmoku.update
%>
<table width="429" style="position:relative;top:30px">
<tr><td>
<fieldset>
<legend align="center" class="txt">Özel Mesaj Sistemi</legend>
<table>
<tr>
<td class="txt">Gönderen :</td>
<td colspan="2"><% =pmoku("gonderen") %></td>
</tr>
<tr>
<td class="txt">Tarih :</td>
<td colspan="2"><% =pmoku("tarih") %></td>
</tr>
<tr>
<td class="txt">Konu :</td>
<td colspan="2"><% =pmoku("konu") %></td>
</tr>

<tr>
<td class="txt">Mesaj :</td>
<td><% =pmoku("mesaj") %>
</td>
</tr>
<tr >
<td><br /><a href="#" onClick="pageload('Sayfalar/pmbox.asp?pm=new','1');return false"  class="link1"><img src="../imgs/Mail_add.gif" border="0" /><br />
Yeni Mesaj<br /></a></td>
<td ><br /><a href="#" onClick="pageload('Sayfalar/pmbox.asp?pm=cevap&id=<%=id%>','1');return false" class="link1"><img src="../imgs/Mail_reply.gif" width="43" height="28" border="0" /><br>
Cevap Yaz</a>
</td><td width="60" align="center"><br />
<a href="#" onClick="pageload('Sayfalar/pmbox.asp?pm=mailsil&id=<%=id%>','1');return false"  class="link1"><img src="../imgs/Mail_delete.gif" border="0" width="43" height="28"/><br>
Mesajı Sil</a></td>

</tr>
</table>
</fieldset>
</td></tr></table>
<%
pmoku.close
set pmoku=nothing
else
Response.Write "Mesaj bulunamadı."
End If


elseif pm="cevap" Then
id=Request.Querystring("id")
if isnumeric(id)=false Then
Response.Write "Mesaj bulunamadı!"
Response.End
End If
dim pmcvp
set pmcvp=Conne.Execute("select * from pmbox where id='"&id&"' and alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"' ")
if not pmcvp.eof Then

%>
<form action="javascript:pmgonder('cevap')" id="cevap" name="cevap">
<table width="460">
<tr><td width="452">
<fieldset>
<legend align="center"  class="txt">Yeni Mesaj</legend>
<table width="417">
<tr>
<td width="106" class="txt">Gönderen :</td>
<td width="319"><select name="gonderen" class="inpt" style="padding-top:2px">
<%if len(char1)>0 Then
Response.Write "<option value="""&trim(char1)&""" style=""height:15px"">"&trim(char1)&"</option>"
End If%>
<%if len(char2)>0 Then
Response.Write "<option value="""&trim(char2)&""" style=""height:15px"">"&trim(char2)&"</option>"
End If%>
<%if len(char3)>0 Then
Response.Write "<option value="""&trim(char3)&""" style=""height:15px"">"&trim(char3)&"</option>"
End If%>
</select></td>
</tr>
<tr>
<td class="txt">Alıcı:</td>
<td><input type="text" name="alici" class="inpt" value="<% =pmcvp("gonderen") %>"/></td>
</tr>
<tr>
<td class="txt">Konu :</td>
<td><input type="text" name="konu" class="inpt" value="RE: <% =pmcvp("konu") %>" /></td>
</tr>
<tr>
<td class="txt">Mesaj :</td>
<td><textarea cols="40" rows="8" class="inpt2" name="mesaj" style="font-size:11px">
</textarea>
<br />
<br />
</td>
</tr>
<tr>
<td>&nbsp;</td>
<td ><input type="submit" value="" style=" background:url(../imgs/next.gif) right; background-repeat:no-repeat; width:95px; height:25px " /></td>
</tr>
</table>
</fieldset>
</td></tr></table></form>
<%else
Response.Write "Mesaj bulunamadı."
End If
elseif pm="gonder" Then
Response.Charset = "utf-8"
Dim Gonderen,Gonderenkontrol
Gonderen=secur(request.Form("gonderen"))
set gonderenkontrol=Conne.Execute("select * from account_char where strcharid1='"&gonderen&"' or strcharid2='"&gonderen&"' or strcharid3='"&gonderen&"' and straccountid='"&Session("username")&"' ")
if trim(gonderen)=char1 or trim(gonderen)=char2 or trim(gonderen)=char3 Then

if not gonderenkontrol.eof Then
Dim alici,alicikontrol
alici=secur(request.Form("alici"))
set alicikontrol=Conne.Execute("select * from account_char where strcharid1='"&alici&"' or strcharid2='"&alici&"' or strcharid3='"&alici&"'")
If Not alicikontrol.eof Then
Dim pmsayi
Set pmsayi=Conne.Execute("select count(alici) pmsayi from pmbox where alici='"&alici&"'")
If pmsayi("pmsayi")>=5 Then
Response.Write alici&" nickli kişinin posta kutusu doludur !"
Response.End
End If
Dim konu,mesaj
konu=secur(request.Form("konu"))
mesaj=replace(secur(request.Form("mesaj")),vbCrLf,"<br>")
if gonderen="" or alici="" or trim(konu)="" or mesaj="" Then
Response.Write "Boş alan bırakmayınız !"
Response.End
End If

if len(konu)>60 Then
Response.Write "Konu çok uzun"
Response.End
elseif len(mesaj)>2000 Then
Response.Write "Mesaj çok uzun"
Response.End
End If
dim newpm
Set newpm = Server.CreateObject("ADODB.Recordset")
sql = "Select * From pmbox"
newpm.open sql,conne,1,3
newpm.addnew
newpm("gonderen")=gonderen
newpm("alici")=alici
newpm("konu")=konu
newpm("mesaj")=mesaj
newpm("tarih")=now()
newpm("durum")=0
newpm.update
%>
<br>Mesajınız Başarıyla Gönderildi !<br><b>Alıcı:</b> <%=alici%><br><b>Konu:</b> <%=konu %>
<%

newpm.close
set newpm=nothing
else
Response.Write "<br><b>Alıcı bulunamadı.</b>"
End If
else
Response.Write gonderen&" nickli karakter sizin hesabınızda değildir !"
Response.End
End If
else
Response.Write gonderen&" nickli karakter sizin hesabınızda değildir !"
Response.End
End If
elseif pm="new" Then
userid=secur(request.querystring("userid")) %>
<br><br>
<form action="javascript:pmgonder('newpm')" method="post" id="newpm" name="newpm">
<table width="460">
<tr><td width="452">
<fieldset>
<legend align="center" class="txt">Yeni Mesaj</legend>
<table width="417">
<tr>
<td width="106"  class="txt">Gönderen :</td>
<td width="319"><select name="gonderen" class="inpt" style="padding-top:2px">
<%if len(char1)>0 Then
Response.Write "<option value="&char1&">"&trim(char1)&"</option>"
End If%>
<%if len(char2)>0 Then
Response.Write "<option value="&char2&">"&trim(char2)&"</option>"
End If%>
<%if len(char3)>0 Then
Response.Write "<option value="&char3&">"&trim(char3)&"</option>"
End If%>
</select></td>
</tr>
<tr>
<td class="txt">Alıcı:</td>
<td><input type="text" name="alici" class="inpt" value="<%=Server.htmlencode(userid)%>"/></td>
</tr>
<tr>
<td class="txt">Konu :</td>
<td><input type="text" name="konu" class="inpt"/></td>
</tr>
<tr>
<td class="txt">Mesaj :</td>
<td><textarea  cols="40" rows="8" name="mesaj" class="inpt2" style="font-size:11px">
</textarea>
<br />
<br />
</td>
</tr>
<tr>
<td>&nbsp;</td>
<td ><input type="submit" value="" style=" background:url(../imgs/next.gif) right; background-repeat:no-repeat; width:95px; height:25px " /></td>
</tr>
</table>
</fieldset>
</td></tr></table></form>
<%
elseif pm="sil" Then
dim id,pmdel,i,pmsl
id=secur(Request.Querystring("id"))
pmdel=request.form("pmdel")

For i=1 To request.form("pmdel").count
if isnumeric(request.form("pmdel")(i))=false Then
Response.Redirect "pmbox.asp"
Response.End
End If

set pmsl=Conne.Execute("select id, alici from pmbox where id='"&request.form("pmdel")(i)&"' and alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"'")
if not pmsl.eof Then

Conne.Execute("delete pmbox where id='"&request.form("pmdel")(i)&"' ")
else
Response.End
End If
next
Response.Redirect "pmbox.asp"


elseif pm="mailsil" Then
id=Request.Querystring("id")
if isnumeric(id)=true Then
Conne.Execute("delete pmbox where id='"&id&"' and alici='"&char1&"' or alici='"&char2&"' or alici='"&char3&"'")
Response.Redirect "pmbox.asp"
else
Response.Redirect "pmbox.asp"
End If

End If
else
Response.Write "Üye girişi yapınız."
End If%>


</html>
<% else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafından kapatılmıştır.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing
Session.CodePage=1254%>
