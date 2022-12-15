<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<%
sayfa=Request.Querystring("sayfa")
select case sayfa
case "1" %>
<form name="aramamotoru" method="post" action="itemara.asp?sayfa=2">
<center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="348" id="AutoNumber1" >
    <tr>
      <td height="50" width="7" rowspan="3">&nbsp;</td>
      <td height="30" width="319" >&nbsp;Aranacak Itemi Ismini Girin :</td>
    </tr>
    <tr>
      <td height="29" width="319" >
      &nbsp;<input size="30" name="aranan" class="inputtxt">
      <input type="submit" name="ok" class="inputtxt" value="  - Ara -  "></td>
    </tr>
  </table>
  </center>
 </form>
<% case "2"
aranan = request("aranan")

if aranan = "" Then
	Response.Write "<b>Lütfen Boþ Alan Býrakmayýnýz!<b><br><br>"
	Response.Write "<INPUT  TYPE=""button"" VALUE=""  << Geri  "" onClick=""history.go(-1)"">"
Else

Set itemara = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from item where strname like '%"&aranan&"%' order by strname"
itemara.Open SQL,conne,1,3

Set itemara2 = Server.CreateObject("ADODB.RecordSet")
SQL = "Select count(*) toplam from item where strname like '%"&aranan&"%'"
itemara2.Open SQL,conne,1,3

Toplam_sonuc = itemara2("toplam")
%>
<center>
<table border="0" width="500" height="25" style="border-collapse: collapse" cellpadding="0" cellspacing="0" bgcolor="#ffcc00">
	<tr>
  <td height="25" background="imajlar/sol_bar.gif" width="14"></td>
  <td height="25" background="imajlar/orta_bar.gif" width="620">Toplam 
  Sonuç Sayýsý : <b><%=Toplam_sonuc%></b></td>
  <td height="25" background="imajlar/sag_bar.gif" width="15">&nbsp;</td></tr></table>
<br/></center>
<%
if itemara.Eof Then
	Response.Write "<b>Sonuç Bulunamadý!<b><br><br>"
	Response.Write "<INPUT class=""buton"" TYPE=""button"" VALUE=""  << Geri  "" onClick=""history.go(-1)"">"
Else
%>
<center>
<table border="0" width="500" height="26" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
<tr>
  <td height="25" background="imajlar/sol_bar.gif" colspan="2"></td>
  <td height="25" background="imajlar/orta_bar.gif" width="250" align="center">
  <p align="left">Item Ismi</td>
  <td height="25" background="imajlar/orta_bar.gif" width="200" align="center">
  Item No :</td>
  
  <td height="25" background="imajlar/sag_bar.gif" colspan="2">&nbsp;</td>
 </tr>
    <%  git = Request.Querystring("git")
	if git="" Then 
	git = 1
	End If
	itemara.pagesize = 50
	ilkkayit=50*(git-1)
	if Toplam_sonuc =<ilkkayit Then
	Response.Write"sayfa bulunamadý"
	else
	itemara.move(ilkkayit)
	
	for i=1 to 50
	if itemara.eof Then exit for%>
 <tr>
  <td height="26" width="7"></td>
  <td height="26" width="7" bgcolor="#FF6600">&nbsp;</td>
  <td width="270" height="26" bgcolor="#FF9900" class="style5">&nbsp;<img src="imajlar/nokta.gif">&nbsp;<b><%=itemara("strname")%></b></td>
  <td width="340" height="26" align="center" bgcolor="#FF9900" class="style5"><b><%=Left(itemara("num"),40)%></b></td>
  <td height="26" width="10" bgcolor="#FF6600">&nbsp;</td>
  <td height="26" width="5">&nbsp;</td>
 </tr>
 <%
 itemara.MoveNext
 next
 %>
 <tr>
  <td height="25" background="imajlar/sol_bar.gif" colspan="2"></td>
  <td height="25" background="imajlar/orta_bar.gif" colspan="3">Sayfalar : 
    <%
        for y=1 to ((Toplam_sonuc-(Toplam_sonuc Mod 50))/50)+1
		if git=y Then
		Response.Write y
		else
		Response.Write "<b> <a href=""default.asp?w8=produces&pro=itemsearchfound&git="&y&"&aranan="&aranan&""">"&y&"</a></b>"
		End If
		next
		%></td>
  <td height="25" background="imajlar/sag_bar.gif" width="5" colspan="2">&nbsp;</td>
 </tr>
</table>
</center>
<%End If
End If
End If
end select %>
<% End If %>
