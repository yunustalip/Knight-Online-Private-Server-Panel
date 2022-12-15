<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0 
Dim menuayar,siteayar,g,y,r,uye,player,karus,human,karusking,humanking,warhero,delos,delos2,bestkarusplayer,besthumanplayer,bestwarrior,bestrogue,bestpriest,bestmage,onlineu,online
set siteayar=Conne.Execute("select ip,sunucuadi from siteayar")
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Statistics'")
If MenuAyar("PSt")=1 Then

If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Statistics")
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
yn("/Statistics")
End If
%>
<style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
</style><br><center><img src="imgs/istatistik.gif"/></center>
<table width="300" height="260" border="1" align="left" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" style="font-weight:bold;  border: 1px #000000;  border-color:#FFFFFF; background-repeat:no-repeat ; border-collapse:collapse; position:relative;top:20px;left:30px">
  <tr>
    <td colspan="2" align="center"><strong><font size="2"></font></strong></td>
  </tr>
  <tr>
    <td width="110">Server Adý</td>
    <td width="129" align="center"><%=siteayar("SunucuAdi")%></td>
  </tr>
  <tr>
    <td>Server IP </td>
    <td align="center"><%=siteayar("IP")%></td>
  </tr>
  <tr>
    <td>Online Oyuncu </td>
    <td align="center"><%Set onlineu =Conne.Execute("Select Count(strcharid) as toplam From CurrentUser")
online = onlineu("toplam")
onlineu.Close
Set onlineu=Nothing
Response.Write online%> </td>
  </tr>
  <tr>
  <td height="22">Yoðunluk</td>
  <td align="left" height="23" background="imgs/servery.gif" style="position:relative;z-index:0; background-repeat: no-repeat">&nbsp;&nbsp;&nbsp;<%	g="<img src=""imgs/green.gif"" style=""position:relative;top:2px"">&nbsp;"
	y="<img src=""imgs/yellow.gif"" style=""position:relative;top:2px"">&nbsp;"
	r="<img src=""imgs/red.gif"" style=""position:relative;top:2px"">&nbsp;"
	online=300
	if online<=1 Then 
	Response.Write g
	elseif online<=3 Then 
	Response.Write g+g	
	elseif online<=10 Then 
	Response.Write g+g+g
	elseif online<=20 Then 
	Response.Write g+g+g+g
	elseif online<=30 Then 
	Response.Write g+g+g+g+g	
	elseif online<=40 Then 
	Response.Write g+g+g+g+g+y
	elseif online<=50 Then 
	Response.Write g+g+g+g+g+y+y+y
	elseif online<=70 Then 
	Response.Write g+g+g+g+g+y+y+y+y
	elseif online<=90 Then 
	Response.Write g+g+g+g+g+y+y+y+y+y
	elseif online<=100 Then 
	Response.Write g+g+g+g+g+y+y+y+y+y+r+r
	elseif online<=110 Then 
	Response.Write g+g+g+g+g+y+y+y+y+y+r+r+r
	elseif online<=120 Then 
	Response.Write g+g+g+g+g+y+y+y+y+y+r+r+r+r
	elseif online <>150 or online =150  Then 
	Response.Write g+g+g+g+g+y+y+y+y+y+r+r+r+r+r
	End If
	%></td>
  </tr>
  <tr>
    <td>Toplam Hesap </td>
    <td align="center"><%Set uye=Conne.Execute("Select count(StrAccountId) as totaluye From tb_user")
Response.Write uye("totaluye")
uye.close
set uye=nothing
%></td>
  </tr>
  <tr>
    <td>Toplam Karakter: </td>
    <td align="center"><%Set player = Conne.Execute ("Select count(StrUserId) As totalplayer From userdata")
Response.Write player("totalplayer")
player.close
set player=nothing
%></td>
  </tr>
  <tr>
    <td>Toplam Karus </td>
    <td align="center"><% Set karus = Conne.Execute ("Select count(StrUserId) as totalkarus From userdata where nation=1")
Response.Write karus("totalkarus")
karus.close
set karus=nothing
%></td>
  </tr>
  <tr>
    <td>Toplam Human </td>
    <td align="center"><% Set human = Conne.Execute ("Select count(StrUserId) as totalhuman From userdata where nation=2")
Response.Write human("totalhuman")
human.close
set human=nothing%></td>
  </tr>
  <tr>
    <td>Karus Kral </td>
    <td align="center"><% Set karusking = Conne.Execute ("Select strKingName From king_system where bynation=1")
	if not karusking.eof and len(trim(karusking(0)))>0 Then
Response.Write "<img src='imgs/king2.gif'>"&"<br/><font color='#FFCC00'><b>&nbsp;"&karusking("strKingName")&"</b></font>" 
End If
karusking.close
set karusking=nothing%></td>
  </tr>
  <tr>
    <td>Human Kral </td>
    <td align="center"><% Set humanking = Conne.Execute ("Select strKingName From king_system where bynation=2")
	if not humanking.eof and len(trim(humanking(0)))>0 Then
Response.Write "<img src='imgs/king2.gif'>"&"<br/><font color='#FFCC00'><b>&nbsp;"&humanking("strKingName")&"</b></font>" 
End If
humanking.close
set humanking=nothing%></td>
  </tr>
    <tr>
    <td>Savaþ Kahramaný</td>
    <td align="center"><b><% set warhero=Conne.Execute("select bynation,strusername from battle")
			if warhero("bynation")="1" Then 
			Response.Write "<font color='#0033FF'>"&warhero("strusername")&"</font>"
			elseif warhero("bynation")="2" Then
			Response.Write "<font color='#FF0000'>"&warhero("strusername")&"</font>"
			End If
			warhero.close
			set warhero=nothing%></b></td>
  </tr>
  <tr>
    <td>Delos Sahibi</td>
    <td align="center"><b><%
set delos=Conne.Execute("select sMasterKnights from KNIGHTS_SIEGE_WARFARE")
set delos2=Conne.Execute("select idname from knights where idnum = '"&delos("sMasterKnights")&"'")
if not delos2.eof  Then
Response.Write delos2("idname")
End If 
delos.close
set delos=nothing
delos2.close
set delos2=nothing%></b></td>
  </tr>
 </table>
 
<table  width="300" height="130" border="1" align="right" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" style="font-weight:bold;  border: 1px #000000;  border-color:#FFFFFF; background-repeat:no-repeat ; border-collapse:collapse; position:relative;top:20px;right:30px">
   <tr>
    <td colspan="2" align="center"><strong><font size="2"></font></strong></td>
  </tr>
  <tr>
    <td>En Ýyi Karus Oyuncu</td>
    <td align="center"><strong><font color="#0033FF">
<% Set bestkarusplayer = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where nation=1 and authority<>0 and authority<>255 ORDER BY Loyalty DESC")
if not bestkarusplayer.eof Then
Response.Write bestkarusplayer("struserid")
else 
Response.Write "-"
End If
bestkarusplayer.close
set bestkarusplayer=nothing %>
    </font></strong></td>
  </tr>
  <tr>
    <td>En Ýyi Human Oyuncu </td>
    <td align="center"><strong><font color="#FF0000">
<% Set besthumanplayer = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where nation=2 and authority=1 ORDER BY Loyalty DESC")
if not besthumanplayer.eof Then
Response.Write besthumanplayer("struserid")
else
Response.Write "-"
End If
besthumanplayer.close
set besthumanplayer=nothing %>
    </font></strong></td>
  </tr>
	
  <tr>
    <td>En Ýyi Warrior Oyuncu </td>
    <td align="center"><strong>
      <% Set bestwarrior = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where authority=1 and Class=101 or authority=1 and Class=105 or authority=1 and Class=106 or authority=1 and Class=201 or authority=1 and class=205 or authority=1 and  class=206  ORDER BY Loyalty DESC")
if not bestwarrior.eof Then
Response.Write bestwarrior("struserid")
else
Response.Write "-"
End If
bestwarrior.close
set bestwarrior =nothing 
%>
    </strong></td>
  </tr>
    <tr>
    <td>En Ýyi Rogue Oyuncu </td>
    <td align="center"><strong>
      <% Set bestrogue = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where authority=1 and Class=102 or class=107 or class=108 or Class=202 or class=207 or class=208 ORDER BY Loyalty DESC")
if not bestrogue.eof Then
Response.Write bestrogue("struserid")
else
Response.Write "-"
End If
bestrogue.close
set bestrogue=nothing%>
    </strong></td>
  </tr>
    <tr>
    <td>En Ýyi Priest Oyuncu </td>
    <td align="center"><strong>
      <% Set bestpriest = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where authority=1 and Class=104 or Class=111 or Class=112 or Class=204 or class=211 or class=212 ORDER BY Loyalty DESC")
if not bestpriest.eof Then
Response.Write bestpriest("struserid")
else
Response.Write "-"
End If
bestpriest.close
set bestpriest=nothing%>
   </strong></td>
  </tr>
    <tr>
    <td>En Ýyi Mage Oyuncu </td>
    <td align="center"><strong>
      <% Set bestmage = Conne.Execute("SELECT TOP 1 strUserId,nation FROM USERDATA where authority=1 and Class=103 or Class=109 or Class=110 or Class=203 or class=209 or class=210 ORDER BY Loyalty DESC")
if not bestmage.eof Then
Response.Write bestmage("struserid")
else
Response.Write "-"
End If
bestmage.close
set bestmage=nothing%>
    </strong></td>
  </tr>
  </table>
<%
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>