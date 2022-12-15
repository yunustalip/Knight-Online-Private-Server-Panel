<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<% Response.expires=0 
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,link,gelenlink_bol,d1,d2,d3,d4,tp,clanid,clantp,toplamclan,sira,humanclan,karusclan,csira,sirala
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='ClanNpDonateRanking'")
If MenuAyar("PSt")=1 Then

If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("/Clan-Np-Donate-Ranking")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("script_name")
if InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
else
REFERER_DOMAIN = left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If 

if REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
else

yn("/Clan-Np-Donate-Ranking")
End If


link = Session("Sayfa")
gelenlink_bol = split(link, "/")
tp=ubound(gelenlink_bol)

if tp=4 Then
d1=gelenlink_bol(4)
elseif tp=5 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
elseif tp=6 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
elseif tp=7 Then
d1=gelenlink_bol(4)
d2=gelenlink_bol(5)
d3=gelenlink_bol(6)
d4=gelenlink_bol(7)
End If %>
<style type="text/css">
body,td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}

</style><center>
<br><img src="imgs/clannpdonate.gif"><br>
<br><br><%
set clanid=Conne.Execute("select k.flag,k.IDNum, k.IDName, k.Members, k.Chief, k.Nation,k.ranking,k.points,k.scape,sum(n.np) as np from knights k, npdonate n where n.clan=k.idnum group by k.idname,idnum,k.IDNum, k.IDName, k.Members, k.Chief, k.Nation,k.ranking,k.points,k.scape,k.flag order by sum(np) desc")

Set clantp = Conne.Execute("Select count(distinct clan) toplamclan From npdonate")
toplamclan=clantp("toplamclan")
%><b><img src="imgs/ilk5_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Ilk 5 Clan(Yanan Kolluk)&nbsp;&nbsp;&nbsp;&nbsp;<img src="imgs/ust_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Üst Clanlar (Pelerinli)&nbsp;&nbsp;&nbsp;&nbsp;<img src="imgs/alt_grade5.gif" width="20" height="20" align="absmiddle">&nbsp;Alt Clanlar(Pelerinsiz)</b><br />
<br><b>Toplam Clan : </b> <%=toplamclan%><br />
<table width="556" border="0" align="center">
  <tr>
	<td width="30" height="16" align="center" background="imgs/menubg.gif"><span class="style1">Sýra </span></td>
	<td width="184" align="center" background="imgs/menubg.gif"><span class="style1">Clan Adý</span></td>	
	<td width="139" align="center" background="imgs/menubg.gif"><span class="style1">NP</span></td>
	<td width="100" align="center" background="imgs/menubg.gif"><span class="style1">Grade</span></td>
	<td width="98" align="center" background="imgs/menubg.gif"><span class="style1">Üye Sayýsý</span></td>
	<td width="83" align="center" background="imgs/menubg.gif"><span class="style1">Irk</span></td>
  </tr>
<%s=1
if not clanid.eof Then
sira=1
humanclan=0
karusclan=0
for sirala=1 to csira
if clanid.eof Then
exit for
End If

clanids = clanid("idnum")
name = clanid("idname")
members = clanid("members")
clannation = clanid("nation")
totalnp = clanid("np")
ranking = clanid("ranking")
points =(clanid("np"))
cape= clanid("scape")
flag= clanid("flag")

if clannation="1" Then
karusclan=karusclan+1
elseif clannation="2" Then
humanclan=humanclan+1
End If

if points<72000 Then
grade=5
elseif points<144000 Then
grade=4
elseif points<360000 Then
grade=3
elseif points<720000 Then
grade=2
elseif points>=720000 Then
grade=1
End If

if clannation="1" and karusclan<6 or clannation="2" and humanclan<6 Then
clangrade="ilk"
else
clangrade="diger"
End If

if grade=5 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade5.gif height=20 width=20>"
elseif grade=4 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade4.gif height=20 width=20>"
elseif grade=3 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade3.gif height=20 width=20>"
elseif grade=2 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade2.gif height=20 width=20>"
elseif grade=1 and flag="2" and clangrade="diger" Then
derece="<img src=imgs/ust_grade1.gif height=20 width=20>"
End If

if grade=5 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade5.gif height=20 width=20>"
elseif grade=4 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade4.gif height=20 width=20>"
elseif grade=3 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade3.gif height=20 width=20>"
elseif grade=2 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade2.gif height=20 width=20>"
elseif grade=1 and flag="1" and clangrade="diger" Then
derece="<img src=imgs/alt_grade1.gif height=20 width=20>"

End If


if grade=5 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade5.gif height=20 width=20>"
elseif grade=4 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade4.gif height=20 width=20>"
elseif grade=3 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade3.gif height=20 width=20>"
elseif grade=2 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade2.gif height=20 width=20>"
elseif grade=1 and clangrade="ilk" Then
derece="<img src=imgs/ilk5_grade1.gif height=20 width=20>"
End If

%>
  <tr bgcolor="#F3D78B" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
	<td align="center"><%Response.Write(sira)%></td>
	<td align="center"><a href="/Clan-Np-Detay/<%=trim(name)&","&clanids%>" style="display:block"  onclick="pageload('/Clan-Np-Detay/<%=name&","&clanids%>');return false"><%=name%></a></td>
	<td align="center"><%=ayir(totalnp)%></td>
	<td align="center"><%=derece%></td>
	<td align="center"><%=members%></td>
	<td align="center"><%=nation(clannation)%></td>
  </tr>
  <%
  sira=sira+1
  clanid.MoveNext
  next
  clanid.close
  set clanid=nothing
  clantp.close
  set clantp=nothing
  else
   Response.Write("<table><tr><td align=""center"">Kurulu Clan Bulunumadý. </td></tr></table>")
  End If %>
  <tr><td colspan="6" align="center">Ilk <%=csira%> Clan Gösteriliyor.<br />
<br />Not: Anlýk Clan sýralamasýdýr.Oyun içi gradeler resetlerde güncellenir.
  <td></tr>
</table>
</center>
<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>