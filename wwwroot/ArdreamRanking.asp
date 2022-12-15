<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%Response.expires=0 
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,link,gelenlink_bol,tp,d1,d2,d3,d4,siralama,siralamatur,siteayr,listeleme,tur,userrank,userranktp,toplamchar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='ArdreamRanking'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("/Ardream-Ranking")
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

yn("/Ardream-Ranking")
End If


link = Session("sayfa")
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
End If

Set siteayr=Conne.Execute("select kulsiralama from siteayar")
ksira=siteayr("kulsiralama")

siralama = secur(d1)
siralamatur = secur(d2)
listeleme=secur(d3)

if siralama="" Then
siralama="ortak"
elseif siralama<>"karus" and siralama<>"elmorad" Then
siralama="ortak"
End If

if siralamatur="Warrior" and siralama="karus" Then
tur="class=101 and authority=1 or class=105 and authority=1 or class=106 and authority=1"
elseif siralamatur="Warrior" and siralama="elmorad" Then
tur="class=201 and authority=1 or class=205 and authority=1 or class=206 and authority=1"
elseif siralamatur="Warrior" and siralama="ortak" Then
tur="class=101 and authority=1 or class=105 and authority=1 or class=106 and authority=1 or class=201 and authority=1 or class=205 and authority=1 or class=206 and authority=1"

elseif siralamatur="Rogue" and siralama="karus" Then
tur="class=102 and authority=1 or class=107 and authority=1 or class=108 and authority=1" 
elseif siralamatur="Rogue" and siralama="elmorad" Then
tur="class=202 and authority=1 or class=207 and authority=1 or class=208 and authority=1"
elseif siralamatur="Rogue" and siralama="ortak" Then
tur="class=102 and authority=1 or class=107 and authority=1 or class=108 and authority=1 or class=202 and authority=1 or class=207 and authority=1 or class=208 and authority=1"

elseif siralamatur="Priest" and siralama="karus" Then
tur="class=104 and authority=1 or class=111 and authority=1 or class=112 and authority=1"
elseif siralamatur="Priest" and siralama="elmorad" Then
tur="class=204 and authority=1 or class=211 and authority=1 or class=212 and authority=1"
elseif siralamatur="Priest" and siralama="ortak" Then
tur="class=104 and authority=1 or class=111 and authority=1 or class=112 and authority=1 or class=204 and authority=1 or class=211 and authority=1 or class=212 and authority=1"

elseif siralamatur="Mage" and siralama="karus" Then
tur="class=103 and authority=1 or class=109 and authority=1 or class=110 and authority=1"
elseif siralamatur="Mage" and siralama="elmorad" Then
tur="class=203 and authority=1 or class=209 and authority=1 or class=210 and authority=1"
elseif siralamatur="Mage" and siralama="ortak" Then
tur="class=103 and authority=1 or class=109 and authority=1 or class=110 and authority=1 or class=203 and authority=1 or class=209 and authority=1 or class=210 and authority=1"

elseif siralama="karus" and siralamatur="" Then
tur="authority=1 and nation=1"
elseif siralama="elmorad" and siralamatur="" Then
tur="authority=1 and nation=2"
elseif siralama="ortak" and siralamatur="" Then
tur="authority=1"
else
tur="authority=1"
End If


if siralama="karus" Then
set userrank=Conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where "&tur&" and level<=59 ORDer By Loyalty DESC,Level DESC,struserid asc")
Set userranktp=Conne.Execute("Select count(*) toplam From USERDATA where "&tur&" and level<=59")

toplamchar=userranktp("toplam")

elseif siralama="elmorad" Then
set userrank=Conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where "&tur&" and level<=59 ORDer By Loyalty DESC,Level DESC,struserid asc")
Set userranktp=Conne.Execute("Select count(*) toplam From USERDATA where "&tur&" and level<=59")

toplamchar=userranktp("toplam")
else
set userrank=Conne.Execute("Select Top "&ksira&" Race,StrUserId,Level,Nation,Class,Loyalty,Knights From USERDATA where "&tur&" and level<=59  ORDer By Loyalty DESC,Level DESC,struserid asc")
Set userranktp=Conne.Execute("Select count(*) toplam From USERDATA where "&tur&" and level<=59")
toplamchar=userranktp("toplam")

End If

If siralama="karus" Then
Response.Write "<br><center><img src=""imgs/karusardreamnp.gif"" alt="""">"
elseif siralama="elmorad" Then
Response.Write "<br><center><img src=""imgs/humanardreamnp.gif"" alt="""">"
else
Response.Write "<br><center><img src=""imgs/ardreamnp.gif"" alt="""">"
End If%>

<style type="text/css">
.nosort {color: #FFFFFF; font-weight:bold}
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
a{
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration:none;
}
</style>
<br><br><br><br>
<b><a href="/Ardream-Ranking" onclick="javascript:pageload('/Ardream-Ranking');return false" class="link1"><img src="imgs/ortak.gif"  border="0" align="texttop">&nbsp;Ortak Sýralama</a>&nbsp;-&nbsp;<a href="/Ardream-Ranking/karus" onclick="javascript:pageload('/Ardream-Ranking/karus');return false" class="link1"><img src="imgs/karuslogo.gif" border="0" align="texttop">&nbsp;Karus Sýralamasý</a>&nbsp;-&nbsp;<a href="/Ardream-Ranking/elmorad" onclick="javascript:pageload('/Ardream-Ranking/elmorad');return false" class="link1"><img src="imgs/elmologo.gif" border="0" align="absmiddle">&nbsp;Human Sýralamasý</a></b>
<br>
<b>Toplam Karakter : </b> <%=toplamchar%><br />
<table width="569" border="0">
  <tr>
    <td align="center" background="imgs/menubg.gif" class="nosort">Sýra </td>
    <td width="150" align="center" background="imgs/menubg.gif" class="nosort">Karakter Adý </td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Level</td>
    <td width="39" align="center" background="imgs/menubg.gif" class="nosort">Irk</td>
    <td width="90" align="center" background="imgs/menubg.gif" class="nosort">Tür</td>
    <td width="140" align="center" background="imgs/menubg.gif" class="nosort">Clan</td>
    <td width="52" align="center" background="imgs/menubg.gif" class="nosort">NP</td>
  </tr>
<%Dim Clan,clanid,idnum,sira,style
if not userrank.eof Then
do while not userrank.eof
if not userrank("Knights")="0" Then
Set clan=Conne.Execute("Select IDNum,IDName From KNIGHTS where IDNum="&userrank("Knights")&"")

if not clan.eof Then

clanid=clan("idname")
idnum=clan("idnum")
else

clanid=""
idnum=""

End If

else
clanid=""
idnum=""
End If

sira=sira+1

%>
<tr bgcolor="#F3D78B" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
	<td align="center"><% Response.Write (sira) %></td>
   	
    <td align="center" id="<%=trim(userrank("strUserId"))%>"><a href="Karakter-Detay/<%=userrank("strUserId")%>" style="display:block" onclick="pageload('Karakter-Detay/<%=userrank("strUserId")%>');return false" ><font style="<%=sirarenk(sira)%>"><%=trim(userrank("strUserId"))%></font></a></td>
    <td align="center"><%=userrank("Level")%></td>
    <td align="center"><% nation(userrank("Nation"))%></td>
    <td align="center"><a href="/Ardream-Ranking/<%Response.Write siralama&"/"
cla(userrank("Class")) %>" onclick="javascript:pageload('/Ardream-Ranking/<%Response.Write siralama&"/"
Response.Write cla(userrank("Class")) %>');return false" class="link1"><%=cla(userrank("Class"))%></a></td>
    <td align="center"><a href="#" onclick="pageload('sayfalar/showclan.asp?goster=<%=idnum%>');return false;"><%=clanid%></a></td>
    <td align="center"><%=userrank("Loyalty")%></td>
  </tr>
<% userrank.MoveNext
  Loop
userrank.close
set userrank=nothing
userranktp.close
set userranktp=nothing

  %><tr><td colspan="8" align="center">Ilk <%=ksira%> karakter gösteriliyor.<br />
  <br />
  Not: Anlýk Np sýralamasýdýr. Oyun içi semboller resetlerde güncellenir. </td>
</tr>
  </table>
  <%else Response.Write("<table><tr><td>Karakter Bulunmamaktadýr.</td></tr></table>")
  	End If %>
    </center>
<%

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing
%>