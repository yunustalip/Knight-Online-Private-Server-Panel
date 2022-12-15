<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<!--#include file="md5.asp"-->
<%Dim MenuAyar,ksira,REFERER_DOMAIN,REFERER_URL,s,online
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Online'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("/User-Ranking")
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

yn("/User-Ranking")
End If
set online=Conne.Execute("Select CURRENTUSER.strCharID,CURRENTUSER.sure,CURRENTUSER.np,CURRENTUSER.np2,USERDATA.strUserId,USERDATA.Level,USERDATA.Nation,USERDATA.Loyalty,USERDATA.Knights,USERDATA.Zone,USERDATA.class,USERDATA.gunluknp1,USERDATA.gunluknp2 From USERDATA, CURRENTUSER where CURRENTUSER.strCharID=USERDATA.strUserId")
%><style>.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}

</style>
<br /><center><img src="imgs/onlineuserlist.gif" /></center>
<br><br>
<table width="740" border="0" align="center">
  <tr>
    <td width="120" height="16" align="center" background="imgs/menubg.gif"><span class="style1">Karakter Adý </span></td>
    <td align="center" background="imgs/menubg.gif"><span class="style1">Level</span></td>
    <td align="center" background="imgs/menubg.gif"><span class="style1">Irk</span></td>
    <td align="center" background="imgs/menubg.gif"><span class="style1">Tür</span></td>
    <td width="80"  align="center" background="imgs/menubg.gif"><span class="style1">NP</span></td>
    <td width="80" align="center" background="imgs/menubg.gif"><span class="style1">Günlük NP</span></td>
    <td width="90" align="center" background="imgs/menubg.gif"><span class="style1">Clan</span></td>
    <td width="100" align="center" background="imgs/menubg.gif"><span class="style1">Bulunduðu Yer</span></td>
    <td width="70" align="center" background="imgs/menubg.gif"><span class="style1">Online Süre</span></td>
	<% if Session("yetki")="1" Then 
	Response.Write "<td width=""40"" align=""center"" background=""imgs/menubg.gif""><span class=""style1"">Gm</span></td>"
	End If%>
  </tr>
<% if not online.eof Then
do while not online.eof 
Set clan =Conne.Execute("Select IDNum,IDName From KNIGHTS Where IDNum='"&online("Knights")&"'")
%>
  <tr bgcolor="#F3D78B"  onmouseover="this.style.background='#E3C06F'" onmouseout="this.style.background='#F3D78B'">
    <td align="center"><a href="Karakter-Detay/<%=online("strCharID")%>" style="display:block" onclick="javascript:pageload('Karakter-Detay/<%=online("strCharID")%>');return false" class="link1"><%=online("strCharID")%></a></td>
    <td align="center"><%=online("Level")%></td>
    <td align="center"><% nation(online("Nation")) %></td>
    <td align="center"><%=cla(online("class"))%></td>
    <td align="center"><%=online("Loyalty")%></td>
    <td align="center"><%=online("gunluknp1")+(online("loyalty")-online("np"))%></td>
    <td align="center"><% if not clan.eof Then 
	Response.Write "<a href=""#"" onclick=""pageload('Clan-Detay/,"&clan("idnum")&"');return false;"" style=""display:block"" class=""link1"">"&clan("IDName")&"</a>"
	else 
	Response.Write ""
	End If %></td>
    <td align="center">
<% if  online("Zone")="21" Then 
Response.Write "Moradon"
elseif online("Zone")="1" Then 
Response.Write "Luferson Castle"
elseif online("Zone")="2" Then 
Response.Write "Elmorad Castle"
elseif online("Zone")="201" Then 
Response.Write "Colony Zone"
elseif online("Zone")="202" Then 
Response.Write "Ardream"
elseif online("Zone")="30" Then 
Response.Write "Delos"
elseif online("Zone")="48" Then 
Response.Write "Arena"
elseif online("Zone")="101" Then 
Response.Write "Lunar War"
elseif online("Zone")="102" Then 
Response.Write "Dark Lunar War"
elseif online("Zone")="103" or online("Zone")="111" Then 
Response.Write "War Zone"
elseif online("Zone")="11" Then 
Response.Write "Karus Eslant"
elseif online("Zone")="12" Then 
Response.Write "El Morad Eslant"
elseif online("Zone")="31" Then 
Response.Write "Bi-Frost"
elseif online("Zone")="51" or online("Zone")="52" or online("Zone")="53" or online("Zone")="54" or online("Zone")="55" Then 
Response.Write "Forgetten Temple Zone"
elseif online("Zone")="32" Then 
Response.Write "Hell Abyss"
elseif online("Zone")="33" Then 
Response.Write "Isiloon Floor"
End If%></td>
<td align="center">
<% dk=round(datediff("s",online("sure"),now )/60)
a=dk mod 60
if dk mod 60=0 Then
Response.Write round(dk/60)&" Saat"
elseif dk>60 Then
Response.Write round(dk/60) &" Saat "&a&" Dk."
else
Response.Write dk&" Dk."
End If
%></td>
<%if Session("yetki")="1" Then%>
<td align="center"><input type=button class="styleform" style=" border-style:solid; background:none; border-color:#FFFFFF" onclick="gpopup('GmPage/gamem.asp?user=ban&nick=<%=online("strCharID")%>')"  value="Banla/Dc"></td>
<%End If%>
  </tr>
<%clan.close
set clan=nothing
online.MoveNext
Loop
else %>
<tr><td align="center">Online Kullanýcý Yok</td></tr>
<% End If 
Response.Write "</table>"
online.close
set online=nothing
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If

MenuAyar.Close
Set MenuAyar=Nothing %>