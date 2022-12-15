<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<!--#include file="md5.asp"-->
<%  Response.expires=0 
Dim MenuAyar,ksira,REFERER_URL,s,REFERER_DOMAIN,banlist,banlisttoplam,kalangun,kalansaat
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='BanList'")
If MenuAyar("PSt")=1 Then
If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("/Ban-List")
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

yn("/Ban-List")
End If
 %>
<style type="text/css">
<!--
.style21 {color: #FFFFFF; font-weight:bold}
-->
</style>
<%
Set banlist = Conne.Execute("Select Race,StrUserId,Level,Nation,Class,Loyalty,Authority,yasakgun,bancount,mutecount,yasaksebep From USERDATA Where Authority=255 or Authority=11 or Authority=2")

Set banlisttoplam = Conne.Execute("Select count(*) toplam From USERDATA Where Authority=255 or Authority=11 or Authority=2")

banlisttoplam=banlisttoplam("toplam")
%><br><img src="imgs/bannedlist.gif"><br><br><br>
<b>Toplam Yasaklý Karakter : </b> <%=banlisttoplam%><br />
<table width="650" border="0" align="center">
  <tr>
   
    <td width="125" align="center" background="imgs/menubg.gif" class="style21">Karakter Adý</td>
    <td width="125" align="center" background="imgs/menubg.gif" class="style21">Kalan Süre</td>
    <td width="175" align="center" background="imgs/menubg.gif" class="style21">Yasak Detayý</td>
    <td width="45" align="center" background="imgs/menubg.gif" class="style21">Level</td>
    <td width="45" align="center" background="imgs/menubg.gif" class="style21">Irk</td>
    <td width="111" align="center" background="imgs/menubg.gif" class="style21">Durum</td>
    <td width="120" align="center" background="imgs/menubg.gif" class="style21">NP</td>
    <%if Session("yetki")="1" Then %>
    <td width="60" align="center" background="imgs/menubg.gif" class="style21">Gm Menü</td>
    <%End If%>
  </tr>
  <% if not banlist.eof Then 
  do while not banlist.eof %>
  <tr>
    <td align="center" bgcolor="#F3D78B"><a href="Karakter-Detay/<%=trim(banlist("strUserId"))%>" onclick="pageload('Karakter-Detay/<%=trim(banlist("strUserId"))%>');chngtitle('<%=banlist("strUserId")%> > Karakter Detay');return false" class="link1"><%=banlist("strUserId")%></a> </td>
    <td align="center" bgcolor="#F3D78B"><%
kalangun=datediff("d",now,banlist("yasakgun"))
kalansaat=datediff("h",now,banlist("yasakgun"))
kalandakika=datediff("n",now,banlist("yasakgun"))
If kalangun>0 Then
Response.Write kalangun&" Gün"
ElseIf kalangun=0 And kalansaat>0 Then
Response.Write kalansaat&" Saat"
ElseIf kalangun=0 And kalansaat=0 And kalandakika>0 Then
Response.Write kalandakika&" Dakika"
ElseIf kalangun<=0 And kalansaat<=0 And kalandakika<=0 Then
Response.Write "Ban Açýlmýþtýr. Oyuna Girilebilir."
End If
%></td>
    <td align="center" bgcolor="#F3D78B"><%=banlist("yasaksebep")%></td>
    <td align="center" bgcolor="#F3D78B"><%=banlist("Level")%></td>
    <td align="center" bgcolor="#F3D78B"><% if banlist("Nation")="1" Then 
	Response.Write "<img src='imgs/karuslogo.gif' />"
	elseif banlist("Nation")="2" Then Response.Write "<img src='imgs/elmologo.gif' />"
	End If %></td>
    <td align="center" bgcolor="#F3D78B"><% if banlist("Authority")="255" Then %><img src="imgs/banned.gif"><%elseif banlist("Authority")="11" or banlist("Authority")="2" Then%><img src="imgs/mute.gif"><%End If%></td>
       <td align="center" bgcolor="#F3D78B"><%=banlist("Loyalty")%></td>
       <%if Session("yetki")="1" Then%>
          <td align="center" bgcolor="#F3D78B"><input type=button class="styleform" style=" border-style:solid; background:none; border-color:#FFFFFF" onclick="javascript:gpopup('<%="gmpage/gamem.asp?username="&Session("username")&"&pwd="&md5(Session("pwd"))&"&user=bankaldir&nick="&banlist("struserID")%>')"  value="<%if banlist("Authority")="255" Then
Response.Write "Ban Kaldýr"
else
Response.Write "Mute Kaldýr"
End If%>"></td>
	   <%End If%>
  </tr>
  <%
  banlist.MoveNext
  Loop
  %>
  </table>
  <% else
  Response.Write("<table><tr><td>Yasaklý Kullanýcý Bulunmamaktadýr.</td></tr></table>")
  End If 
  else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing %>