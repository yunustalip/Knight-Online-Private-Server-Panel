<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%if menuayar("clannpsiralamasi")=1 Then

lnk=secur(Request.Querystring)
linkp=split(lnk,",")
if instr(lnk,",")>0 Then
goster=linkp(1)
else

End If

if isnumeric(goster)=false or goster="" Then
Response.Write("<br><br><b>Clan Bulunamadý.</b>")
Response.End
End If


Set clan = Conne.Execute("Select IDName,Chief,IDNum,Chief,ViceChief_1,ViceChief_2,ViceChief_3,Members,Nation,points,scape,ranking,flag,createtime From KNIGHTS Where IDNum = "&goster)

if not clan.eof Then
Set usert = Conne.Execute("Select n.np,n.clan,n.userid,u.struserid,u.knights,u.level From npdonate n, userdata u Where clan='"&goster&"' and n.userid=u.struserid")

set totalnp=Conne.Execute("select sum(np) from npdonate where clan='"&clan("idnum")&"' ")


points=totalnp(0)
cape=clan("scape")
ranking=clan("ranking")
flag=clan("flag")

if cape="-1" Then
pelerin="Yok"
else
pelerin="Var"
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

if ranking=1 or ranking=2 or ranking=3 or ranking=4 or ranking=5 Then
clangrade="ilk"
else
clangrade="diger"
End If

if grade=5 and flag="2" Then
derece="<img src=../imgs/ust_grade5.gif height=20 width=20>"
elseif grade=4 and flag="2" Then
derece="<img src=../imgs/ust_grade4.gif height=20 width=20>"
elseif grade=3 and flag="2" Then
derece="<img src=../imgs/ust_grade3.gif height=20 width=20>"
elseif grade=2 and flag="2" Then
derece="<img src=../imgs/ust_grade2.gif height=20 width=20>"
elseif grade=1 and flag="2" Then
derece="<img src=../imgs/ust_grade1.gif height=20 width=20>"
End If

if grade=5 and flag="1" Then
derece="<img src=../imgs/alt_grade5.gif height=20 width=20>"
elseif grade=4 and flag="1" Then
derece="<img src=../imgs/alt_grade4.gif height=20 width=20>"
elseif grade=3 and flag="1" Then
derece="<img src=../imgs/alt_grade3.gif height=20 width=20>"
elseif grade=2 and flag="1" Then
derece="<img src=../imgs/alt_grade2.gif height=20 width=20>"
elseif grade=1 and flag="1" Then
derece="<img src=../imgs/alt_grade1.gif height=20 width=20>"

End If


if grade=5 and clangrade="ilk" Then
derece="<img src=../imgs/ilk5_grade5.gif height=20 width=20>"
elseif grade=4 and clangrade="ilk" Then
derece="<img src=../imgs/ilk5_grade4.gif height=20 width=20>"
elseif grade=3 and clangrade="ilk" Then
derece="<img src=../imgs/ilk5_grade3.gif height=20 width=20>"
elseif grade=2 and clangrade="ilk" Then
derece="<img src=../imgs/ilk5_grade2.gif height=20 width=20>"
elseif grade=1 and clangrade="ilk" Then
derece="<img src=../imgs/ilk5_grade1.gif height=20 width=20>"
End If

End If

if not cape=-1 Then
	if len(cape)=3 Then
	capem=left(cape,1)
	pelerinm="../imgs/cape/cloak_m_0"&capem&".gif"
	cape=mid(cape,2,2)
	End If
	if len(cape)=1 Then
	cape=0&cape
	End If
End If
%><br>
<img src="imgs/clandetay.gif">

<style type="text/css">
<!--
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 16px;
	color:#FF0000;
}
.style22 {
	color: #CC0000;
	font-weight: bold;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
.stil1siralama {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	color: #990000;
	font-weight: bold;
}
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}

-->
</style>

<% if not clan.eof Then %>
<table width="550"  border="0" align="center" <%if len(capem)>0 Then Response.Write "style=""position:relative;left:40"""%>>

 <tr>
     <td align="center" valign="top" ><% if clan("nation")="2" Then
	 Response.Write "<img src=""../imgs/elmoc.gif"" width=""100"" height=""69"">"
	 elseif clan("nation")="1" Then
	 Response.Write "<img src=""../imgs/karusc.gif"" width=""100"" height=""69"">"
	 else
	 End If%>	 </td>
    <td width="235" valign="top">
   	  <table width="200" border="0" align="center">
  <tr>
    <td align="left" bgcolor="#FF9900" class="style1">Clan Adý :</td>
    <td align="center" bgcolor="#FFCC66" class="stil1siralama"><%=clan("IDName")%></td>
    </tr>
  <tr>
    <td align="left" bgcolor="#FF9900" class="style1">Toplam Np :</td>
    <td align="center" bgcolor="#FFCC66"><%=ayir(totalnp(0))%></td>
    </tr>
  <tr>
    <td align="left" bgcolor="#FF9900" class="style1">Pelerin :</td>
    <td align="center" bgcolor="#FFCC66"><%=pelerin%></td>
  </tr>
  <tr>
    <td align="left" bgcolor="#FF9900" class="style1">Grade :</td>
    <td align="center" bgcolor="#FFCC66"><%=derece%></td>
  </tr>
</table></td>
    
    <td valign="top" ><table width="219" border="0">
      <tr>
        <td width="92" align="left" bgcolor="#FF9900" class="style1">Lider</td>
        <td width="111" bgcolor="#FFCC66"><%=clan("Chief")%></td>
      </tr>
      <tr>
        <td align="left" bgcolor="#FF9900" class="style1">1.Asistan</td>
        <td bgcolor="#FFCC66"><%=clan("ViceChief_1")%></td>
      </tr>
      <tr>
        <td align="left" bgcolor="#FF9900" class="style1">2.Asistan</td>
        <td bgcolor="#FFCC66"><%=clan("ViceChief_2")%></td>
      </tr>
      <tr>
        <td align="left" bgcolor="#FF9900" class="style1">3.Asistan</td>
        <td bgcolor="#FFCC66"><%=clan("ViceChief_3")%></td>
      </tr>
      <tr>
        <td bgcolor="#FF9900" class="style1">Kuruluþ Tarihi</td>
        <td bgcolor="#FFCC66"><%=left(clan("createtime"),10)%></td>
      </tr>
    </table></td>
    <td align="center" valign="top" >
<%if not cape=-1 Then
	pelerin="../imgs/cape/cloak_c_"&cape&".gif"
	Response.Write "<img src="&pelerin&" width=""96"" height=""96"">"
	if len(capem)>0 Then
	Response.Write "<img src="&pelerinm&" width=""96"" height=""96"" style=""position:relative;left:-96"">"
	End If
	End If%></td>
 </tr>
</table>
<table width="465" border="0" align="center">
  <tr>
    <td align="center"><strong>Clan Üyeleri (Toplam Üye: 
        <% =clan("members") %> ) </strong></td>
</tr></table>
<table width="550" border="0" align="center">
  <tr>
    <td width="194" align="center" bgcolor="#FF6600"><strong class="style1">Karakter Adý </strong></td>
    <td width="58" align="center" bgcolor="#FF6600"><strong class="style1">Level</strong></td>
	    <td width="199" align="center" bgcolor="#FF6600"><strong class="style1">Baðýþlanan NP</strong></td>
		<td width="75" align="center" bgcolor="#FF6600"><strong class="style1">Durum</strong></td>
  </tr>

  <% if not usert.eof Then 
  do while not usert.eof
  set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&usert("strUserId")&"'")%>
   <tr bgcolor="#FFCC00">
    <td align="center"><a href="../Karakter-Detay/<%=usert("strUserId")%>" onclick="javascript:pageload('Karakter-Detay/<%=usert("strUserId")%>');return false"><%=usert("strUserId")%></a></td>
    <td align="center"><%=usert("Level")%></td>
	<td align="center"><%=ayir(usert("np"))%></td>
	    <td align="center"><%if not onlinek.eof Then
		Response.Write "<font color='#FF0000'>Oyunda !</font>"
		elseif onlinek.eof Then
		Response.Write "<font color='#666666'>Çevrimdýþý</font>" 
		End If%></td>
  </tr><%
  usert.Movenext
  if not usert.eof Then 
   set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&usert("strUserId")&"'")%>
     <tr bgcolor="#FFCC66">
    <td align="center" bgcolor="#FF9900"><a href="../Karakter-Detay/<%=usert("strUserId")%>" onclick="javascript:pageload('Karakter-Detay/<%=usert("strUserId")%>');return false"><%=usert("strUserId")%></a></td>
    <td align="center" bgcolor="#FF9900"><%=usert("Level")%></td>
	<td align="center" bgcolor="#FF9900"><%=ayir(usert("np"))%></td>
		    <td align="center" bgcolor="#FF9900"><%if not onlinek.eof Then
		Response.Write "<font color='#FF0000'>Oyunda !</font>"
		elseif onlinek.eof Then
		Response.Write "<font color='#666666'>Çevrimdýþý</font>" 
		End If%></td>
  </tr>
   <%
  usert.Movenext
  End If
  Loop
 else 
 Response.Write("<tr><td>Clanda Üye Yok</td></tr>")
  End If %>
</table>
<% else Response.Write"Böyle Bir Clan Bulunmamaktadýr."
 End If
 else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if
MenuAyar.Close
Set MenuAyar=Nothing%>