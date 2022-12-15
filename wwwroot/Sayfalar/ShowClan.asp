<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="../guvenlik.asp"-->
<%response.expires=0

Dim MenuAyar
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='ShowClan'")
If MenuAyar("PSt")=1 Then
Dim lnk,linkp,goster
lnk=Session("sayfa")
linkp=split(lnk,",")
if instr(lnk,",")>0 Then
goster=linkp(1)
else

End If

if isnumeric(goster)=false or goster="" Then
Response.Redirect("../Clan-Ranking")
Response.End
End If

Dim Clan,usert,totalnp,points
Set Clan = Conne.Execute("Select IDName,Chief,IDNum,Chief,ViceChief_1,ViceChief_2,ViceChief_3,Members,Nation,points,scape,ranking,flag,createtime From KNIGHTS Where IDNum = "&goster&"")

if not clan.eof Then
Set usert = Conne.Execute("Select strUserId,Knights,Loyalty,Level From USERDATA  Where Knights = "&clan("IDNum")&" and authority<>255 order by loyalty")
set totalnp=Conne.Execute("select sum(loyalty) from userdata where Knights = "&clan("IDNum")&" and authority<>0 and authority<>255 having sum(loyalty)>0")

if totalnp.eof Then
points=0
else
points=totalnp(0)
End If

Dim cape,ranking,flag,pelerin,grade,clangrade,derece,capem,pelerinm
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
derece="<img src=../imgs/ust_grade_5.bmp>"
elseif grade=4 and flag="2" Then
derece="<img src=../imgs/ust_grade_4.bmp>"
elseif grade=3 and flag="2" Then
derece="<img src=../imgs/ust_grade_3.bmp>"
elseif grade=2 and flag="2" Then
derece="<img src=../imgs/ust_grade_2.bmp>"
elseif grade=1 and flag="2" Then
derece="<img src=../imgs/ust_grade_1.bmp>"
End If

if grade=5 and flag="1" Then
derece="<img src=../imgs/alt_grade_5.bmp>"
elseif grade=4 and flag="1" Then
derece="<img src=../imgs/alt_grade_4.bmp>"
elseif grade=3 and flag="1" Then
derece="<img src=../imgs/alt_grade_3.bmp>"
elseif grade=2 and flag="1" Then
derece="<img src=../imgs/alt_grade_2.bmp>"
elseif grade=1 and flag="1" Then
derece="<img src=../imgs/alt_grade_1.bmp>"

End If


if grade=5 and clangrade="ilk" Then
derece="<img src=../imgs/ilk_grade_5.bmp  align=""center"">"
elseif grade=4 and clangrade="ilk" Then
derece="<img src=../imgs/ilk_grade_4.bmp>"
elseif grade=3 and clangrade="ilk" Then
derece="<img src=../imgs/ilk_grade_3.bmp>"
elseif grade=2 and clangrade="ilk" Then
derece="<img src=../imgs/ilk_grade_2.bmp>"
elseif grade=1 and clangrade="ilk" Then
derece="<img src=../imgs/ilk_grade_1.bmp>"
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
.style12 {
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

<% if not clan.eof Then %><br><br>
<table width="550"  border="0" align="center" style="position:relative;top:30px;<%if len(capem)>0 Then Response.Write "left:48px"%>">

 <tr>
    <td align="center" valign="top" ><% if clan("nation")="2" Then
	 Response.Write "<img src=""../imgs/elmora.gif"" width=""100"" height=""80"">"
	 elseif clan("nation")="1" Then
	 Response.Write "<img src=""../imgs/krs.gif"" width=""100"" height=""80"">"
	 else
	 End If%>
	</td>
    <td width="340" align="center" valign="top">
   <table width="340" border="0" align="center">
  <tr>
    <td align="left" background="imgs/menubg.gif" class="style1"><strong>Clan Adý</strong></td>
    <td align="center" bgcolor="#F3D78B" class="stil1siralama"><b><%=clan("IDName")%></b></td>
    <td align="center" width="74" background="imgs/menubg.gif"  class="style1"><strong>Grade</strong></td>
      <tr>
    <td align="left" background="imgs/menubg.gif"><strong class="style1">Toplam Np</strong></td>
    <td align="center" bgcolor="#F3D78B"><%=ayir(points)%></td>
     <td rowspan="6" valign="top" align="center" width="64"><%=derece%></td></tr>
    </tr>

      <tr>
        <td width="92"  background="imgs/menubg.gif" bgcolor="#FF9900"><strong class="style1">Lider</strong></td>
        <td bgcolor="#F3D78B" align="center"><a href="../Karakter-Detay/<%=Trim(clan("Chief"))%>" onclick="javascript:pageload('Karakter-Detay/<%=Trim(clan("Chief"))%>');return false" class="link1"><%=clan("Chief")%></a></td>
      </tr>
      <tr>
        <td background="imgs/menubg.gif" ><strong class="style1">1.Asistan</strong></td>
        <td bgcolor="#F3D78B" align="center"><a href="../Karakter-Detay/<%=Trim(clan("ViceChief_1"))%>" onclick="javascript:pageload('Karakter-Detay/<%=Trim(clan("ViceChief_1"))%>');return false" class="link1"><%=clan("ViceChief_1")%></a></td>
      </tr>
      <tr>
        <td background="imgs/menubg.gif" ><strong class="style1">2.Asistan</strong></td>
        <td bgcolor="#F3D78B" align="center"><a href="../Karakter-Detay/<%=Trim(clan("ViceChief_2"))%>" onclick="javascript:pageload('Karakter-Detay/<%=Trim(clan("ViceChief_2"))%>');return false" class="link1"><%=clan("ViceChief_2")%></a></td>
      </tr>
      <tr>
        <td background="imgs/menubg.gif" ><strong class="style1">3.Asistan</strong></td>
        <td bgcolor="#F3D78B" align="center"><a href="../Karakter-Detay/<%=Trim(clan("ViceChief_3"))%>	" onclick="javascript:pageload('Karakter-Detay/<%=Trim(clan("ViceChief_3"))%>');return false" class="link1"><%=clan("ViceChief_3")%></a></td>
      </tr>
      <tr>
        <td background="imgs/menubg.gif" ><strong class="style1">Kuruluþ Tarihi</strong></td>
        <td bgcolor="#F3D78B" align="center"><%=left(clan("createtime"),10)%></td>
        </tr>
    </table></td>
        <td valign="top" >
	<table style="position:relative;left:-2px">
	<tr>
	<td background="imgs/menubg.gif" width="96" align="center"><strong class="style1">Pelerin</strong></td>
	</tr></table><%if not cape=-1 Then
	pelerin="../imgs/cape/cloak_c_"&cape&".gif"
	Response.Write "<img src="&pelerin&" width=""96"" height=""96"">"
	if len(capem)>0 Then
	Response.Write "<img src="&pelerinm&" width=""96"" height=""96"" style=""position:relative;left:-96"">"
	End If
	Else
	Response.Write "Yok"
	End If%>
</td>
 </tr>
</table>
<table width="550" border="0" align="center" style="position:relative;top:30px;">
  <tr>
    <td align="center" colspan="4"><strong>Clan Üyeleri (Toplam Üye: 
        <% =clan("members") %> ) </strong></td>
</tr>
  <tr>
    <td width="200" align="center" background="imgs/menubg.gif"><strong class="style1">Karakter Ad&#305; </strong></td>
    <td width="55" align="center"  background="imgs/menubg.gif"><strong class="style1">Level</strong></td>
    <td width="199" align="center" background="imgs/menubg.gif"><strong class="style1">National Point</strong></td>
    <td width="75" align="center"  background="imgs/menubg.gif"><strong class="style1">Durum</strong></td>
  </tr>

  <%Dim Onlinek
  If not usert.Eof Then 
  Do While Not usert.Eof
Set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&usert("strUserId")&"'") %>
   <tr bgcolor="#f3d78b" onmouseover="this.style.background='#D5AB4A'" onmouseout="this.style.background='#F3D78B'">
    <td align="center"><a href="../Karakter-Detay/<%=Trim(usert("strUserId"))%>" onclick="pageload('Karakter-Detay/<%=Trim(usert("strUserId"))%>');return false" style="display:block" class="link1"><%=usert("strUserId")%></a></td>
    <td align="center"><%=usert("Level")%></td>
    <td align="center"><%=ayir(usert("Loyalty"))%></td>
    <td align="center"><%if not onlinek.eof Then
		Response.Write "<font color='#FF0000'><strong>Oyunda !</strong></font>"
		elseif onlinek.eof Then
		Response.Write "<font color='#666666'>Çevrimdýþý</font>" 
		End If%></td>
   </tr>
   <%
  usert.Movenext
  Loop
	
 Else 
 Response.Write("<tr><td>Clanda Üye Yok</td></tr>")
  End If %>
</table>
<% else Response.Write"Böyle Bir Clan Bulunmamaktadýr."
 End If
 else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>