<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"--><html>
<base href="http://<%=Request.ServerVariables("server_name")%>">
<body bgcolor="#F9EED8" style="margin-left:0px;margin-top:0px" oncontextmenu="return false" onselectstart="return false">
<%response.expires=0
Response.Charset = "iso-8859-9"
If Not Request.ServerVariables("Script_Name")="/404.asp"  Then
yn("default.asp")
End If

Dim REFERER_URL,REFERER_DOMAIN
REFERER_URL = Request.ServerVariables("HTTP_REFERER")

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If

If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else

yn("default.asp")
End If
%>
<script type="text/javascript">
function kulbilgi(sid,detay){
$.ajax({
   url: 'userbilgi/'+sid+'/'+detay+'.html',
   success: function(ajaxCevap) {
      $('div#userbilgi').html(ajaxCevap);
   }
});
}
</script><script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<style>
<style type="text/css">

body,th {
	color: #FFFFFF;
	font-family:Verdana, Arial, Helvetica, sans-serIf;
	font-size:10px;
}
 
td {
	color: #000000;
	font-family:Verdana, Arial, Helvetica, sans-serIf;
	font-size:10px;
}
.style6 {
	color: #FFFFFF;
	font-weight: bold;
}
</style><table width="100%" oncontextmenu="return false"><tr><td>
<%Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='UserBilgi'or PId='HiddenOther'")
If MenuAyar("PSt")=1 Then
If Session("login")="ok" Then

Dim lnk,linkp,sid,detay
lnk=secur(Session("sayfa"))
linkp=split(lnk,"/")
sid=linkp(4)
If sid="Administrat" Then
sid="Administrator"
End If
If ubound(linkp)>4 Then
detay=linkp(5)
End If
If sid="" Then
Response.End
End If 

dim chard,usery
Set chard = Conne.Execute("Select * From USERDATA where strUserId='"&sid&"'")
If not chard.eof Then 
set usery =  Conne.Execute("select * from account_char where straccountid='"&Session("username")&"'")

If detay="" Then
detay="karakter"
End If

MenuAyar.Movenext
If menuayar("PSt")=0 or trim(sid)=trim(usery("strcharid1")) or trim(sid)=trim(usery("strcharid2")) or trim(sid)=trim(usery("strcharid3")) Then
If detay="friend" Then
 %>
<table width="301" height="376" border="0" align="left" cellpadding="0" cellspacing="0" style="background:url(/imgs/friendlist.gIf); background-repeat:no-repeat; ">
 <tr>
   <td>
   <div style="position:relative;top:-47px;"><a href="userbilgi/<%=sid%>/karakter" onClick="kulbilgi('<%=sid%>','karakter');return false" onMouseOver="MM_swapImage('char','','/imgs/tabcharon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabchar.gIf" name="char" id="char" border="0"></a></div>
   <div style="position:relative;top:-72px;left:110px"><a href="userbilgi/<%=sid%>/clan" onClick="kulbilgi('<%=sid%>','clan');return false" onMouseOver="MM_swapImage('clan','','/imgs/tabclanon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabclan.gIf" name="clan" id="clan" border="0"></a></div>
   <div style="position:relative;top:-97px;left:218px"><a href="userbilgi/<%=sid%>/friend" onClick="kulbilgi('<%=sid%>','friend');return false" onMouseOver="MM_swapImage('friend','','/imgs/tabfriendon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabfriend.gIf" name="friend" id="friend" border="0"></a></div>
     <div style="position:relative; left: 20px; top: -65px; width: 262px; height: 159px;overflow:auto; ">
     <table width="100%" border="0" >

	<%
	Set users = Conne.Execute("Select * From FRIEND_LIST Where strUserID = '"&chard("struserid")&"'")
	If not users.eof Then
	Dim x,renk
	for x=1 to 24
	set onlineler=Conne.Execute("select strcharid from currentuser where strcharid='"&users("strFriend"&x)&"'")
	If not onlineler.eof Then
	renk="#3F0"
	else
	renk="#CCC"
	End If
	%>
     <tr ><td width="22%" style="color:<%=renk%>;font-family: Arial, Helvetica, sans-serIf;font-size:12px">
     <%=users("strFriend"&x)%>
	</td>
    </tr>
     <%next	
	 End If %>
     </table></div></td></tr>
</table>
<%elseIf detay="clan" Then

Set clang =Conne.Execute("Select IDName,Chief,IDNum,Chief,ViceChief_1,ViceChief_2,ViceChief_3,Members,Nation,points,scape,ranking,flag,createtime From KNIGHTS Where IDNum = "&chard("Knights")&"")

If not clang.eof Then
Dim totaluser,cape,ranking,flag
set totaluser=Conne.Execute("select count(*) as tops from userdata where knights='"&clang("idnum")&"' ")

cape=clang("scape")
ranking=clang("ranking")
flag=clang("flag")

Dim pelerin
If cape="-1" Then
pelerin="Yok"
Else
pelerin="Var"
End If

Dim points,grade

If points<72000 Then
grade=5
elseIf points<144000 Then
grade=4
elseIf points<360000 Then
grade=3
elseIf points<720000 Then
grade=2
elseIf points>=720000 Then
grade=1
End If

dim clangrade,derece
If ranking=1 or ranking=2 or ranking=3 or ranking=4 or ranking=5 Then
clangrade="ilk"
else
clangrade="diger"
End If

If grade=5 and flag="2" Then
derece="<img src=""imgs/ust_grade_5.bmp"">"
elseIf grade=4 and flag="2" Then
derece="<img src=/imgs/ust_grade_4.bmp >"
elseIf grade=3 and flag="2" Then
derece="<img src=/imgs/ust_grade_3.bmp >"
elseIf grade=2 and flag="2" Then
derece="<img src=/imgs/ust_grade_2.bmp >"
elseIf grade=1 and flag="2" Then
derece="<img src=/imgs/ust_grade_1.bmp >"
End If

If grade=5 and flag="1" Then
derece="<img src=""imgs/alt_grade_5.bmp"">"
elseIf grade=4 and flag="1" Then
derece="<img src=""imgs/alt_grade_4.bmp"">"
elseIf grade=3 and flag="1" Then
derece="<img src=""imgs/alt_grade_3.bmp"">"
elseIf grade=2 and flag="1" Then
derece="<img src=""imgs/alt_grade_2.bmp"">"
elseIf grade=1 and flag="1" Then
derece="<img src=""imgs/alt_grade_1.bmp"">"
End If


If grade=5 and clangrade="ilk" Then
derece="<img src=""imgs/ilk_grade_5.bmp"">"
elseIf grade=4 and clangrade="ilk" Then
derece="<img src=""imgs/ilk_grade_4.bmp"">"
elseIf grade=3 and clangrade="ilk" Then
derece="<img src=""imgs/ilk_grade_3.bmp"">"
elseIf grade=2 and clangrade="ilk" Then
derece="<img src=""imgs/ilk_grade_2.bmp"">"
elseIf grade=1 and clangrade="ilk" Then
derece="<img src=""imgs/ilk_grade_1.bmp"">"
End If


%>
<table width="301" height="376" border="0" align="left" cellpadding="0" cellspacing="0" style="background:url(/imgs/clanimage.gif); background-repeat:no-repeat;">
 <tr>
   <td>
   <div style="position:relative;top:12px;"><a href="userbilgi/<%=sid%>/karakter" onClick="kulbilgi('<%=sid%>','karakter');return false" onMouseOver="MM_swapImage('char','','/imgs/tabcharon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabchar.gIf" name="char" id="char" border="0"></a></div>
   <div style="position:relative;top:-13px;left:110px"><a href="userbilgi/<%=sid%>/clan" onClick="kulbilgi('<%=sid%>','clan');return false" onMouseOver="MM_swapImage('clan','','/imgs/tabclanon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabclan.gIf" name="clan" id="clan" border="0"></a></div>
   <div style="position:relative;top:-38px;left:218px"><a href="userbilgi/<%=sid%>/friend" onClick="kulbilgi('<%=sid%>','friend');return false" onMouseOver="MM_swapImage('friend','','/imgs/tabfriendon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabfriend.gIf" name="friend" id="friend" border="0"></a></div>
     <div style="position:relative; left: 10px; top: 80px; width: 262px; height: 159px;overflow:auto; ">
     <table width="100%" border="0" >

	<%Dim users,onlineler
	Set users = Conne.Execute("Select strUserId,Knights,Loyalty,Level,fame,class From USERDATA Where Knights = "&clang("IDNum")&" order by fame asc")
	do while not users.eof
	set onlineler=Conne.Execute("select strcharid from currentuser where strcharid='"&users("struserid")&"'")
	If not onlineler.eof Then

	%>
     <tr ><td width="22%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"><%If users("fame")=0 Then
Response.Write ""
elseIf users("fame")=1 Then
Response.Write "Lider"
elseIf users("fame")=2 Then
Response.Write "Asistan"
elseIf users("fame")=5 Then
Response.Write "Üye"
End If%></td>
	<td width="42%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"><%=users("struserid")%></td>
   	<td width="10%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"><%=users("level")%></td>
    <td width="26%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"><%=cla2(users("class"))%></td>
    </tr>
     <%End If
	 users.movenext
	 loop%>
     </table></div>
     <div style="position:relative; left: 170px; top: -146px; width: 113px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
	 <%If chard("fame")=0 Then
Response.Write ""
elseIf chard("fame")=1 Then
Response.Write "Lider"
elseIf chard("fame")=2 Then
Response.Write "Asistan"
elseIf chard("fame")=5 Then
Response.Write "Üye"
End If%></div>
     <div style="position:relative; left: 120px; top: -195px; width: 137px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
       <center>
         <%=clang("idname")%>
       </center>
     </div>
     <div style="position:relative; left: 170px; top: -155px; width: 41px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
       <center>
         <%=totaluser("tops")%>
       </center>
     </div>
    <div style="position:relative; left: 18px; top: -225px; width: 75px; height: 68px;"><%=derece%></div>
    </td></tr>
</table>
<%

clang.close
set clang=nothing
else%>
<table width="301" height="376" border="0" align="left" cellpadding="0" cellspacing="0" style="background:url(/imgs/clanimage.gif); background-repeat:no-repeat; ">
 <tr>
   <td>
   <div style="position:relative;top:12px;"><a href="userbilgi/<%=sid%>/karakter" onClick="kulbilgi('<%=sid%>','karakter');return false" onMouseOver="MM_swapImage('char','','/imgs/tabcharon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabchar.gIf" name="char" id="char" border="0"></a></div>
   <div style="position:relative;top:-13px;left:110px"><a href="userbilgi/<%=sid%>/clan" onClick="kulbilgi('<%=sid%>','clan');return false" onMouseOver="MM_swapImage('clan','','/imgs/tabclanon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabclan.gIf" name="clan" id="clan" border="0"></a></div>
   <div style="position:relative;top:-38px;left:218px"><a href="userbilgi/<%=sid%>/friend" onClick="kulbilgi('<%=sid%>','friend');return false" onMouseOver="MM_swapImage('friend','','/imgs/tabfriendon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabfriend.gIf" name="friend" id="friend" border="0"></a></div>
     <div style="position:relative; left: 10px; top: 80px; width: 262px; height: 159px;overflow:auto; ">
     <table width="100%" border="0" >

	<tr ><td width="22%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"></td>
	<td width="42%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"></td>
   	<td width="10%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"></td>
    <td width="26%" style="color:#3F0;font-family: Arial, Helvetica, sans-serIf;font-size:10px"></td>
    </tr>
     
     </table></div>
     <div style="position:relative; left: 200px; top: -150px; width: 113px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
	</div>
     <div style="position:relative; left: 120px; top: -195px; width: 137px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
       <center>
         
       </center>
     </div>
     <div style="position:relative; left: 170px; top: -155px; width: 41px; height: 17px; font-size:12px; font-weight:bold;color:#E6DEBB;font-family: Arial, Helvetica, sans-serIf;">
       <center>
         
       </center>
     </div>
    <div style="position:relative; left: 18px; top: -225px; width: 75px; height: 68px;"></div>
    </td></tr>
</table>
<%End If
chard.close
set chard=nothing

elseIf detay="karakter" Then
Dim clang,lvlup
Set clang =Conne.Execute("Select IDNum,IDName From KNIGHTS Where IDNum = "&chard("Knights")&"")
Set lvlup = Conne.Execute("Select Level,Exp From LEVEL_UP Where Level = "&chard("Level")&"")
%>

<table width="301" height="382" border="0" align="left" cellpadding="0" cellspacing="0" style="background:url(/imgs/statbar.gIf); background-repeat:no-repeat; ">
  <tr>
    <td colspan="3"><div style="position: relative; top: 24px; left: 0px;"><a href="userbilgi/<%=sid%>/karakter" onClick="kulbilgi('<%=sid%>','karakter');return false" onMouseOver="MM_swapImage('char','','/imgs/tabcharon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabchar.gIf" name="char" id="char" border="0"></a></div>
<div style="position: relative; top: -1px; left: 110px;"><a href="userbilgi/<%=sid%>/clan" onClick="kulbilgi('<%=sid%>','clan');return false" onMouseOver="MM_swapImage('clan','','/imgs/tabclanon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabclan.gIf" name="clan" id="clan" border="0"></a></div>
<div style="position: relative; top: -26px; left: 218px;"><a href="userbilgi/<%=sid%>/friend" onClick="kulbilgi('<%=sid%>','friend');return false" onMouseOver="MM_swapImage('friend','','/imgs/tabfriendon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabfriend.gIf" name="friend" id="friend" border="0"></a></div></td>
  </tr>
  <tr>
    <td height="25" colspan="3" class="style6" style="padding-left:15px;padding-top:-15px"><%=chard("strUserId")%></td>
    <td colspan="2" rowspan="2"></td>
    <td width="56" rowspan="2"></td>
  </tr>
  <tr>
    <td height="20" colspan="3" valign="top" class="style6" style="padding-left:15px"><font color="#DFC68C">
      <%=cla2(chard("Class"))%>
    </font></td>
  </tr>
  <tr>
    <td height="18" colspan="2"></td>
    <td width="48" align="center" class="style6" style="padding-left:-55px"><%=chard("Level")%></td>
    <td align="center" class="style6"></td>
    <td colspan="2" align="center" valign="middle" class="style6"><% If chard("Nation")="1" Then
		  Response.Write "Karus"
          elseIf chard("Nation")="2" Then
		  Response.Write "El-Morad"
          End If %></td>
  </tr>
  <tr>
    <td height="32" colspan="2"></td>
    <td height="32" colspan="3" align="center" class="style6"><%=chard("Exp")%> / <%=lvlup("Exp")%></td>
    <td height="32" align="center"></td>
  </tr>
  <tr>
    <td height="23" colspan="2"></td>
    <td height="23" colspan="3" align="center" valign="middle" class="style6"><%=chard("Loyalty")%> / <%=chard("LoyaltyMonthly")%></td>
    <td height="23" align="center"></td>
  </tr>
  <tr>
    <td width="44" valign="bottom">&nbsp;</td>
    <td height="46" colspan="2" align="center" valign="bottom" class="style6"><%=chard("Strong")%></td>
    <td width="55" valign="bottom" class="style6">&nbsp;</td>
    <td width="55" align="center" valign="bottom" class="style6"><%=chard("Cha")%></td>
    <td align="center" valign="bottom">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="48" colspan="2" align="center" valign="middle" class="style6"><%=chard("Sta")%></td>
    <td width="55">&nbsp;</td>
    <td align="center" valign="middle" class="style6"><%=chard("Intel")%></td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="25" colspan="2" align="center" valign="middle" class="style6"><%=chard("Dex")%></td>
    <td width="55">&nbsp;</td>
    <td height="25" align="center" valign="middle" class="style6"><%=chard("Points")%></td>
    <td height="25" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td height="70"></td>
  </tr>
</table>
<%
lvlup.close
set lvlup=nothing
clang.close
set clang=nothing
chard.close
set chard=nothing

End If
End If
else
Response.Write("Karakter Bulunamadi")
End If 

else
Response.Write "<br><b>Lütfen Giris yapiniz.</b>"
End If

MenuAyar.Close
Set MenuAyar=Nothing
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serIf;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
%></td></tr>
</table></body></html>