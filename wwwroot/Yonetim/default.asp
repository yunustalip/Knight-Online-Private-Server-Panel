<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"-->
<!--#include file="../md5.asp"-->
<%Response.CharSet = "windows-1254" 
Session.lcid = 1055
Session.CodePage = 1254

If Session("strAccountID")="" Or Not Session("durum")="esp" Then
Response.Redirect("login.asp")
Response.End
End If
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Control Panel</title>
<link href="css/styles.css" rel="stylesheet" type="text/css" >
<script type="text/javascript" src="_inc/user.js"></script>
<script type="text/javascript" src="overlib.js"></script>
<script src="_inc/jquery-1.4.2.js"></script>
<script type="text/javascript">
function rsm(val)
{
  $.ajax({
	type: 'GET',
	url: 'resim.asp',
	data: 'resim='+val,
	success: function(Cevap) {
		$('#resim').val(Cevap);
	}
  });
}
function sayfad(syf)
{
  $.ajax({
	type: 'GET',
	url: 'default.asp'+syf,
	success: function(sCevap) {
		$('div#sayfa').html(sCevap);
	}
  });
}
</script>
 <script language="JavaScript" type="text/JavaScript">
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
<script language="JavaScript" type="text/javascript" src="wysiwyg.js"></script>
<script language="JavaScript" type="text/javascript">
function HepsiniSec(divId)
{
var elmnts = document.getElementById(divId)
.getElementsByTagName('input');
for (var i = 0; i < elmnts.length; i++)
{
If (elmnts[i].id.indexOf('text') == 0)
{
If (elmnts[i].value ==! 1)
{
elmnts[i].value = 1;
}
}
}
}

function Temizle(divId)
{
var elmnts = document.getElementById(divId)
.getElementsByTagName('input');
for (var i = 0; i < elmnts.length; i++)
{
If (elmnts[i].id.indexOf('text') == 0)
{
If (elmnts[i].value !== 1)
{
elmnts[i].value = '';
}
}
}
}
</script>
<style type="text/css">
<!--
.style8 {font-size: 16px}
-->
</style>
</head>

<body background="../imgs/kol-bkgrd.jpg">
<style>
body{
color:#F9EFD7;
}
.menu td a{
background:url("../imgs/ybg.gif");
height:32px;
color:#89640B;
azimuth:center
}

.menu td a:hover{
color:#76550A;
text-decoration:none
}
.menu td a:hover{
background:url("../imgs/ybgon.gif");
color:#F9EFD7;
height:32px;
}
</style>
<table width="644" height="281" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td colspan="2" align="center">
	<img src="xmgs/105450.gif" alt="image" width="545" height="69" border="0">
	</td>
</tr>
  <tr>
    <td width="181" valign="top" bgcolor="#ffffff">
    <table width="181" border="0" border-color="#CCCCCC" >
    <tr>
    <td valign="top">
    
	<table width="177" border="0" class="menu">
	<tr>
        <td align="left">
	<a href="default.asp?w8=siteayarlari" style="display:block;"><img src="../imgs/home.gif" width="32" height="32" border="0" align="absmiddle" />
	Ana Sayfa Ayarları</a></td>
      </tr>
 	<tr>
        <td align="left" ><a href="default.asp?w8=0" style="display:block"><img src="../imgs/Control Panel.gif" width="32" height="32" border="0" align="absmiddle" />
	Site Sayfa Ayarları</a></td>
	</tr>
 	<tr>
        <td align="left" ><a href="default.asp?w8=menusettings" style="display:block"><img src="../imgs/database_process.gif" width="32" height="32" border="0" align="absmiddle" />
	Menü Ayarları</a></td>
	</tr>
	<tr>
        <td align="left" ><a href="#" onclick="window.open('EnterpriseManager.asp','','fullscreen=0,top=0,left=0,resizable=1,status=1,scrollbars=1,menubar=0,toolbar=0,height='+screen.availHeight+', width='+screen.availWidth+10);return false" style="display:block"><img src="../imgs/sql-query.gif" width="32" height="32" border="0" align="absmiddle" />
	SQL EnterPrise Manager</a></td>
      </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=kod" style="display:block"><img src="../imgs/sql-query.gif" width="32" height="32" border="0" align="absmiddle" />
	Query Analyzer</a></td>
      </tr>
      <tr>
        <td  align="left" ><a href="default.asp?w8=manager" style="display:block"><img src="../imgs/Netdrive Connected.gif" width="32" height="32" border="0" align="absmiddle" />
	Server Yöneticisi</a></td>
      </tr>
      <tr>
        <td align="left"><a href="default.asp?w8=res" style="display:block"><img src="../imgs/database_accept.gif" width="32" height="32" border="0" align="absmiddle" />
	Server Reset </a></td>
      </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=ticket" style="display:block"><img src="../imgs/Mail_new.gif" width="32" height="27" border="0" align="absmiddle" />
	Gelen Ticketler</a></td>
    </tr>
      <tr>
        <td  align="left" ><a href="default.asp?w8=anket" style="display:block"><img src="../imgs/1253482283_mime_txt.gif" width="32" height="32" border="0" align="absmiddle" />
	Anket Yönetimi</a></td>
      </tr>
      <tr>
        <td  align="left" ><a href="default.asp?w8=news" style="display:block"><img src="../imgs/news.gif" width="32" height="32" border="0" align="absmiddle" />
	Haber & Event Yönetimi</a></td>
      </tr>
      <tr>
        <td align="left" ><a href="default.asp?w8=filecontrol" style="display:block"><img src="../imgs/dos.gif"  border="0" align="absmiddle"/>
	Dosya Yönetimi</a></td>
      </tr>
      <tr>
        <td  align="left" ><a href="default.asp?w8=pus" style="display:block"><img src="../imgs/85.gif" width="32" height="30"  border="0" align="absmiddle"/>
	Power Up Store </a></td>
      </tr>
     <tr>
        <td  align="left" ><a href="default.asp?w8=premium" style="display:block"><img src="../imgs/cashu.gif" width="32" height="32" border="0" align="absmiddle" />
	Premium &amp; Cash Editor </a></td>
      </tr>
      <tr>
        <td align="left" ><a href="default.asp?w8=ozelitem" style="display:block"><img src="../imgs/../imgs/86.gif" width="31" height="32"  border="0" align="absmiddle">
	Özel Item Editorü </a></td>
      </tr> 
      <tr>
        <td align="left" ><a href="default.asp?w8=1" style="display:block"><img src="../imgs/user_info_32.gif" width="31" height="32"  border="0" align="absmiddle">
	User Editor </a></td>
      </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=10" style="display:block"><img src="../imgs/BriefCase.gif" width="32" height="26"  border="0" align="absmiddle"/>
	Inventory Editor</a></td>
    </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=20" style="display:block"><img src="../imgs/Bank.gif" border="0" align="absmiddle"/>
	Banka Editor</a></td>
    </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=baslangic" style="display:block"><img src="../imgs/mappe_128.gif" border="0" align="absmiddle"/>
	Başlangıç Item Editor</a></td>
    </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=bankabaslangic" style="display:block"><span align="left"><img src="../imgs/Box.gif" width="32" height="26"  border="0" align="absmiddle" /></span>
	<span align="center" style="height:32px;text-align:center;">Banka Başlangıç Item Editor</span></a></td>
    </tr>
     <tr>
        <td align="left" ><a href="default.asp?w8=itemmall" style="display:block"><img src="../imgs/32x32_mail.gif" width="32" height="32" border="0" align="absmiddle">
	Online Item  Gönder </a></td>
      </tr>
     <tr>
        <td align="left" ><a href="default.asp?w8=12" style="display:block"><img src="../imgs/1253487231_edit.GIF" width="32" height="32" border="0" align="absmiddle">
	Item Editor </a></td>
      </tr>
     <tr>
        <td align="left" ><a href="default.asp?w8=itemfinder" style="display:block"><img src="../imgs/iSearch.gif" width="32" height="32" border="0" align="absmiddle">
	Item Finder </a></td>
      </tr>
          <tr>
        <td align="left" ><a href="default.asp?w8=7" style="display:block"><img src="../imgs/contacts.gif"  border="0" align="absmiddle"/>
	Account Editor</a></td>
      </tr>
    <tr>
        <td align="left" ><a href="default.asp?w8=2" style="display:block"><img src="../imgs/405.gif" width="32" height="32" border="0" align="absmiddle" />
	Char Irk/Class Değişimi </a></td>
    </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=9" style="display:block"><img src="../imgs/452.gif" width="32" height="32" border="0" align="absmiddle" />
	Account Irk Değişimi</a></td></tr>
		<tr>
        <td align="left" ><a href="default.asp?w8=clanedit" style="display:block"><img src="../imgs/Workgroup Network.gif" width="32" height="32" border="0" align="absmiddle" />
	Clan Editor</a></td></tr>
     <tr>
	<td align="left" ><a href="default.asp?w8=version" style="display:block"><img src="../imgs/kx2-control-panel.gif" width="31" height="32"  border="0" align="absmiddle" />
	Version Editor </a></td>
      </tr>
     <tr>
	<td align="left" ><a href="default.asp?w8=town" style="display:block"><img src="../imgs/32x32_home.gif" width="32" height="32" border="0" align="absmiddle" />
	Town Koordinat Editor</a></td>
      </tr>
      <tr>
        <td align="left" ><a href="default.asp?w8=npc" style="display:block"><img src="../imgs/Freecell.gif" width="32" height="32" border="0" align="absmiddle" />
	Npc Editor </a></td>
      </tr>
      <tr>
        <td align="left" ><a href="default.asp?w8=monster" style="display:block"><img src="../imgs/212.gif" width="32" height="32" border="0" align="absmiddle" />
	Monster Editor </a></td>
      </tr>
        <tr>
	<td align="left" ><a href="default.asp?w8=level" style="display:block"><img src="../imgs/Chart_1.gif" width="31" height="32"  border="0" align="absmiddle" />
	Level Up Editor </a></td>
      </tr>
      <tr>
        <td align="left" ><a href="default.asp?w8=upgrade" style="display:block"><img src="../imgs/chart.gif" width="32" height="32" border="0" align="absmiddle" />
	Upgrade Editor </a></td>
      </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=idbul" style="display:block"><img src="../imgs/402.gif" width="32" height="32" border="0" align="absmiddle" />
	Login ID Bulucu</a></td>
      </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=produces" style="display:block"><img src="../imgs/database_search.gif" width="32" height="32" border="0" align="absmiddle" />
	Procedureler</a></td>
       </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=pmbox" style="display:block"><img src="../imgs/kwrite.gif" width="32" height="32" border="0" align="absmiddle" />
	Özel Mesajlar</a></td>
       </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=backup" style="display:block" ><img src="../imgs/database_previous.png" width="32" height="32" border="0" align="absmiddle" >
    Database Yedekleme</a></td>
       </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=logs" style="display:block" ><img src="../imgs/applications-utilities.gif" width="32" height="32" border="0" align="absmiddle" > Site Logları</a></td>
       </tr>
	<tr>
        <td align="left" ><a href="default.asp?w8=onlineuser" style="display:block" ><img src="../imgs/info.gif" width="32" height="32" border="0" align="absmiddle" > Online Kullanıcı Detay</a></td>
       </tr>
	<tr>
        <td align="center" ><a href="default.asp?w8=errorlogs" style="display:block" ><img src="../imgs/Windows Security.gif" width="32" height="32" border="0" align="left" > Sitede Meydana Gelen Teknik Hatalar</a></td>
       </tr>
       <tr>
        <td align="left" ><a href="logout.asp" style="display:block" onClick="return confirm('Çıkış Yapmak Istediğinizden Eminmisiniz?')"><img src="../imgs/remove.gif" align="absmiddle" width="32" height="32" border="0" />
	Çıkış</a></td>
       </tr>
    </table>
    
	</td>
    </tr>
    </table>
</td>
    <td width="800" valign="top" bgcolor="#FFFFFF">
	<table width="790" height="362" border="0">
<tr>
        <td width="15">&nbsp;</td>
        <td width="745" valign="top">
        <span id="txtHint">
<%Dim w8
w8=Request.Querystring("w8")
If w8="" Then
w8="siteayarlari"
End If
select Case w8
Case "clanedit"
If Session("durum")="esp" Then
If Trim(Request.QueryString("islem"))="edit" Then
id=Trim(Request.QueryString("id"))
dim clanbul
Set ClanBul=Conne.Execute("Select * From Knights Where IDNum="&id)
Response.Write("<table><tr><td>IDnum</td><td><input name=""idnum"" value="""&ClanBul("idnum")&""">"&vbcrlf)
Response.Write("<tr><td>IDName</td><td><input name=""idname"" value="""&Trim(ClanBul("idname"))&""">"&vbcrlf)
Response.Write("<tr><td>Clan Gradesi</td><td><select name=""flag""><option value=""1""")
If Clanbul("flag")="1" Then Response.Write(" Selected ")
Response.Write(">Alt Clan(Pelerinsiz)</option><option value=""2""")
If Clanbul("flag")="2" Then Response.Write(" Selected ")
Response.Write(">Üst Clan(Pelerinli)</option>")
Response.Write("<tr><td>Irk</td><td><select name=""nation""><option value=""1""")
If Clanbul("nation")="1" Then Response.Write(" Selected ")
Response.Write(">Karus</option><option value=""2""")
If Clanbul("nation")="2" Then Response.Write(" Selected ")
Response.Write(">Human</option></td></tr>")
Response.Write("<tr><td>Clan Üye Sayısı</td><td><input name=""members"" value="""&ClanBul("members")&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>Clan Lideri</td><td><input name=""chief"" value="""&Trim(ClanBul("chief"))&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>1. Asistan</td><td><input name=""vicechief_1"" value="""&Trim(ClanBul("vicechief_1"))&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>2. Asistan</td><td><input name=""vicechief_2"" value="""&Trim(ClanBul("vicechief_2"))&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>3. Asistan</td><td><input name=""vicechief_3"" value="""&Trim(ClanBul("vicechief_3"))&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>Toplam Np (Points)</td><td><input name=""points"" value="""&Trim(ClanBul("points"))&"""></td></tr>"&vbcrlf)
Response.Write("<tr><td>Ranking(Sıra)</td><td><input name=""points"" value="""&Trim(ClanBul("ranking"))&"""></td></tr>"&vbcrlf)
cape=ClanBul("sCape")
	If Not cape=-1 Then
	If Len(cape)=3 Then
	capem=left(cape,1)
	pelerinm="../imgs/cape/cloak_m_0"&capem&".gif"
	cape=mid(cape,2,2)
	End If
	if len(cape)=1 Then
	cape=0&cape
	End If

	pelerin="../imgs/cape/cloak_c_"&cape&".gif"
	Response.Write "<tr><td valign=""top"">PELERIN: </td><td><img src="""&pelerin&""" align=""absmiddle"">"
	if len(capem)>0 Then
	Response.Write "<img src="&pelerinm&" style=""position:relative;left:-128"" align=""absmiddle""><br>"
	End If
	End If
	capem=""
	pelerin=""
	pelerinm=""
	Response.Write("<a HREF=""#"" onClick=""acYardim01()"">Pelerin Seç</a>"&vbcrlf)
	Response.Write("<script>function acYardim01(){yeniPencere = window.open('', 'yardim01', 'height=500,width=600,resizable=1,status=0,scrollbars=1,menubar=0,toolbar=0') "&vbcrlf)
	Response.Write("yeniPencere.document.write('")
	dim pelerinler
	Set Pelerinler=Conne.Execute("Select sCapeIndex,strName,byGrade From KNIGHTS_CAPE")
	dim sayi
	sayi=1
	Response.write "<table>"
	Do While Not Pelerinler.Eof
	cape=Pelerinler(0)
	If Not cape=-1 Then
	If Len(cape)=3 Then
	capem=Left(cape,1)
	pelerinm="../imgs/cape/cloak_m_0"&capem&".gif"
	cape=Mid(cape,2,2)
	End If
	If Len(cape)=1 Then
	cape="0"&cape
	End If

	pelerin="../imgs/cape/cloak_c_"&cape&".gif"
	Response.Write("<td id="""&sayi&""" onclick=""selectedbox('"&sayi&"')"">")
	Response.Write "<img src="""&pelerin&""" align=""absmiddle"" >"
	if len(pelerinm)>0 Then
	Response.Write "<img src="""&pelerinm&""" style=""position:relative;left:-128""  align=""absmiddle"">"
	Else
	Response.Write("<img src=""../imgs/blank.gif"" width=""128"">")
	End If
	
	End If
	cape=""
	pelerinm=""
	pelerin=""
	If sayi mod 2 = 0 Then
	Response.Write "</td></tr><tr>"
	Else
	Response.write "</td>"
	End If

	Pelerinler.MoveNext
	sayi=sayi+1
	
	Loop
	Response.Write("');yeniPencere.document.close()}</script>")
Response.Write("</td></tr></table>")%>
<script>function selectedbox(slot){
 for (var i=1; i<<%=sayi%>; i++){
  var spn = document.getElementById(i);
  spn.style.border = "1px solid #4C4B36";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
	}
</script>
<%Response.End
End If

Dim Clans
Set Clans=Conne.Execute("Select * From Knights Order By IdName")

Response.Write("<div style=""height:500px;width:400px;font-weight:bold;overflow:scroll""><table width=""300""><tr><td style=""color:#FF0000;font-weight:bold""><u>IDNUM</u></td><td style=""color:#FF0000;font-weight:bold""><u>Clan ID Name</u></td>")
Do While Not Clans.Eof
Response.Write "<tr ><td style=""font-weight:bold"">"&Clans("idnum")&"</td><td style=""font-weight:bold""><a href=""default.asp?w8=clanedit&islem=edit&id="&clans("idnum")&""">"&Clans("IdName")&"</a></td></tr>"
Clans.MoveNext
Loop
Response.Write("</table></div>")

Else
Response.Redirect("default.asp")
End If

Case "itemfinder"
If Session("durum")="esp" Then
%> <script language="javascript">
    function itemfind(){
    $.ajax({
       type: 'get',
       url: 'itemfinder.asp',
data: $('#ifind').serialize(),
       start: $('div#itemsonuc').html('<center><img src="../imgs/38-1.gif"><br><b>Aranıyor...</b></center>'),
error: function(ajaxCevap) {
          $('div#itemsonuc').html(ajaxCevap);
       },
       success: function(ajaxCevap) {
          $('div#itemsonuc').html(ajaxCevap);
       }
    });
    }
    </script>
<form action="javascript:itemfind()" name="ifind" id="ifind">
<table>
<tr>
<td>Item Numarası</td><td>Aranacak Yer</td><td></td>
</tr>
<tr>
<td><input type="text" name="itemno"></td>
<td><select name="fmode">
<option value="1">Inventory</option>
<option value="2">Banka</option>
</select></td><td><input type="submit" value="Ara"></td>
</tr>
</table></form><div id="itemsonuc"></div>
<%If Request.Querystring("islem")="sil" Then
	dwid=Trim(Request.Querystring("dwid"))
	pos=Trim(Request.Querystring("pos"))
	strUserID=Trim(Request.Querystring("struserid"))
fmode=cint(trim(Request.Querystring("fmode")))
If fmode=1 Then
If pos<>"" And IsNumeric(pos) Then pos=CInt(pos)
	Set rs=Conne.Execute("select strCharID from CURRENTUSER where strCharID='" & strUserID & "'")
	If Not rs.eof Then
		Response.Write "Karakter Online iken bu işlemi yapamazsınız."
		Response.End
	Else
	If dwid<>0 And strUserID<>"" And pos<>"" Then
		dodel=false
		dim EqItem(3,41)
		k = 0
		j = 0
	Do While j < 42
sqlc = "select cast(cast(substring(cast(strserial as varbinary(400))," & k + 4 & ", 1)+"
sqlc = sqlc+"substring(cast(strserial as varbinary(400)), " & k + 3 & ", 1)+"
sqlc = sqlc+"substring(cast(strserial as varbinary(400)), " & k + 2 & ", 1)+"
sqlc = sqlc+"substring(cast(strserial as varbinary(400))," & k + 1 & ", 1) as varbinary(4)) as int(4)) as strserial,"
sqlc = sqlc+"cast(cast(substring(cast(strItem as varbinary(400))," & k + 4 & ", 1)+"
sqlc = sqlc+"substring(cast(strItem as varbinary(400)), " & k + 3 & ", 1)+"
sqlc = sqlc+"substring(cast(strItem as varbinary(400)), " & k + 2 & ", 1)+"
sqlc = sqlc+"substring(cast(strItem as varbinary(400))," & k + 1 & ", 1) as varbinary(4)) as int(4)) as id,"
sqlc = sqlc+"cast(cast(substring(cast(strItem as varbinary(400))," & k + 6 & ", 1)+"
sqlc = sqlc+"substring(cast(strItem as varbinary(400)), " & k + 5 & ", 1) as varbinary(2)) as smallint(2)) as dur,"
sqlc = sqlc+"cast(cast(substring(cast(strItem as varbinary(400))," & k + 8 & ", 1)+"
sqlc = sqlc+"substring(cast(strItem as varbinary(400)), " & k + 7 & ", 1) as varbinary(2)) as smallint(2)) as Count from USERDATA where strUserId='" & strUserID & "'"
	Set rs=Conne.Execute(sqlc)

	If Rs("Id") >= 0 Then
	EqItem(0, j) = Rs("Id")
	EqItem(1, j) = Rs("Dur")
	EqItem(2, j) = Rs("Count")
	EqItem(3, j) = Rs("strserial")

	If  pos=j Then
	EqItem(0, j)=0
	EqItem(1, j)=0
	EqItem(2, j)=0
	EqItem(3, j)=0
	dodel=true
	End If
	End If
	j = j + 1
	k = k + 8
	Rs.Close
	Loop

	If dodel=true Then
	k = 0
	j = 0
Do While j < 42
Sql = "Update USERDATA Set strItem=cast(substring(cast(strItem as varbinary(400)),1," & k & ")+"
Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),4,1)+"
Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),3,1)+"
Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),2,1)+"
Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),1,1)+"
Sql = Sql + "substring(cast(0 as varbinary(2)),2,1)+"
Sql = Sql + "substring(cast(0 as varbinary(2)),1,1)+"
Sql = Sql + "substring(cast(0 as varbinary(2)),2,1)+"
Sql = Sql + "substring(cast(0 as varbinary(2)),1,1) as varbinary(400)) where strUserId='" & strUserID & "'"
Conne.Execute Sql
	j = j + 1
	k = k + 8
	Loop

	k = 0
	j = 0
Do While j < 42
Sql = "Update USERDATA Set strserial=cast(substring(cast(strserial as varbinary(400)),1," & k & ")+"
Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),4,1)+"
Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),3,1)+"
Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),2,1)+"
Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),1,1)+"
Sql = Sql + "substring(cast(" & EqItem(1, j) & " as varbinary(2)),2,1)+"
Sql = Sql + "substring(cast(" & EqItem(1, j) & " as varbinary(2)),1,1)+"
Sql = Sql + "substring(cast(" & EqItem(2, j) & " as varbinary(2)),2,1)+"
Sql = Sql + "substring(cast(" & EqItem(2, j) & " as varbinary(2)),1,1) as varbinary(400)) where strUserId='" & strUserID & "'"
Conne.Execute Sql
	j = j + 1
	k = k + 8
	Loop

	End If
	End If
	End If
End If
End If 
If fmode=2 Then
strAccountId=strUserID
	If pos<>"" And IsNumeric(pos) Then pos=CInt(pos)

	Set rs=Conne.Execute("select strCharID from dbo.CURRENTUSER where strCharID='" & strUserID & "'")
	If Not rs.eof Then
Response.Write "karakter online iken bu işlemi yapamazsınız"
	Else
		If dwid<>0 And strAccountId<>"" And pos<>"" Then
			dodel=false
			redim EqItem(3,192)
			kk = 0
			j = 0
			Do While j < 192
				Sql = "select "
				Sql = Sql & " cast(cast(substring(cast(WarehouseData as varbinary(1600))," & kk + 4 & ", 1)+substring(cast(WarehouseData as varbinary(1600)), " & kk + 3 & ", 1)+substring(cast(WarehouseData as varbinary(1600)), " & kk + 2 & ", 1)+substring(cast(WarehouseData as varbinary(1600))," & kk + 1 & ", 1) as varbinary(4)) as int(4)) as dwid,"
				Sql = Sql & " cast(cast(substring(cast(strSerial as varbinary(1600))," & kk + 8 & ", 1)+substring(cast(strSerial as varbinary(1600))," & kk + 7 & ", 1)+substring(cast(strSerial as varbinary(1600))," & kk + 6 & ", 1)+substring(cast(strSerial as varbinary(1600))," & kk + 5 & ", 1)+substring(cast(strSerial as varbinary(1600)), "& kk + 4 & ", 1)+substring(cast(strSerial as varbinary(1600))," & kk + 3 & ", 1)+substring(cast(strSerial as varbinary(1600)), "& kk + 2 & ", 1)+substring(cast(strSerial as varbinary(1600))," & kk + 1 & ", 1) as varbinary(8)) as bigint) as strSerial,"
				Sql = Sql & " cast(cast(substring(cast(WarehouseData as varbinary(1600))," & kk + 6 & ", 1)+substring(cast(WarehouseData as varbinary(1600)), " & kk + 5 & ", 1) as varbinary(2)) as smallint(2)) as dur,"
				Sql = Sql & " cast(cast(substring(cast(WarehouseData as varbinary(1600))," & kk + 8 & ", 1)+substring(cast(WarehouseData as varbinary(1600)), " & kk + 7 & ", 1) as varbinary(2)) as smallint(2)) as Count"
				Sql = Sql & " from WAREHOUSE where strAccountID='" & strAccountId & "'"
				Set rs= Server.CreateObject("ADODB.Recordset")
				Rs.Open Sql, Conne, 1, 1
				If Rs("dwid") >= 0 Then
					EqItem(0, j) = Rs("dwid")
					EqItem(1, j) = Rs("Dur")
					EqItem(2, j) = Rs("Count")
					EqItem(3, j) = Rs("strserial")
					If  pos=j Then
						EqItem(0, j)=0
						EqItem(1, j)=0
						EqItem(2, j)=0
						EqItem(3, j)=0
						dodel=true
					End If
				End If
				j = j + 1
				kk = kk + 8
				Rs.Close
			Loop

			If dodel Then
				kk = 0
				j = 0
				Do While j < 192
					Sql = "Update warehouse Set WarehouseData=cast(substring(cast(WarehouseData as varbinary(1600)),1," & kk & ")+"
					Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),4,1)+"
					Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),3,1)+"
					Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),2,1)+"
					Sql = Sql + "substring(cast(" & EqItem(0, j) & " as varbinary(4)),1,1)+"
					Sql = Sql + "substring(cast(" & EqItem(1, j) & " as varbinary(2)),2,1)+"
					Sql = Sql + "substring(cast(" & EqItem(1, j) & " as varbinary(2)),1,1)+"
					Sql = Sql + "substring(cast(" & EqItem(2, j) & " as varbinary(2)),2,1)+"
					Sql = Sql + "substring(cast(" & EqItem(2, j) & " as varbinary(2)),1,1) as varbinary(1600)) where strAccountID='" & strAccountID & "'"
					Conne.Execute Sql
				
					Sql = "Update warehouseSetstrserial=cast(substring(cast(strserial as varbinary(1600)),1," & kk & ")+"
					Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),4,1)+"
					Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),3,1)+"
					Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),2,1)+"
					Sql = Sql + "substring(cast(" & EqItem(3, j) & " as varbinary(4)),1,1)+"
					Sql = Sql + "substring(cast(0 as varbinary(2)),2,1)+"
					Sql = Sql + "substring(cast(0 as varbinary(2)),1,1)+"
					Sql = Sql + "substring(cast(0 as varbinary(2)),2,1)+"
					Sql = Sql + "substring(cast(0 as varbinary(2)),1,1) as varbinary(1600)) where strAccountID='" & strAccountID & "'"
					Conne.Execute Sql
					j = j + 1
					kk = kk + 8
				Loop
				
			Else
			End If
		End If
	End If

Response.Write "<script> alert('"&struserid&", Nickli Karakterden "&dwid&" Nolu Item Silinmiştir.')</script>"
End If
Else
Response.Redirect("default.asp")
End If
Case "errorlogs"
If Session("durum")="esp" Then
If Request.Querystring("islem")="sil" Then
dim id
id=Request.Querystring("id")
Conne.Execute("delete from errorlogs where id="&id)
Response.Redirect("default.asp?w8=errorlogs")
ElseIf Request.Querystring("islem")="truncate" Then
Conne.Execute("truncate table errorlogs")
End If
dim errorlog,hatano,ay
Set errorlog=Conne.Execute("select * from ErrorLogs order by errortime")
dim a(100),b(100),c(100)
hatano=1
If errorlog.eof Then
Response.Write "<br><b>Herhangi Bir Teknik Hata Uyarısı Bulunmadı.</b>"
Response.End
End If
Response.Write "<a href=""default.asp?w8=errorlogs&islem=truncate"">Kayıtları Temizle</a>"
do while not errorlog.eof
ay=split(errorlog(4),",")
Response.Write "<table border=""1""><tr><td>"&errorlog(1)&", "&errorlog(0)&"<br>"&errorlog(2)&"<br><b>"&ay(0)&", "&ay(1)
If ubound(ay)=4 Then
Response.Write ", Karakter: "&ay(4)
End If
Response.Write "</b><br>"


Response.Write "<br>Tarayıcı Tipi:<br> "&ay(2)&"<br><br>Tarih:<br>"&ay(3)&"</td><td><a href=""default.asp?w8=errorlogs&islem=sil&id="&errorlog("id")&""">SIL</a></td></tr></table><br>"
errorlog.movenext
hatano=hatano+1
Loop
Else
Response.Redirect("default.asp")
End If
Case "onlineuser"
If Session("durum")="esp" Then
%><!--#include file="../webonline.asp"--><style>
td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
}

</style><table border="1" cellpadding="5" cellspacing="1">
<tr>
<td>Session ID</td>
<td>IP</td>
<td>Bulunduğu Sayfa</td>
<td>Son Güncellenme Tarihi</td>
<td>Tarayıcı</td>
</tr>
<%AktifKullanicilariSil
dim aktifkullanicilar,akul,xy,userinf
aktifkullanicilar=application("aktifkullanicilistesi")

akul=split(aktifkullanicilar,"|")
for xy=0 to ubound(akul)-1
userinf=split(akul(xy),"#")
Response.Write "<tr><td>"&userinf(0)&"</td><td>"&userinf(2)&"</td><td>"&userinf(3)&"</td><td>"&userinf(1)&"</td><td>"&userinf(4)&"</td></tr>"
next%>
</table>
<%Else
Response.Redirect("default.asp")
End If
Case "backup"
If Session("durum")="esp" Then
islem=Request.Querystring("islem")
Dim connx
Set connx = Server.CreateObject("ADODB.Connection")
connx.open= "driver={SQL Server};server="&sunucu&";database=koyedek;uid="&kullanici&";pwd="&sifre&"" %>
<br><a href="default.asp?w8=backup&islem=backup" >Tüm User Bilgilerinin Yedeğini Al</a><br><br><a href="default.asp?w8=backup&islem=backupup">Tüm User Bilgilerinin Yedeklerini Geri Yükle</a>
<br><br><a href="default.asp?w8=backup&islem=userbackup">Bir Userin Bilgilerini Yedekle</a>
<br><br><a href="default.asp?w8=backup&islem=userbackupup">Bir Userin Bilgilerini Geri Yükle</a>
<%Dim users
If islem="userbackupup" Then
Set users=connx.execute("select * from userdata order by strUserID")%><table width="400"><tr><td style="font-weight:bold" align="center">Kullanıcı Adı</td><td style="font-weight:bold" >Son Yedeklenme Tarihi</td></tr><%do while not users.eof 
Response.Write "<tr><td><b>"&users(0)&"</b></td><td><b>"&users(19)&"</b></td><td><a href=""default.asp?w8=backup&islem=kullanicibackup&id="&users(0)&""">Geri Yükle</a></td>"
Users.MoveNext
Loop
ElseIf islem="userbackup" Then
Set users=Conne.Execute("select * from userdata order by strUserID")%><table width="400"><tr><td style="font-weight:bold" align="center">Kullanıcı Adı</td></tr><%do while not users.eof 
Response.Write "<tr><td><b>"&users(0)&"</b></td><td><a href=""default.asp?w8=backup&islem=kullanicibackupup&id="&users(0)&""">Yedekle</a></td>"
users.MoveNext
Loop
ElseIf islem="kullanicibackup" Then
Dim userid,usr,userbckp
userid=Request.Querystring("id")
Set usr=Connx.Execute("select * from userdata where struserId='"&userid&"'")
Set userbckp = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from userdata where struserId='"&userid&"'"
userbckp.Open SQL,conne,1,3
If userbckp.eof Then
userbckp.addnew
userbckp("struserId") = userid
End If
userbckp("Rank") = usr("Rank")
userbckp("Title") = usr("Title")
userbckp("Level") = usr("Level")
userbckp("Exp") = usr("Exp")
userbckp("Loyalty") = usr("Loyalty")
userbckp("Knights") = usr("Knights")
userbckp("Fame") = usr("Fame")
userbckp("Strong") = usr("Strong")
userbckp("Sta") = usr("Sta")
userbckp("Dex") = usr("Dex")
userbckp("Intel") = usr("Intel")
userbckp("Cha") = usr("Cha")
userbckp("Points") = usr("Points")
userbckp("Gold") = usr("Gold")
userbckp("strSkill") = usr("strSkill")
userbckp("strItem") = usr("strItem")
userbckp("strSerial") = usr("strSerial")
userbckp.Update
Response.Write "<br><br>"&userid &" Nickli Karakterin Bilgilerinin Yedekleri Geri Yüklenmiştir."

ElseIf islem="kullanicibackupup" Then
userid=Request.Querystring("id")
Set usr=Conne.Execute("select * from userdata where struserId='"&userid&"'")
connx.execute("delete userdata where struserId='"&userid&"'")
Set userbckp = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from userdata where struserId='"&userid&"'"
userbckp.Open SQL,connx,1,3
userbckp.addnew
userbckp("struserId") = userid
userbckp("Rank") = usr("Rank")
userbckp("Title") = usr("Title")
userbckp("Level") = usr("Level")
userbckp("Exp") = usr("Exp")
userbckp("Loyalty") = usr("Loyalty")
userbckp("Knights") = usr("Knights")
userbckp("Fame") = usr("Fame")
userbckp("Strong") = usr("Strong")
userbckp("Sta") = usr("Sta")
userbckp("Dex") = usr("Dex")
userbckp("Intel") = usr("Intel")
userbckp("Cha") = usr("Cha")
userbckp("Points") = usr("Points")
userbckp("Gold") = usr("Gold")
userbckp("strSkill") = usr("strSkill")
userbckp("strItem") = usr("strItem")
userbckp("strSerial") = usr("strSerial")
userbckp.Update
Response.Write "<br><br>"&userid &" Nickli Karakterin Bilgilerinin Yedekleri Alınmıştır."
ElseIf islem="backup" Then
Conne.Execute("exec USERDATA_BACKUP")
Response.Write "<br><br>Tüm Userlerin Bilgilerinin Yedekleri Alınmıştır."
ElseIf islem="backupup" Then
connx.execute("exec USERDATA_BACKUP")
Response.Write "<br><br>Tüm Userlerin Bilgilerinin Yedekleri Geri Yüklenmiştir."
End If
Else
Response.Redirect("default.asp")
End If
Case "manager"
If Session("durum")="esp" Then
islem=Request.Querystring("islem")%>
<script>function rdrct(url,msg){
return confirm(msg)
window.location=url;
}
</script>
<%Response.Write "<br><a onclick=""javascript:return confirm('Server Ve Bulunduğunuz Pencere Kapanacaktır.\nBu İşlemi Yapmak İstediğinize Eminmisiniz?')"" href=""default.asp?w8=manager&islem=pckapat""><img src=""../imgs/shutdown.gif"" border=""0"" align=""absmiddle""> Serverı Kapat</a><br><br><a onclick=""javascript:return confirm('Server Ve Bulunduğunuz Pencere Kapanacaktır.\nBu İşlemi Yapmak İstediğinize Eminmisiniz?')"" href=""default.asp?w8=manager&islem=pcrestart""><img src=""../imgs/down.gif"" align=""absmiddle"" border=""0""> Serverı Yeniden Başlat</a><br><br><a href=""#"" onclick=""javascript:window.open('indir.asp','','fullscreen=0,top='+screen.availHeight/4+',left='+screen.availWidth/4+',resizable=1,status=1,scrollbars=1,menubar=0,toolbar=0,height=300,width=500');return false"">Servera Dosya Upload Et</a><br><br><br><a href=""#"" onclick=""javascript:window.open('dosyaindir.asp','','fullscreen=0,top='+screen.availHeight/4+',left='+screen.availWidth/4+',resizable=1,status=1,scrollbars=1,menubar=0,toolbar=0,height=300,width=500');return false"">Servera Dosya İndir</a><br>"
Response.Write "<br><br><form action=""default.asp?w8=manager&islem=runprogram"" method=""post""><b>Program Yolu: </b><input type=""text"" name=""pyol"" size=""45"" value=""C:\""><input type=""submit"" value=""Çalıştır""><br><br><b>Program Açmak için Programın Tam Yolunu Yazmanız Gerekmektedir.<br>Örn: C:\windows\system32\notepad.exe</b>"
If islem="pckapat" Then

Set wsh = CreateObject("WScript.Shell")
wsh.run("C:\windows\system32\Shutdown.exe -s -t 00")

ElseIf islem="pcrestart" Then

Set wsh = CreateObject("WScript.Shell")
wsh.run("C:\windows\system32\Shutdown.exe -r -t 00")

ElseIf islem="runprogram" Then

Set wsh = CreateObject("WScript.Shell")
wsh.run(request.form("pyol"))

End If
Else
Response.Redirect("default.asp")
End If

Case "monster"
If Session("durum")="esp" Then
If Request.Querystring("islem")="" Then%>
<form action="default.asp?w8=monster&islem=search" method="post">
<table width="300">
<tr>
<td>Monster Adını Giriniz</td><td><input type="text" name="monsterid"></td><td><input type="submit" value="Ara"></td></tr>
</form>
<%ElseIf Request.Querystring("islem")="search" Then
Set mon=Conne.Execute("select * from k_monster where strname like '%"&Request.Form("monsterid")&"%' ")
If not mon.eof Then
Response.Write "<br><b><u>Monster Seçiniz: </u></b><br><br><table width=""250"" cellpadding=0>"
do while not mon.eof%>
<tr><td><a href="default.asp?w8=monster&islem=edit&monsterid=<%=mon("ssid")%>"><%=mon("strname")%></a></td><td><%=mon("ssid")%></td>
<%mon.movenext
Loop
Else
Response.Write "Monster Bulunamadı"
End If
ElseIf Request.Querystring("islem")="edit" Then
monsterid=Request.Querystring("monsterid")
Set mons=Conne.Execute("select * from k_monster where ssid="&monsterid)

If not mons.eof Then
Response.Write "<form action=""default.asp?w8=monster&islem=kayit"" method=""post""><table width=""400"" align=""left"">"
Response.Write "<tr><td>Monster Adı: </td><td><input type=""text"" name=""strname"" value="""&mons("strname")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Monster Kodu: </td><td><input type=""text"" name=""monsterkod"" value="""&mons("ssid")&"""><input type=""hidden"" name=""monsterssid"" value="""&mons("ssid")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Exp: </td><td><input type=""text"" name=""iexp"" value="""&mons("iexp")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Np: </td><td><input type=""text"" name=""iloyalty"" value="""&mons("iloyalty")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Para(Coin): </td><td><input type=""text"" name=""imoney"" value="""&mons("imoney")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Level: </td><td><input type=""text"" name=""sLevel"" value="""&mons("sLevel")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Hp: </td><td><input type=""text"" name=""iHpPoint"" value="""&mons("iHpPoint")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Mp: </td><td><input type=""text"" name=""sMpPoint"" value="""&mons("sMpPoint")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Atack: </td><td><input type=""text"" name=""sAtk"" value="""&mons("sAtk")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Defans: </td><td><input type=""text"" name=""sAc"" value="""&mons("sAc")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>Damage: </td><td><input type=""text"" name=""sDamage"" value="""&mons("sDamage")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>1 İtem: </td><td><input type=""text"" name=""iWeapon1"" value="""&mons("iWeapon1")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td>2 İtem: </td><td><input type=""text"" name=""iWeapon2"" value="""&mons("iWeapon2")&"""></td></tr>"&vbcrlf
Response.Write "<tr><td colspan=""2""><input type=""submit"" value=""KAYDET"" style=""width:280px;height:50px""></td></tr>"
Response.Write "</table>"
End If

Set mitem=Conne.Execute("select * from k_monster_item where sindex="&monsterid)
If not mitem.eof Then%>
<script type="text/javascript">
 function selectedbox(slot){
 for (var i=1; i<6; i++){
  var spn = document.getElementById(i);
  spn.style.border = "";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
	}
function itemozell(ssid,slot,itemno){
  $.ajax({
	type: 'GET',
	url: 'mitem.asp?ssid='+ssid+'&slot='+slot+'&itemno='+itemno,
	success: function(ajaxCevap) {
		$('div#itemoz').html(ajaxCevap);
	}
  });
}
</script>
<%Response.Write "<table width=""103"" align=""right"" border=""1"" cellspacing=""1"" style=""background-repeat:no-repeat;position:relative;right:350px;top:20px"" background=""../imgs/drop.gif"">"
Response.Write "<tr><td><img src=""../item/3.jpg"" onmouseover=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=#DFC68C>"&mons("imoney")&"', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();""></td>"
for x=1 to 5
itemno=mitem("iItem0"&x)
Set det=Conne.Execute("select itemtype,delay,num,kind,strname,damage,weight,duration,ac,Evasionrate,Hitrate,daggerac,swordac,axeac,spearac,bowac,maceac,firedamage,icedamage,LightningDamage,poisondamage,hpdrain,MPDamage,mpdrain,MirrorDamage,StrB,StaB,MaxHpB,DexB,IntelB,MaxMpB,ChaB,FireR,coldr,LightningR,magicr,poisonr,curser,reqstr,ReqSta,reqdex,reqintel,reqcha,Countable from ITEM where num='"&itemno&"'")
If not det.eof Then
dtype=det("ItemType")
speed=det("delay")
kinds=det("kind")

If speed>0 and speed<90 Then
delay = "Atack Speed : Very Fast<br>"
ElseIf speed>89 and speed<111 and not kinds=>91 and not kinds=<95 Then
delay = "Atack Speed : Fast<br>"
ElseIf speed>110 and speed<131 Then
delay = "Atack Speed : Normal<br>"
ElseIf speed>130 and speed<151 Then
delay = "Atack Speed : Slow<br>"
ElseIf speed>150 and speed<201 Then
delay = "Atack Speed : Very Slow<br>"
Else
delay=""
End If

If kinds=11 Then
kind="Dagger"
ElseIf kinds =21 Then
kind="One-handed Sword"
ElseIf kinds = 22 Then
kind="Two-handed Sword"
ElseIf kinds =31 Then
kind= "Axe"
ElseIf kinds = 32 Then
kind="Two-handed Axe"
ElseIf kinds = 41 Then
kind="Club"
ElseIf kinds = 42 Then
kind="Two-handed Club"
ElseIf kinds = 51 Then
kind="Spear"
ElseIf kinds = 52 Then
kind="Long Spear"
ElseIf kinds = 60 Then
kind="Shield"
ElseIf kinds = 70 Then
kind="Bow"
ElseIf kinds = 71 Then
kind="Crossbow"
ElseIf kinds = 91 Then
kind="Earring"
ElseIf kinds = 92 Then
kind="Necklace"
ElseIf kinds = 93 Then
kind="Ring"
ElseIf kinds = 94 Then
kind="Belt"
ElseIf kinds = 95 Then
kind="Lune Item"
ElseIf kinds = 98 Then
kind="Upgrade Scroll"
ElseIf kinds = 110 Then
kind="Staff"
ElseIf kinds = 120 Then
kind="Staff"
ElseIf kinds = 210 Then
kind="Warrior Armor"
ElseIf kinds = 220 Then
kind="Rogue Armor"
ElseIf kinds = 230 Then
kind="Magician Armor"
ElseIf kinds = 240 Then
kind="Priest Armor"
Else
kind=""
End If

If dtype=0 Then
If dtype=0 and kinds=255 Then
dtype="(Scroll)"
renk="white"
Else
dtype="(Non Upgrade Item)"
renk="white"
End If
ElseIf dtype=1 Then
dtype="(Magic Item)"
renk="blue"
ElseIf dtype=2 Then
dtype="(Rare Item)"
renk="yellow"
ElseIf dtype=3 Then
dtype="(Craft Item)"
renk="lime"
ElseIf dtype=4 Then
dtype="(Unique Item)"
renk="#DFC68C"
ElseIf dtype=5 Then
dtype="(Upgrade Item)"
renk="#CE8DC5"
ElseIf dtype=6 Then
dtype="(Event Item)"
renk="lime"
End If

If det("Damage")>0 Then 
atack="Attack Power : "&det("Damage") & "<br>"
Else
atack=""
End If
If det("Weight")>1 Then 
If len(det("Weight"))=2 Then
sf="00"
ElseIf len(det("weight"))=3 Then
sf="0"
End If
weight="Weight : "&left(det("Weight"),2)&"."&mid(det("weight"),3,1)&sf&"<br>"
Else
weight=""
End If

If det("Duration")>1 and det("ItemType")=0 Then
duration="Quantity : "&det("Duration") & "<br>"
ElseIf det("Duration")>1  Then 
duration="Max Durability : "&det("Duration") & "<br>"
Else
duration=""
End If
If det("Ac")>0 Then 
defans="Defense Ability : "&det("Ac") & "<br>"
Else
defans=""
End If
If det("Evasionrate")>0 Then
dodging="Increase Dodging Power by : "&det("Evasionrate")&"<br>"
Else
dodging=""
End If
If det("Hitrate")>0 Then
incap="Increase Attack Power by  : "&det("Hitrate")&"<br>"
Else
incap=""
End If


If det("DaggerAc")>0 Then 
daggerac="Defense Ability (Dagger) : "&det("DaggerAc") & "<br>"
Else
daggerac=""
End If
If det("SwordAc")>0 Then 
swordac="Defense Ability (Sword) : "&det("SwordAc") & "<br>"
Else
swordac=""
End If
If det("MaceAc")>0 Then 
clubac="Defense Ability (Club) : "&det("MaceAc") & "<br>"
Else
clubac=""
End If
If det("AxeAc")>0 Then 
axeac="Defense Ability (Axe) : "&det("AxeAc") & "<br>"
Else
axeac=""
End If
If det("SpearAc")>0 Then 
spearac="Defense Ability (Spear) : "&det("SpearAc") & "<br>"
Else
spearac=""
End If
If det("BowAc")>0 Then 
bowac="Defense Ability (Arrow) : "&det("BowAc") & "<br>"
Else
bowac=""
End If
If det("FireDamage")>0 Then 
firedam="Flame Damage : "&det("FireDamage") & "<br>"
Else
firedam=""
End If
If det("IceDamage")>0 Then 
icedam="Glacier Damage : "&det("IceDamage") & "<br>"
Else
icedam=""
End If
If det("LightningDamage")>0 Then 
ligthdam="Lightning Damage : "&det("LightningDamage") & "<br>"
Else
ligthdam=""
End If
If det("PoisonDamage")>0 Then 
posdam="Poison Damage : "&det("PoisonDamage") & "<br>"
Else
posdam=""
End If
If det("HPDrain")>0 Then
hpdrain="HP Recovery : "&det("HPDrain")&"<br>"
Else
hpdrain=""
End If
If det("MPDamage")>0 Then
mpdamage="MP Damage : "&det("MPDamage")&"<br>"
Else
mpdamage=""
End If
If det("MPDrain")>0 Then
mpdrain="MP Recovery : "&det("MPDrain")&"<br>"
Else
mpdrain=""
End If
If det("MirrorDamage")>0 Then 
mirrordam="Repel Physical Damage : "&det("MirrorDamage") & "<br>"
Else
mirrordam=""
End If
If det("StrB")>0 Then 
strbon="Strength Bonus : "&det("StrB") & "<br>"
Else
strbon=""
End If
If det("StaB")>0 Then 
canbonus="Health Bonus : "&det("StaB") & "<br>"
Else
canbonus=""
End If
If det("DexB")>0 Then 
dexbon="Dexterity Bonus : "&det("DexB") & "<br>" 
Else
dexbon=""
End If
If det("IntelB")>0 Then 
intbon="Intelligence Bonus : "&det("IntelB") & "<br>"
Else
intbon=""
End If
If det("ChaB")>0 Then 
magicbon="Magic Power Bonus : "&det("ChaB") & "<br>"
Else
magicbon=""
End If
If det("MaxHpB")>0 Then 
hpbon="HP Bonus : "&det("MaxHpB") & "<br>"
Else
hpbon=""
End If
If det("MaxMpB")>0 Then 
mpbon="MP Bonus : "&det("MaxMpB") & "<br>"
Else
mpbon=""
End If

If det("FireR")>0 Then 
fireres="Resistance to Flame : "&det("FireR") & "<br>"
Else
fireres=""
End If
If det("ColdR")>0 Then 
glares="Resistance to Glacier : "&det("ColdR") & "<br>"
Else
glares=""
End If
If det("LightningR")>0 Then 
lightres="Resistance to Lightning : "&det("LightningR") & "<br>"
Else
lightres=""
End If
If det("MagicR")>0 Then 
magicres="Resistance to Magic : "&det("MagicR") & "<br>"
Else
magicres=""
End If
If det("PoisonR")>0 Then 
posres="Resistance to Poison : "&det("PoisonR") & "<br>"
Else
posres=""
End If
If det("CurseR")>0 Then 
curseres="Resistance to Curse : "&det("CurseR") & "<br>"
Else
curseres=""
End If

If det("ReqStr")>0 Then 
If det("ReqStr")>userbilgi("strong") Then
reqstr="<font color=red>Required Strength : "&det("ReqStr")&"</font><br>"
Else
reqstr="Required Strength : "&det("ReqStr") & "<br>"
End If
Else
reqstr=""
End If
If det("ReqSta")>0 Then 
If det("ReqSta")>userbilgi("sta") Then
reqstr="<font color=red>Required Health : "&det("ReqSta") & "<br>"
Else
reqhp="Required Health : "&det("ReqSta") & "<br>"
End If
Else
reqhp=""
End If
If det("ReqDex")>0 Then 
If det("ReqDex")>userbilgi("dex") Then
reqstr="<font color=red>Required Dexterity : "&det("ReqDex") & "<br>"
Else
reqdex="Required Dexterity : "&det("ReqDex") & "<br>"
End If
Else
reqdex=""
End If
If det("ReqIntel")>0 Then 
If det("ReqIntel")>userbilgi("intel") Then
reqstr="<font color=red>Required Intelligence : "&det("ReqIntel") & "<br>"
Else
reqint="Required Intelligence : "&det("ReqIntel") & "<br>"
End If
Else
reqint=""
End If
If det("ReqCha")>0 Then 
If det("ReqCha")>userbilgi("cha") Then
reqstr="<font color=red>Required Magic Power : "&det("ReqCha") & "<br>"
Else
reqcha="Required Magic Power : "&det("ReqCha") & "<br>"
End If
Else
reqcha=""
End If

If det("Countable")=1 Then
drtn=itemler("stacksize")
ElseIf det("duration")>0 and det("ItemType")=0 and kinds <> 95 Then
drtn=itemler("durability")
Else
drtn=""
End If

iname=server.htmlencode(det("strname"))

If det("strb")=24 or det("stab")=24 or det("dexb")=24 or det("ChaB")=24 or det("intelb")=24 Then
itemname2=replace(iname, "(+0)" , "(+10)" )
Else
itemname2=iname
End If

itemname=server.htmlencode(itemname2)

items="<img width=""45"" height=""45"" src=""../item/"&resim2(itemno)&""" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&itemname&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&Durability&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

Response.Write "<td id="""&x&""" align=""center"" onclick=""selectedbox(this.id);itemozell('"&monsterid&"','"&x&"','"&itemno&"')"">"&items&"</td>"&vbcrlf
Else
Response.Write "<td id="""&x&""" align=""center"" onclick=""selectedbox(this.id);itemozell('"&monsterid&"','"&x&"','"&itemno&"')"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"&vbcrlf
End If
If x mod 2=1 Then
Response.Write "</tr>"&vbcrlf&"<tr>"&vbcrlf
End If
next
%></tr></table><div id="itemoz" align="center"></div></form>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<script>
function itemekle(rsm,num,dur,countable,charid){
var ids=document.getElementById('inventoryslot').value
document.getElementById(ids).innerHTML="<img src=../item/"+rsm+">";
document.getElementById('num').value=num
document.getElementById('dropyuzde').value='100'

var ssid=document.getElementById('ssid').value
var itemno=document.getElementById('num').value;
var dropyuzde=document.getElementById('dropyuzde').value;
  $.ajax({
	type: 'GET',
	url: 'mitemkaydet.asp?ssid='+ssid+'&slot='+ids+'&num='+itemno+'&dropyuzde='+dropyuzde
  });
}
function itemkayit()
{
var ids=document.getElementById('inventoryslot').value;
var ssid=document.getElementById('ssid').value;
var itemno=document.getElementById('num').value;
var dropyuzde=document.getElementById('dropyuzde').value;
  $.ajax({
	type: 'GET',
	url: 'mitemkaydet.asp?ssid='+ssid+'&slot='+ids+'&num='+itemno+'&dropyuzde='+dropyuzde
  });
}

function itemsil(slot)
{
var ids=document.getElementById('inventoryslot').value
document.getElementById(ids).innerHTML='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
document.getElementById('num').value=0;
document.getElementById('dropyuzde').value=0;
var ssid=document.getElementById('ssid').value;
  $.ajax({
	type: 'GET',
	url: 'mitemkaydet.asp?ssid='+ssid+'&slot='+ids+'&num=0&dropyuzde=0'
  });
}
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'bkeyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<form action="javascript:chng();"  method="post" id="search" name="search">
	<table style="position:relative;left:-100px">
	<tr align="center" >
	<td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr ><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option></select>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select>
	</tr><tr><td colspan="5">
      <input type="text" style="width:280px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr ><td colspan="5">
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
    </td>
	</td></tr> 
</table>
</form>
<%End If
ElseIf Request.Querystring("islem")="kayit" Then
strname=Request.Form("strname")
monsterkod=Request.Form("monsterkod")
monsterssid=Request.Form("monsterssid")
iexp=Request.Form("iexp")
iloyalty=Request.Form("iloyalty")
imoney=Request.Form("imoney")
sLevel=Request.Form("sLevel")
iHpPoint=Request.Form("iHpPoint")
sMpPoint=Request.Form("sMpPoint")
sAtk=Request.Form("sAtk")
sAc=Request.Form("sAc")
sDamage=Request.Form("sDamage")
iWeapon1=Request.Form("iWeapon1")
iWeapon2=Request.Form("iWeapon2")
Conne.Execute("Update k_monster Set strname='"&strname&"', sSid="&monsterkod&", iexp="&iexp&", iloyalty="&iloyalty&", imoney="&imoney&", sLevel="&sLevel&", iHpPoint="&iHpPoint&", sMpPoint="&sMpPoint&", sAtk="&sAtk&", sAc="&sAc&", sDamage="&sDamage&", iWeapon1="&iWeapon1&", iWeapon2="&iWeapon2&" where ssid="&monsterssid&"   ")
Response.Redirect("default.asp?w8=monster")

End If
Else
Response.Redirect("default.asp")
End If
Case "anket"
If Session("durum")="esp" Then
Dim Anket
Set Anket=Conne.Execute("select * from anket")
If Not anket.eof Then%>
<form action="default.asp?w8=anket&islem=kaydet" method="post">
Anket Sorusu: <input type="text" name="anketsoru" value="<%=trim(anket("anketsoru"))%>" size="50"><br>
Anket Seçenek 1: <input type="text" name="anketsec1" value="<%=trim(anket("anketsec1"))%>" size="30"><br>
Anket Seçenek 2: <input type="text" name="anketsec2" value="<%=trim(anket("anketsec2"))%>" size="30"><br>
Anket Seçenek 3: <input type="text" name="anketsec3" value="<%=trim(anket("anketsec3"))%>" size="30"><br>
Anket Seçenek 4: <input type="text" name="anketsec4" value="<%=trim(anket("anketsec4"))%>" size="30"><br>
Anket Seçenek 5: <input type="text" name="anketsec5" value="<%=trim(anket("anketsec5"))%>" size="30"><br>
<input type="submit" value="Kaydet">
</form>
<%If Request.Querystring("islem")="kaydet" Then
Conne.Execute("Update anket Set anketsoru='"&Request.Form("anketsoru")&"', anketsec1='"&Request.Form("anketsec1")&"', anketsec2='"&Request.Form("anketsec2")&"', anketsec3='"&Request.Form("anketsec3")&"', anketsec4='"&Request.Form("anketsec4")&"', anketsec5='"&Request.Form("anketsec5")&"' ")
Response.Redirect("default.asp?w8=anket")
End If
Else
Conne.Execute("insert into anket values('','','','','','')")
Response.Redirect("default.asp?w8=anket")
End If
Else
Response.Redirect("default.asp")
End If
Case "itemmall"
If Session("durum")="esp" Then%>
<script language="javascript">
function userara(){
$.ajax({
type: 'post',
url: 'userbul.asp',
data: $('#struserid').serialize(),
success: function(ajaxCevap) {
$('#userler').html(ajaxCevap);
}
});
}
function userbul(userid){
document.getElementById('struserid').value=userid
}
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'keyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
	$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<table>
<tr valign="top">
<td><form action="default.asp?w8=itemmall&s=2" method="POST" name="itemn" >
<font face="Verdana" style="font-size:9pt;"><b>Kullanıcı Adını Yazın: </b></font>
<input  type="text" name="struserid" id="struserid" onKeyUp="userara()" autocomplete="off">
<font face="Verdana" style="font-size:11px;"><b>İtem No Girin: </b></font>
<input type="text" name="itemno" id="itemn"><input type="submit" value="İtem Gönder">
<div name="userler" id="userler"></div>
</form></td>
  </tr>
    <tr valign="top">
    <td><form action="javascript:chng();"  method="post" id="search" name="search">
	<table >
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select></td>
	</tr></table>
      <input type="text" style="width:200px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
      <div id="sonuc" style="width:300px; background-color:silver;">.. Item İsmini Yazın</div>
<%If Request.Querystring("s")="2" Then
charid=Request.Form("struserid")
itemno=Request.Form("itemno")
Set accbl=Conne.Execute("select * from account_char where strcharid1='"&charid&"' or strcharid2='"&charid&"' or strcharid3='"&charid&"'")
If not accbl.eof Then
Set itemnk=Conne.Execute("select * from item where num="&itemno)
If not itemnk.eof Then
Conne.Execute("insert into WEB_ITEMMALL values('"&accbl("straccountid")&"','"&charid&"',1,"&itemno&",1,'"&now&"','','"&itemnk("strname")&"','','')")
Else
Response.Write "Item Bulunamadı"
End If
End If
End If
Else
Response.Redirect("default.asp")
End If
Case "pmbox"
If Session("durum")="esp" Then 
dim pmler
Set pmler=Conne.Execute("select * from pmbox order by tarih")
If pmler.eof Then
Response.Write "<br><b>Özel Mesaj Bulunamadı!</b>"
Else
Response.Write "<table width=""630""><tr><td colspan=""5"" style=""font-weight:bold""><a href=""default.asp?w8=pmgonder&kime=herkes"">Tüm Kullanıcılara Özel Mesaj Gönder</a></td></tr><tr><td colspan=""5"" style=""font-weight:bold""><a href=""default.asp?w8=pmgonder&kime=kullanici"">Kişiye Özel Mesaj Gönder</a></td></tr><tr><td align=""center"" colspan=""4"" style=""font-weight:bold"">Kullanıcıların Birbirlerine Gönderdikleri Özel Mesajlar</td></tr><tr><td style=""font-weight:bold"">Gönderen</td><td style=""font-weight:bold"">Alıcı</td><td style=""font-weight:bold"">Konu</td><td style=""font-weight:bold"">Tarih</td><td>&nbsp;</td></tr>"
do while not pmler.eof 
Response.Write "<tr><td>"&pmler("gonderen")&"</td><td>"&pmler("alici")&"</td><td>"&pmler("konu")&"</td><td>"&pmler("tarih")&"</td><td><a href=""default.asp?w8=pmoku&id="&pmler("id")&""">OKU</a></td></tr>"
pmler.movenext
Loop
Response.Write "</table>"
End If
Else
Response.Redirect("default.asp")
End If
Case "pmgonder"
dim kime
kime=Request.Querystring("kime")
If kime="kullanici" Then
dim uses
Set uses=Conne.Execute("select struserid from userdata order by struserid")

%><form action="default.asp?w8=pmgonder&kime=kullanici&islem=pmkayit" method="post">
<table>
<tr>
<td><b>Alıcının Hesap Adı: </b><select name="accid">
<%If not uses.eof Then
do while not uses.eof
Response.Write "<option value="""&uses("struserid")&""">"&uses("struserid")&"</option>"
uses.movenext
Loop
End If%></td>
</tr>
<tr><td><b>Konu:</b> <input type="text" name="konu" size="50">
</td></tr>
<tr><td><textarea name="pmgonder" id="pmgonder"></textarea>
       <script language="JavaScript">
  generate_wysiwyg('pmgonder');
</script>
</td></tr>

<tr><td>
<input type="submit" value="Pm Gönder" class="inputstyle"  style="width:600; font-size:11px;">
</td></tr>
</table>
</form>
<%
If Request.Querystring("islem")="pmkayit" Then

Set pmekle = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from pmbox"
pmekle.Open SQL,conne,1,3
pmekle.addnew
pmekle("gonderen")="Game Master"
pmekle("alici")=Request.Form("accid")
pmekle("konu")=Request.Form("konu")
pmekle("mesaj")=Request.Form("pmgonder")
pmekle.Update
pmekle.close
Set pmekle=nothing

End If
ElseIf kime="herkes" Then%>
<form action="default.asp?w8=pmgonder&kime=herkes&islem=pmkayit" method="post">
<table>

<tr><td><b>Konu:</b> <input type="text" name="konu" size="50">
</td></tr>
<tr><td><textarea name="pmgonder" id="pmgonder"></textarea>
       <script language="JavaScript">
  generate_wysiwyg('pmgonder');
</script>
</td></tr>

<tr><td>
<input type="submit" value="Pm Gönder" class="inputstyle"  style="width:600; font-size:11px;">
</td></tr>
</table>
</form>
<%If Request.Querystring("islem")="pmkayit" Then
dim kullanicilari,pmekle,kullanicilar
Set kullanicilar=Conne.Execute("select struserid from userdata order by struserid")
do while not kullanicilar.eof
Set pmekle = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from pmbox"
pmekle.Open SQL,conne,1,3
pmekle.addnew
pmekle("gonderen")="Game Master"
pmekle("alici")=kullanicilar("struserid")
pmekle("konu")=Request.Form("konu")
pmekle("mesaj")=Request.Form("pmgonder")
pmekle.Update
pmekle.close
Set pmekle=nothing
kullanicilar.movenext
Loop
End If
End If
Case "pmoku"
If Session("durum")="esp" Then
id=Request.Querystring("id")
dim pmok
Set pmok=Conne.Execute("select * from pmbox where id="&id)
Response.Write "<table width=""600""><tr><td align=""center"" colspan=""4"" style=""font-weight:bold"">Kullanıcıların Birbirlerine Gönderdikleri Özel Mesajlar</td></tr><tr><td style=""font-weight:bold"">Gönderen</td><td style=""font-weight:bold"">Alıcı</td><td style=""font-weight:bold"">Konu</td><td style=""font-weight:bold"">Tarih</td><td>&nbsp;</td></tr><tr><td>"&pmok("gonderen")&"</td><td>"&pmok("alici")&"</td><td>"&pmok("konu")&"</td><td>"&pmok("tarih")&"</td></tr><tr><td colspan=""5""><br><b>Mesaj</b><br><br>"&pmok("mesaj")&"</td></tr></table>"
Else
Response.Redirect "default.asp"
End If
Case "logs"
If Session("durum")="esp" Then 
If Request.Querystring("islem")="temizle" Then
Conne.Execute("truncate table logs")
Response.Redirect("default.asp?w8=logs")
End If
dim logs
Set logs=Conne.Execute("select * from logs order by id")
If not logs.eof Then
Response.Write "<table width=""700"" ><tr><td width=""100"" align=""center""><b>Ip</b></td><td width=""400"" align=""center""><b>İşlem</b></td><td width=""120"" align=""center""><b>İşlem Tarihi</b></td><td width=""10"">&nbsp;</td></tr>"
do while not logs.eof
Response.Write "<tr><td>"&logs("ip")&"</td><td>"&logs("islem")&"</td><td>"&logs("islemtarihi")&"</td><td><a href=""default.asp?w8=logs&islem=sil&id="&logs("id")&""">Sil</a></td></tr>"
logs.movenext
Loop
Response.Write "<tr><td><a href=""default.asp?w8=logs&islem=temizle"">Kayıtları Temizle</a></td></tr></table>"
Else
Response.Write "<br><b>Kayıt Bulunamadı.</b>"
End If

If Request.Querystring("islem")="sil" and Request.Querystring("id")<>"" Then
Conne.Execute("delete logs where id="&Request.Querystring("id")&"")
Response.Redirect("default.asp?w8=logs")
End If

Else
Response.Redirect("default.asp")
Response.End
End If

Case "bankabaslangic"
If Session("durum")="esp" Then %>


    <script language="javascript">
    function icerikal(){
    $.ajax({
	type: 'GET',
	url: 'bankabaslangicitemleribul.asp',
	start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
	success: function(ajaxCevap) {
	$('#itemler').html(ajaxCevap);
       }
    });
    }
    </script>

   <b>Banka Başlangıç İtem Editörü</b><br><br>
<input type="button" value="Banka Başlangıç Itemlerini Bul" onClick="icerikal();" style="background:black;color:white"/>
</form><span name="itemler" id="itemler"></span>
<%
Else 
End If 
Case "baslangic"
If Session("durum")="esp" Then%>
 <script language="javascript">
    function icerikal(){
    $.ajax({
       type: 'GET',
       url: 'baslangicitembul.asp',
       data: 'class='+ encodeURI( document.getElementById("class").value ),
	   start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
    }


    </script>
<form action="javascript:icerikal()" method="post"><b>Başlangıç Itemleri Düzenle</b><br><br>
Class Seçiniz: <select name="class" id="class" >
<option value="1">Warrior</option>
<option value="2">Rogue</option>
<option value="3">Mage</option>
<option value="4">Priest</option>
</select> <input type="submit" value="Başlangıç Düzenle">
</form>
<span name="itemler" id="itemler"></span>
<script src="_inc/jquery.hotkeys.js"></script>
<script language="javascript">
function icerikal(){
if(document.getElementById('but')!=null&&document.getElementById('but').disabled==false){
if(confirm('\n\n\n\n\n\n\n\n\n\n\n                 ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! \n>>>>>>>>> INVENTORYDEKI ITEMLERI KAYIT ETMEDEN ÇIKIYORSUNUZ. <<<<<<<<<<<<<<<<<\n>>>>>EĞER KAYIT ETMEDEN ÇIKARSANIZ YAPTIĞINIZ DEĞİŞİKLİKLER GEÇERLİ OLMAZ<<<<<   \n>>>>>>>>>>>> ITEMLERI KAYIT ETMEK IÇIN VAZGEÇE TIKLAYINIZ<<<<<<<<<<<<<<<<<\n\n\n\n\n\n\n\n\n\n\n')){
$.ajax({
       type: 'GET',
       url: 'baslangicitembul.asp',
       data: 'class='+ encodeURI( document.getElementById("class").value ),
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
       $('#itemler').html(ajaxCevap);
       }
    });
document.getElementById("class").blur();
}
else{
return false;
}
}
$.ajax({
       type: 'GET',
       url: 'baslangicitembul.asp',
       data: 'class='+ encodeURI( document.getElementById("class").value ),
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
}

function loadpage(syf){
 $.ajax({
       type: 'GET',
       url: syf,
       start: $('#itemler').html('<img src="../imgs/38-1.gif"><br>Itemler Yükleniyor...'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
$('#chrs').fadeOut(0);
    }

function userara(){
$.ajax({
   type: 'post',
   url: 'userbul.asp',
   data: 'struserid='+encodeURI( document.getElementById("charid").value ),
   start:$('#chrs').html("<center><img src='../imgs/38-1.gif'><br>Aranıyor. Lütfen Bekleyin.</center>"),
   success: function(ajaxCevap) {
      $('#chrs').html(ajaxCevap);
   }
});
$('#chrs').fadeIn(0);
}
function userbul(userid){
document.getElementById('charid').value=userid;
}

function kontrolet(olay){
olay = olay || event;
if (olay.keyCode==13){
icerikal();
$('#chrs').fadeOut(0);
}
else{
userara();
$('#chrs').fadeIn(0);
}
}


 function domo(){
jQuery(document).bind('keydown', 'Ctrl+s',function (evt){itemkayitall();document.getElementById('but').disabled=true; return false;});
jQuery(document).bind('keydown', 'Ctrl+c',function (evt){$cnum=$('#num').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#cicon').val($cicon);
$('#citemmname').val($itemmname);
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+x',function (evt){$cnum=$('#num').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#citemmname').val($itemmname);
$('#cicon').val($cicon);
$('#num').val('0');
$('#dur').val('0');
$('#stacksize').val('0');
$('div#itemmname').html('');
itemkayit();
$('#'+$cinventoryslot).html('<img height="45" width="45" src="../imgs/blank.gif">');
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+v',function (evt){$cnum=$('#cnum').val();
$cdur=$('#cdur').val();
$cstacksize=$('#cstacksize').val();
$cinventoryslot=$('#cinventoryslot').val();
$inventoryslot=$('#inventoryslot').val();
$cicon=$('#cicon').val();
$itemmname=$('#citemmname').val();
if($cicon=='')$cicon='<img height="45" width="45" src="../imgs/blank.gif">';
document.getElementById('but').disabled=false;
$('#num').val($cnum);
$('#dur').val($cdur);
$('#stacksize').val($cstacksize);
$('#itemmname').html($itemmname);
$('#'+$inventoryslot).html($cicon);
itemkayit();
return false;
});


jQuery(document).bind('keydown', 'del',function (evt){var ids=document.getElementById('inventoryslot').value;
	document.getElementById('but').disabled=false;
	itemsil(ids);
});


jQuery(document).bind('keydown', 'left',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==0){slt=42}
	selectedbox(slt-1);
	itemozell('',slt-1);
});

jQuery(document).bind('keydown', 'right',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==41){slt=-1}
	selectedbox(slt+1);
	itemozell('',slt+1);
});

jQuery(document).bind('keydown', 'up',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt<14&&slt>0){slt=slt-3}
	if (slt<21&&slt>13){slt=13}
	if (slt>20&&slt<42){slt=slt-7}
	selectedbox(slt);
	itemozell('',slt);
});

jQuery(document).bind('keydown', 'down',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt<11&&slt>=0){slt=slt+3}
	else if(slt==11){slt=13} 
	else if (slt==13||slt==12){slt=14}
	else if (slt>13&&slt<35){slt=slt+7}
	else if (slt>34&&slt<42){slt=0}
	selectedbox(slt);
	itemozell('',slt);
});

}
            
            
    jQuery(document).ready(domo);

	function itemsil(slot){
	var ids=document.getElementById('inventoryslot').value
	document.getElementById(ids).innerHTML='<img src="../imgs/blank.gif">';
	document.getElementById('num').value='0';
	document.getElementById('dur').value='0';
	document.getElementById('stacksize').value='0';
	cal('ui_button2.mp3');
	itemkayit();
	document.getElementById('itemmname').innerHTML='&nbsp;';
	document.getElementById('but').disabled=false;
	if (slot>13){
	document.getElementById('adet'+slot).innerHTML='&nbsp;';
		}
	}


function selectedbox(slot){
 for (var i=0; i<42; i++){
  var spn = document.getElementById(i);
  spn.style.border = "1px solid #4C4B36";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
	}



function itemkayit(){
  $.ajax({
	type: 'post',
	url:'baslangicitemkaydet.asp?islem=one',
	data:'num='+$('#num').val()+'&charid='+$('#charid').val()+'&dur='+$('#dur').val()+'&stacksize='+$('#stacksize').val()+'&inventoryslot='+$('#inventoryslot').val() 
  });
}




function itemkayitall(){
  $.ajax({
	type: 'post',
	url: 'baslangicitemkaydet.asp?islem=all',
	data: 'num='+$('#num').val()+'&charid='+$('#charid').val()+'&dur='+$('#dur').val()+'&stacksize='+$('#stacksize').val()+'&inventoryslot='+$('#inventoryslot').val()
  });
}



 
function loadpage(syf){
   $.ajax({
   type: 'GET',
   url: syf,
   start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
   success: function(ajaxCevap) {
   $('#itemler').html(ajaxCevap);
    }
    });
    }

function itemekle(rsm,num,dur,countable,name,detay,slot){

if (slot==5||slot==6||slot==7||slot==8||slot==9){
cal('item_armor.mp3')
}
else if(slot==1||slot==2||slot==3||slot==4){
cal('item_weapon.mp3')
}
else{
cal('ui_button2.mp3')
}

var ids=document.getElementById('inventoryslot').value;
document.getElementById(ids).innerHTML="<img src=\"../item/"+rsm+"\" onMouseOver=\"return overlib('"+detay+"', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10)\" onMouseOut=\"return nd();\">";
document.getElementById('num').value=num;
document.getElementById('dur').value=dur;
document.getElementById('stacksize').value=countable;

if (ids>13){
document.getElementById('adet'+ids).innerHTML='&nbsp;';
}
document.getElementById('itemmname').innerHTML='<b>'+name+'</b>';
itemkayit()
document.getElementById('but').disabled=false
}


function dis(){
document.getElementById('but').disabled=true
}

window.onbeforeunload = cikis_yap


function cikis_yap(){
if (document.getElementById('but').disabled==false){
return ('\n\n\n\n\n\n\n\n\n\n\n                 ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! \n>>>>>>>>> INVENTORYDEKI ITEMLERI KAYIT ETMEDEN ÇIKIYORSUNUZ. <<<<<<<<<<<<<<<<<\n>>>>>EĞER KAYIT ETMEDEN ÇIKARSANIZ YAPTIĞINIZ DEĞİŞİKLİKLER GEÇERLİ OLMAZ<<<<<   \n>>>>>>>>>>>> ITEMLERI KAYIT ETMEK IÇIN VAZGEÇE TIKLAYINIZ<<<<<<<<<<<<<<<<<\n\n\n\n\n\n\n\n\n\n\n')

}
}


function stacksizeupdate(slot,no){
if (slot>13){
if (no==''||no==0||no=='0'){
no='&nbsp;'
}
document.getElementById('adet'+slot).innerHTML=no
}
}

</script>
<%Else
Response.Redirect "default.asp"
End If
Case "town"
If Session("durum")="esp" Then
islem=Request.Querystring("islem")
If islem="" Then
islem=1
End If
If islem=1 Then
Set tn=Conne.Execute("select * from start_position") %>
<b>Town Koordinat Editor</b><br><br>
<form action="default.asp?w8=town&islem=2" method="post">
<select name="zoneno">
<%do while not tn.eof
Set tw=Conne.Execute("select * from zone_info where zoneno="&tn("zoneid")&"")
Response.Write "<option value="&tn("zoneid")&">"&tw("bz")&"</option>"
tn.movenext
Loop
Response.Write "</select><input type=""submit"" value=""Düzenle""></form>"
ElseIf islem="2" Then
zoneno=Request.Form("zoneno")
Set zoneinfo=Conne.Execute("select * from start_position where zoneid="&zoneno&"")
Response.Write "<b>Town Koordinat Editor</b><br><br><form action=""default.asp?w8=town&islem=3"" method=""post""><input type=""hidden"" name=""zoneno"" value="""&zoneno&""">"
Response.Write "Karus X: <input type=""text"" name=""karusx"" value="&zoneinfo("skarusx")&"><br>"
Response.Write "Karus Y: <input type=""text"" name=""karusy"" value="&zoneinfo("skarusz")&"><br>"
Response.Write "ElMorad X: <input type=""text"" name=""elmoradx"" value="&zoneinfo("selmoradx")&"><br>"
Response.Write "ElMorad Y: <input type=""text"" name=""elmorady"" value="&zoneinfo("selmoradz")&"><br>"
Response.Write "Menzil(Max Uzaklık) X: <input type=""text"" name=""rangex"" value="&zoneinfo("brangex")&"><br>"
Response.Write "Menzil(Max Uzaklık) Y: <input type=""text"" name=""rangez"" value="&zoneinfo("brangez")&"><br>"
Response.Write "<input type=""submit"" value=""Kaydet"""
Response.Write "</form>"
ElseIf islem="3" Then 
karusx=Request.Form("karusx")
karusy=Request.Form("karusy")
elmoradx=Request.Form("elmoradx")
elmorady=Request.Form("elmorady")
rangex=Request.Form("rangex")
rangez=Request.Form("rangez")
zoneno=Request.Form("zoneno")
Conne.Execute("Update start_positionSetsKarusX="&karusx&" , sKarusZ="&karusy&" , sElmoradX="&elmoradx&" , sElmoradZ="&elmorady&"  , bRangeX='"&rangex&"' , bRangeZ='"&rangez&"' where ZoneID="&zoneno&" ")
Response.Redirect("default.asp?w8=town")
End If
Else
Response.Redirect("default.asp")
End If
Case "ozelitem"
If Session("durum")="esp" Then
islem=Request.Querystring("islem")
If islem="" Then%>
<script type="text/javascript">
function chng(val)
{
  $.ajax({
	type: 'GET',
	url: 'keyw.asp',
	data: $('#keyw').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<table width="303" border="0" align="left">
<form action="default.asp?w8=ozelitem&islem=bul" method="post">
  <tr>
    <td colspan="2" align="center"><strong>Özel İtem Title Editörü</strong></td>
    </tr>
  <tr>
    <td width="114">İtem kodunu yazınız:</td>
    <td width="179"><input type="text" name="itemnum"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Ara"> </td>    
  </tr>
</form>
</table>
<form action="default.asp?w8=ozelitem&islem=kulbul" method="post">
<table width="303" border="0" align="right">
  <tr>
    <td colspan="2" align="center"><strong>Karakter Title Editörü</strong></td>
    </tr>
  <tr>
    <td width="114">Karakter Nick Yazınız:</td>
    <td width="179"><input type="text" name="username"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Ara"> </td>
    
  </tr>
</table><br>
</form><br><br>
<script type="text/javascript">
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'keyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<form action="javascript:chng();"  method="post" id="search" name="search">
<table align="left">
<tr valign="top">
    <td>
	<table >
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select>
	</tr></table>
      <input type="text" style="width:200px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr><td>
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
    </td>
	</tr> 
</table>
</form>
<br><br><br><br><br><br>
<br>Not: Title verdiğiniz karakterin itemi giyebilmesi için itemin title değeriyle aynı yapmanız gerekmektedir.<br><br>
<table align="center">
<tr>
<td width="700" align="center" colspan="5"><b>Özel Itemli Kullanıcılar</b></td>
</tr>
<tr><b><td><b>Kullanıcı Adı</td><td><b>Item Num</td><td><b>Item Serial</td><td><b>Alış Zamanı</td><td><b>Bitiş Zamanı</b></td></tr>
<%Set preitem=Conne.Execute("select * from privateitems order by strname")
do while not preitem.eof%>
<tr>
<td><%=preitem("strname")%></td>
<td><%=preitem("num")%></td>
<td><%=preitem("serial")%></td>
<td><%=preitem("aliszamani")%></td>
<td><%=preitem("aliszamani")+preitem("sure")%></td>
<td></td>
</tr>
<%preitem.movenext
Loop%>
</table>
<%ElseIf islem="kulbul" Then
username=Request.Form("username")
Set userbul=Conne.Execute("select * from userdata where struserid='"&username&"'")
If userbul.eof Then
Response.Write "<br><b>Karakter Bulunamadı</b>"
Response.End
End If%><br><br>
<form action="default.asp?w8=ozelitem&islem=kulkaydet" method="post">
<table width="200" align="center">
<tr>
<td colspan="2" align="center"><b>User Title Editör</b></td>
</tr>
<tr>
<td width="75">Kullanıcı Adı: </td>
<td><%=username%><input type="hidden" value="<%=username%>" name="username"></td>
</tr>
<tr>
<td>Title: </td>
<td><input type="text" value="<%=userbul("title")%>" name="title"></td>
</tr>
<tr><td colspan="2" align="center"><input type="submit" value="Kaydet"></td></tr>
</table></form>
<%ElseIf islem="kulkaydet" Then
username=Request.Form("username")
title=Request.Form("title")
Set gunc=Conne.Execute("Update userdata Set title="&title&" where struserid='"&username&"'")
Response.Write "<br><b>Karakter Güncellendi!</b>"
ElseIf islem="bul" Then
itemnum=Request.Form("itemnum")
Set ara=Conne.Execute("select num,reqtitle,strname,class from item where num="&itemnum&"")
%>
<form action="default.asp?w8=ozelitem&islem=kaydet" method="post">
<table width="200" border="0">
  <tr>
    <td colspan="2" align="center"><%=ara("strname")%><input type="hidden" value="<%=ara("num")%>" name="itemnum"></td>
      </tr>
  <tr>
    <td>Item Title</td>
    <td><input type="text" value="<%=ara("reqtitle")%>" name="title"></td>
  </tr>
  <tr>
    <td>Class</td>
    <td><select name="cla">
    <optgroup label="Her Class Kullanabilir">
	<option value="0" <%If ara("class")=0 Then Response.Write "selected"%>>Her Class</option>
</optgroup>
<optgroup label="Warrior">
<option value="1" <%If ara("class")=1 Then Response.Write "selected"%>>Warrior 1-10 lwl</option>
<option value="5" <%If ara("class")=5 Then Response.Write "selected"%>>Warrior 10-60 lwl</option>
<option value="6" <%If ara("class")=6 Then Response.Write "selected"%>>Warrior 60-80 lwl</option>
</optgroup>
<optgroup label="Rogue">
<option value="2" <%If ara("class")=2 Then Response.Write "selected"%>>Rogue 1-10 lwl</option>
<option value="7" <%If ara("class")=7 Then Response.Write "selected"%>>Rogue 10-60 lwl</option>
<option value="8" <%If ara("class")=8 Then Response.Write "selected"%>>Rogue 60-80 lwl</option>
</optgroup>
<optgroup label="Priest">
<option value="4" <%If ara("class")=4 Then Response.Write "selected"%>>Priest 1-10 lwl</option>
<option value="11" <%If ara("class")=11 Then Response.Write "selected"%>>Priest 10-60lwl</option>
<option value="12" <%If ara("class")=12 Then Response.Write "selected"%>>Priest 60-80 lwl</option>
</optgroup>
<optgroup label="Mage">
<option value="3" <%If ara("class")=3 Then Response.Write "selected"%>>Mage 1-10 lwl</option>
<option value="9" <%If ara("class")=9 Then Response.Write "selected"%>>Mage 10-60 lwl</option>
<option value="10" <%If ara("class")=10 Then Response.Write "selected"%>>Mage 60-80 lwl</option>
</optgroup>
</select>
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Kaydet"></td>
    
  </tr>
</table>
</form>

<%ElseIf islem="kaydet" Then
itemnum=request.Form("itemnum")
clas=Request.Form("cla")
title=request.Form("title")
Set x=Conne.Execute("Update item Set reqtitle="&title&", class="&clas&" where num="&itemnum&"")
Response.Write "<br>Item Güncellendi bu Itemi giymesini itediğiniz oyuncunun title değerini "&title&" Yapmalısınız" 
End If
Else
Response.Redirect("default.asp")
End If
Case "news"
If Session("durum")="esp" Then
dim events,habers,haberler,evnts
events=Request.Querystring("events")
habers=secur(Request.Querystring("haber"))
Set haberler = Conne.Execute("Select * from haberler")
Set evnts = Conne.Execute("Select * from events")
If habers="" Then%>
<table width="386" border="0">
<tr>
<td colspan="3" align="center"><a href="default.asp?w8=news&haber=ekle">Haber ekle</a></td>
</tr>
<tr><td width="161">&nbsp;</td>
</tr>
<%If not haberler.eof Then
do while not haberler.eof%>
  <tr>
    <td><%=haberler("baslik")%></td>
    <td width="119"><%=haberler("tarih")%></td>
    <td width="60"><a href="default.asp?w8=news&haber=duzenle&id=<%=haberler("id")%>">Düzenle</a></td>
    <td width="28"><a href="default.asp?w8=news&haber=sil&id=<%=haberler("id")%>">Sil</a></td>
  </tr>

<% haberler.movenext
Loop
Else
Response.Write "Haber bulunamadı"
End If %>
</table><br>

<table width="486" border="0">
<tr>
<td colspan="3" align="center"><a href="default.asp?w8=news&events=ekle">Event ekle</a></td>
</tr>

<%If not evnts.eof Then
do while not evnts.eof%>
  <tr>
    <td><%=evnts(1)%></td>
    <td width="119"><%=evnts(2)%></td>
    <td width="60"><a href="default.asp?w8=news&events=duzenle&id=<%=evnts(0)%>">Düzenle</a></td>
    <td width="28"><a href="default.asp?w8=news&events=sil&id=<%=evnts(0)%>">Sil</a></td>
  </tr>
<% evnts.movenext
Loop
Else
Response.Write "Event bulunamadı"
End If %>
</table>
<% ElseIf habers="ekle" Then%>
<form action="default.asp?w8=news&haber=kaydet" method="post">
<table width="200" border="0">
  <tr>
    <td>Başlık</td>
    <td><input type="text" name="baslik"></td>
  </tr>
  <tr>
    <td>Haber</td>
    <td><textarea name="haber" id="haber"></textarea></td>
  </tr>
  <tr>
    <td>Gönderen</td>
    <td><input type="text" name="gonderen"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Ekle"></td>
    </tr>
</table>
</form>
       <script language="JavaScript">
  generate_wysiwyg('haber');
</script> 

<%ElseIf habers="duzenle" Then
id=secur(Request.Querystring("id"))
dim haberduzen
Set haberduzen=Conne.Execute("select * from haberler where id='"&id&"'")%>
<form action="default.asp?w8=news&haber=kaydet2&id=<%=id%>" method="post">
<table width="200" border="0">
  <tr>
    <td>Başlık</td>
    <td><input type="text" name="baslik" value="<%=haberduzen("baslik")%>"></td>
  </tr>
  <tr>
    <td>Haber</td>
    <td><textarea name="haber" id="haber"><%=haberduzen("haber")%></textarea></td>
  </tr>
  <tr>
    <td>Gönderen</td>
    <td><input type="text" name="gonderen" value="<%=haberduzen("gonderen")%>"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Ekle"></td>
    </tr>
</table>
</form>
       <script language="JavaScript">
  generate_wysiwyg('haber');
</script> 
<%ElseIf habers="kaydet2" Then
id=secur(Request.Querystring("id"))
dim haberler2,baslik,haber,gonderen
Set haberler2 = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from haberler where id='"&id&"'"
haberler2.Open SQL,conne,1,3

baslik=secur(Request.Form("baslik"))
haber=secur(Request.Form("haber"))
gonderen=secur(Request.Form("gonderen"))

haberler2("baslik")=baslik
haberler2("haber")=haber
haberler2("gonderen")=gonderen
haberler2("tarih")= date()
haberler2.Update
haberler2.close
Set haberler2=nothing
Response.Redirect "default.asp?w8=news"

ElseIf habers="kaydet" Then
Set haberler = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from haberler"
haberler.Open SQL,conne,1,3
haberler.addnew
haberler("baslik")=Request.Form("baslik")
haberler("haber")=Request.Form("haber")
haberler("gonderen")=Request.Form("gonderen")
haberler("tarih")= date()
haberler.Update

Response.Redirect "default.asp?w8=news"
ElseIf habers="sil" Then
id=secur(Request.Querystring("id"))
Conne.Execute("delete haberler where id='"&id&"'")
Response.Redirect "default.asp?w8=news"
End If
If events="duzenle" Then
id=Request.Querystring("id")
dim evg
Set evg=Conne.Execute("select * from events where id="&id)
%><br><br>
<form action="default.asp?w8=news&events=kaydet&id=<%=id%>" method="post">
<table width="300" border="0">
  <tr>
    <td>Event </td>
    <td><input type="text" name="1" value="<%=evg(1)%>" size="50"></td>
  </tr>
  <tr>
    <td>Tarih </td>
    <td><input type="text" name="2" value="<%=evg(2)%>"></td>
  </tr>

  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Düzenle"></td>
    </tr>
</table>
</form>
<%ElseIf events="kaydet" Then
id=Request.Querystring("id")
dim events2
Set events2 = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from events where id='"&id&"'"
events2.Open SQL,conne,1,3
events2(1)=Request.Form("1")
events2(2)=Request.Form("2")
events2.Update
events2.close
Set events2=nothing
Response.Redirect "default.asp?w8=news"
ElseIf events="ekle" Then%><br>
<form action="default.asp?w8=news&events=kaydet2&id=<%=id%>" method="post">
<table width="300" border="0">
  <tr>
    <td>Event </td>
    <td><input type="text" name="1"  size="50"></td>
  </tr>
  <tr>
    <td>Tarih </td>
    <td><input type="text" name="2" value="<%=date%>"></td>
  </tr>

  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Ekle"></td>
    </tr>
</table>
</form>
<%ElseIf events="kaydet2" Then
Set events2 = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from events"
events2.Open SQL,conne,1,3
events2.addnew
events2(1)=Request.Form("1")
events2(2)=Request.Form("2")
events2.Update
events2.close
Response.Redirect "default.asp?w8=news"
ElseIf events="sil" Then
id=secur(Request.Querystring("id"))
Conne.Execute("delete events where id='"&id&"'")
Response.Redirect "default.asp?w8=news"
End If
Else
Response.Redirect "default.asp"
End If
'--------------------------------------------------------------------------------------------------------
Case "filecontrol"
If Session("durum")="esp" Then
dim files,dosya,sql
files=secur(Request.Querystring("files"))
Set dosya = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from download"
dosya.Open SQL,conne,1,3

If files="" Then%>
<table width="325" border="0">
<tr>
<td colspan="3" align="center"><a href="default.asp?w8=filecontrol&files=ekle"><span class="style5">Dosya ekle</span></a></td>
</tr>
<tr><td width="144">&nbsp;</td>
</tr>
<%If not dosya.eof Then
do while not dosya.eof%>
  <tr>
    <td><%=dosya("dosyaismi")%></td>
    <td width="101"><%=dosya("tarih")%></td>
	<td width="48"><a href="default.asp?w8=filecontrol&files=duzenle&id=<%=dosya("id")%>">Düzenle</a></td>
    <td width="14"><a href="default.asp?w8=filecontrol&files=sil&id=<%=dosya("id")%>">Sil</a></td>
  </tr>

<% dosya.movenext
Loop
Else
Response.Write "Dosya bulunamadı"
End If %>
</table>
<% ElseIf files="ekle" Then%>
<form action="default.asp?w8=filecontrol&files=kaydet" method="post">
<table width="200" border="0">
  <tr>
    <td>Dosya ismi</td>
    <td><input type="text" name="dosyaismi"></td>
  </tr>
  <tr>
    <td>Açıklama</td>
    <td><textarea name="aciklama" id="aciklama"></textarea></td>
  </tr>
  <tr>
    <td>Adres</td>
    <td><input type="text" name="adres"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Ekle"></td>
    </tr>
</table>
</form>
       <script language="JavaScript">
  generate_wysiwyg('aciklama');
</script> 

<%ElseIf files="duzenle" Then
id=secur(Request.Querystring("id"))
Dim dosyaduzen
Set dosyaduzen=Conne.Execute("select * from download where id='"&id&"'")%>
<form action="default.asp?w8=filecontrol&files=kaydet2&id=<%=id%>" method="post">
<table width="200" border="0">
  <tr>
    <td>Dosya ismi</td>
    <td><input type="text" name="dosyaismi" value="<%=dosyaduzen("dosyaismi")%>"></td>
  </tr>
  <tr>
    <td>Açıklama</td>
    <td><textarea name="aciklama" id="aciklama"><%=dosyaduzen("aciklama")%></textarea></td>
  </tr>
  <tr>
    <td>Adres</td>
    <td><input type="text" name="adres" value="<%=dosyaduzen("adres")%>"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input  type="submit" class="inputstyle" value="Ekle"></td>
    </tr>
</table>
</form>
       <script language="JavaScript">
  generate_wysiwyg('aciklama');
</script> 
<%ElseIf files="kaydet2" Then
id=secur(Request.Querystring("id"))
dim Dosya2,dosyaismi,aciklama,adres
Set dosya2 = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from download where id='"&id&"'"
dosya2.Open SQL,conne,1,3

dosyaismi=Request.Form("dosyaismi")
aciklama=Request.Form("aciklama")
adres=Request.Form("adres")
dosya2("dosyaismi")=dosyaismi
dosya2("aciklama")=aciklama
dosya2("adres")=adres
dosya2("tarih")= date()
dosya2.Update
Response.Redirect "default.asp?w8=filecontrol"

ElseIf files="kaydet" Then

dosyaismi=secur(Request.Form("dosyaismi"))
aciklama=secur(Request.Form("aciklama"))
adres=Request.Form("adres")
dosya.addnew
dosya("dosyaismi")=dosyaismi
dosya("aciklama")=aciklama
dosya("adres")=adres
dosya("tarih")= date()
dosya.Update
Response.Redirect "default.asp?w8=filecontrol"
ElseIf files="sil" Then
id=secur(Request.Querystring("id"))
Conne.Execute("delete download where id='"&id&"'")
Response.Redirect "default.asp?w8=filecontrol"
End If

Else
Response.Redirect "default.asp"
End If

Case "kod"

If Session("durum")="esp" Then

	Response.Buffer=True
	Action=Request("Action")
	SqlQuery=Request("SqlQuery")

 %>
 <div  class="style6 style8"><img src="../imgs/query_analyzer.jpeg" align="absmiddle"> SQL Query Analyzer</div>
<!-- #INCLUDE FILE="Library/WebGrid.asp" -->
		<script language="JavaScript" src="Library/WebGrid.js"></script>
		<link href="Style/Style.css" type="text/css" rel="stylesheet">
		<link href="Style/WebGrid/Classic/Grid.css" type="text/css" rel="stylesheet">
<table cellspacing="0" cellpadding="0" height="100%" width="100%" border="0" align="center">			
	<tr>				
		<td>					
			<table class="cssTableOutset" cellSpacing="3" cellPadding="0" width="100%" border="0">						
				<tr>							
					<td width="100%" style="font-size:12px;"><b>New Query</b></td>								
				</tr>					
			</table>				
		</td>			
	</tr>
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="100%">
			<table class="cssTableOutset" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">					
						<table width="100%" height="100%" cellSpacing="3" cellPadding="3" border="0">	
							<form name="frmSource" method="post" action="default.asp?w8=kod&Action=Exec">
							<tr>
								<td width="100%" height="90%" align="center">
									<textarea wrap="off" name="SqlQuery" id="SqlQuery" cols="110" rows="20"><%=SqlQuery%></textarea>
								</td>
							</tr>							
							</form>
						</table>   					
					</td>						
				</tr>	
			</table>
		</td>						
	</tr>	
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="30" valign="middle">
			<table class="cssTableOutset" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="middle" align="left">					
						<table cellSpacing="1" cellPadding="0" border="0">
							<tr>
								<td vAlign="top"><button id="cmdOK" onclick="document.frmSource.submit();" style="width:75px;height:25px;">Execute</button></td>
							</tr>
						</table> 
					</td>						
				</tr>	
			</table>
		</td>						
	</tr>
<%
If Action="Exec" Then
	'On Error Resume Next 
	Set Rs=Conne.Execute(SqlQuery)
	If Err<>0 Then 
		ErrMsg=Replace(Err.Description,"'","\'")
		Response.Write("<script>alert('" & ErrMsg & "');</script>")
	ElseIf Rs.State=1 Then 
%>
 	<tr>				
		<td height="300">
			<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">						
						<table width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td width="100%" height="100%" align="center">
<%
		Response.Write(WebGrid("QueryResults", Rs, 25, "","",""))
		Rs.close		
%>
								</td>						
							</tr>					
						</table>
					</td>						
				</tr>					
			</table>
		</td>						
	</tr>	
<%
	Else 
		Response.Write("<script>alert('Your Query Successfuly Executed...')</script>") 
	End If 
	conne.close 
End If 
%>
</table>
<% 
Else
Response.Redirect("default.asp")
End If

Case "kodgir" 
If Session("durum")="esp" Then
Dim qatxt,kodgir
qatxt=Request.Form("qatxt")
If Not trim(qatxt)="" Then
Set kodgir=Conne.Execute(qatxt)
End If
Response.Write("Query Analyzer kodu çalıştırıldı.")
Else
Response.Redirect("default.asp")
End If

Case "produces"
If Session("durum")="esp" Then
%><table width="200">
  <tr>
    <td><a href="default.asp?w8=produces&pro=freepot" class="produces">Sınırsız Hp Mp Pot</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=allcharsmaradon" class="produces">Bütün Charları Maradona At</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=weight0" class="produces">Bütün İtemlerin Ağırlığı 0</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=21gb" class="produces">Herkese 21 GB Ver</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=75fix" class="produces">Okçu 75 Skill (32 K Bug) Fix</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=deletebow" class="produces">Bowları Kaldır</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=clanadd" class="produces">Useri Clana koy</a></td>
  </tr>
  
  <tr>
    <td><a href="default.asp?w8=produces&pro=npsifirla" class="produces">Npleri Sıfırla</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=kralyap" class="produces">Useri Kral Yap</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=dropsifirla" class="produces">Drop Sıfırla</a></td>
  </tr>
  <tr>
    <td><a href="default.asp?w8=produces&pro=deleteacidholy" class="produces">Acid Potion ve Holy Waterleri Kaldır</a></td>
  </tr>
</table>
<br />
<br />
<%
pro=Request.Querystring("pro")

select Case pro
Case "freepot"
Set freepot=Conne.Execute("Update magic Set useitem='0' where useitem='389014000'"&"Update magic Set useitem='0' where useitem='389082000'")
Response.Write "Potlar sınırsız yapıldı."
Case "allcharsmaradon"
Set maradon=Conne.Execute("Update Userdata Set zone = '21'")
Response.Write "Bütün Charlar Maradona Atıldı. "
Case "weight0"
Set item0=Conne.Execute("Update ITEM Set Weight = '0'")
Response.Write "İtemlerin Ağırlıkları Kaldırıldı.."
Case "21gb"
Set gb21=Conne.Execute("Update userdata Set gold = '2100000000'")
Response.Write "Herkese 21 GB Verildi."
Case "75fix"
Set fix75=Conne.Execute("Update MAGIC_TYPE2 Set AddDamagePlus = 1 where NeedArrow = 1 or NeedArrow = 3 or NeedArrow = 5 ")
Response.Write "75 Skill Fixed !"
Case "deletebow"
Set deletebow=Conne.Execute("delete from item where strname like '%bow%'")
Response.Write "Bowlar Silindi."
Case "deleteacidholy"
Set acidholy=Conne.Execute("UPDATE ITEM SET SellingGroup = 0 Where Num = 389083000"&"Update ITEM Set ReqTitle = '100' , ReqLevel = '81', ReqRank = '1' where Num = '379101000'"&"Update ITEM Set ReqTitle = '100' , ReqLevel = '81', ReqRank = '1' where Num = '379102000'"&"Update ITEM Set ReqTitle = '100' , ReqLevel = '81', ReqRank = '1' where Num = '379103000'")
Case "npsifirla"
Set np=Conne.Execute("UPDATE USERDATA SET Loyalty = 0 ")
Case "dropsifirla"
Set drop=Conne.Execute("Update k_monster_item Set iItem01 = '0'"&"Update k_monster_item Set iItem02 = '0'"&"Update k_monster_item Set iItem03 = '0'"&"Update k_monster_item Set iItem04 = '0'"&"Update k_monster_item Set iItem05 = '0'")
Case "kralyap"%>
<form action="default.asp?w8=produces&pro=kralyapok" method="post">
Username : <input type="text" name="kralnick"/><br />
<input type="submit" value="Gönder" />
</form>
<br />
<% Case "kralyapok"
kralnick=Request.Form("kralnick")
Set nations=Conne.Execute("select nation from USERDATA where strUserId = '"&kralnick&"'")
Set clanid=Conne.Execute("select sIDnum from KNIGHTS_USER where StrUserID = '"&kralnick&"'")
Set clanname=Conne.Execute("select IDName from KNIGHTS where IDNum = "&clanid("sIDnum")&"")
Set accid=Conne.Execute("select strAccountID from ACCOUNT_CHAR where '"&kralnick&"' = strCharID1 or '"&kralnick&"' = strCharID2 or '"&kralnick&"' = strCharID3")
vt("Update KING_SYSTEM Set strKingName = '"&kralnick&"' where byNation = "&nations("nation")&"")
vt("Update USERDATA Set Rank = 0 where nation="&nations("nation"))
vt("Update USERDATA Set Rank = 1 where strUserId = '"&kralnick&"'")


Case "clanadd"%>
<form action="default.asp?w8=produces&pro=clanadd2" method="post" >
Kullanıcı adı : <input type="text" name="charid"/><br />
Clan ın Numarası : <input type="text" name="clanno"/><br />
<input type="submit" value="gönder" />
</form>
<%
Case "clanadd2"
charid=Request.Form("charid")
clanno=Request.Form("clanno")
Set clanadd=Conne.Execute("Update Userdata Set Knights = '"&clanno&"', Fame = 5 where struserid = '"&charid&"' ")
Set clanadd2=Conne.Execute("insert into knights_user values ('"&clanno&"','"&charid&"')")
Set clanadd3=Conne.Execute("Update knights Set members = members+1 where IDnum = '"&clanno&"' ")
Response.Write (charid&",&nbsp;"&clanno&"&nbsp; Nolu Clana Alınmıştır.")

Case Else
end select
 End If
Case "version"
If Session("durum")="esp" Then
Set version = Conne.Execute("select * from VERSION")
islem=Request.Querystring("islem")
If islem="" Then %>
<a href="default.asp?w8=version&islem=ekle">Yeni Patch Ekle</a>
<br><br><table width="200" border="1">
  <tr>
    <td>Version</td>
    <td>Filename</td>
    <td>Compname</td>
    <td>Hisversion</td>
  </tr>
  <% do while not version.eof %>
  <tr>
    <td><%=version(0)%></td>
    <td><%=version(1)%></td>
    <td><%=version(2)%></td>
    <td><%=version(3)%></td>
    <td><%Response.Write "<a href='default.asp?w8=version&islem=sil&id="&version(0)&"'>Sil</a>"%></td>
  </tr>
<%version.movenext
Loop%>
</table>

<%ElseIf islem="ekle" Then%>
<table width="200" border="1">
  <tr>
    <td>Version</td>
    <td>Filename</td>
    <td>Compname</td>
    <td>Hisversion</td>
  </tr><form action="default.asp?w8=version&islem=kaydet" method="post">
  <tr>
    <td><input type="text" name="1"></td>
    <td><input type="text" name="2"></td>
    <td><input type="text" name="3"></td>
    <td><input type="text" name="4"></td>
    <td><input type="submit" value="Ekle"></td>
  </tr></form>
</table>
<%ElseIf islem="kaydet" Then
v1=Request.Form("1")
v2=Request.Form("2")
v3=Request.Form("3")
v4=Request.Form("4")
Conne.Execute("insert into version values('"&v1&"','"&v2&"','"&v3&"','"&v4&"')")
Response.Redirect("default.asp?w8=version")
ElseIf islem="sil" Then
id=Request.Querystring("id")
Conne.Execute("delete version where version='"&id&"'")
Response.Redirect("default.asp?w8=version")
End If
Else
Response.Redirect("default.asp")
End If
Case "npc"
If Session("durum")="esp" Then
Set npcler=Conne.Execute("select * from k_npc order by strname")
%>
<center><h3>Npc Editor</h3></center>
<script>
function lnpc(url) {
window.open('npc.asp?zoneid='+url, 'Window2', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width='+screen.availWidth+10+',height='+screen.availHeight+',top=0,left=0')
}
</script>

<%If Request.Querystring("is")="duzenle" Then
npcid=trim(Request.Form("npcid"))
If not npcid="" Then
Set knpc=Conne.Execute("select * from k_npc where ssid="&npcid&"")%>
<form action="default.asp?w8=npc&is=kaydet" method="post" >
<input type="hidden" value="<%=knpc("ssid")%>" name="ssid">
<table width="350" border="0">
  <tr>
    <td width="91"><b>Npc Adı:</b></td>
    <td width="150"><input type="text" value="<%=trim(knpc("strname"))%>" style="width:200px" name="strname"></td>
    </tr>
  <tr>
    <td width="91">Exp:</td>
    <td width="150"><input type="text" value="<%=knpc("iexp")%>" name="iexp"></td>
  </tr>
  <tr>
    <td>Np:</td>
    <td><input type="text" value="<%=knpc("iloyalty")%>" name="iloyalty"></td>
  </tr>
  <tr>
    <td>Hp:</td>
    <td><input type="text" value="<%=knpc("ihppoint")%>" name="ihppoint"></td>
  </tr>
  <tr>
    <td>Mp:</td>
    <td><input type="text" value="<%=knpc("smppoint")%>" name="smppoint"></td>
  </tr>
  <tr>
    <td>Para:</td>
    <td><input type="text" value="<%=knpc("imoney")%>" name="imoney"></td>
  </tr>
  <tr>
    <td>Sağ El Silah No:</td>
    <td><input type="text" value="<%=knpc("iweapon1")%>" name="iweapon1"></td>
  </tr>
  <tr>
    <td>Sol El Silah No:</td>
    <td><input type="text" value="<%=knpc("iweapon2")%>" name="iweapon2"></td>
  </tr>
    <tr>
    <td>Defans:</td>
    <td><input type="text" value="<%=knpc("sac")%>" name="sac"></td>
  </tr>
  <tr>
  <td colspan="2" align="center"><input type="submit" value="Kaydet">&nbsp;<a href="default.asp?w8=npc&is=sil&id=<%=npcid%>" ><font class="style5">Npcyi sil</font></a></td>
  </tr>
</table>
</form>
<%End If
ElseIf Request.Querystring("is")="kaydet" Then
Conne.Execute("Update k_npc Set strname='"&Request.Form("strname")&"', iexp="&Request.Form("iexp")&",iloyalty="&Request.Form("iloyalty")&",ihppoint="&Request.Form("ihppoint")&",smppoint="&Request.Form("smppoint")&",imoney="&Request.Form("imoney")&",iweapon1="&Request.Form("iweapon1")&",iweapon2="&Request.Form("iweapon2")&",sac="&Request.Form("sac")&" where ssid="&Request.Form("ssid")&"")
Response.Redirect("default.asp?w8=npc")
End If
Else
response.rediret("default.asp")
End If
Case "upgrade"
If Session("durum")="esp" Then%>
<table width="600" border="0">
	<tr>
    <td colspan="2" align="center">Oranını değiştirmek istemediğinizin kutucuğunu boş bırakın.<br>Oranı Ayarlarken 100% için 10000 50% için 5000 %75 için 7500 gibi oranlar girmelisiniz.<br>&nbsp;</td>
    </tr>
  <tr>
    <td width="323"><form method="post" action="default.asp?w8=upgrade1">
<table width="281" border="1">
  <tr>
    <td colspan="2" align="center" class="style5">Silah Ve Armor Upgrade</td>
    </tr>
  <tr>
    <td width="48" align="center">+1</td>
    <td width="217">
      <input type="text" name="up1" id="up1">
    </td>
  </tr>
  <tr>
    <td align="center">+2</td>
    <td><input type="text" name="up2" id="up2"></td>
  </tr>
  <tr>
    <td align="center">+3</td>
    <td><input type="text" name="up3" id="up3"></td>
  </tr>
  <tr>
    <td align="center">+4</td>
    <td><input type="text" name="up4" id="up4"></td>
  </tr>
  <tr>
    <td align="center">+5</td>
    <td><input type="text" name="up5" id="up5"></td>
  </tr>
  <tr>
    <td align="center">+6</td>
    <td><input type="text" name="up6" id="up6"></td>
  </tr>
  <tr>
    <td align="center">+7</td>
    <td><input type="text" name="up7" id="up7"></td>
  </tr>
  <tr>
    <td align="center">+8</td>
    <td><input type="text" name="up8" id="up8"></td>
  </tr>
  <tr>
    <td align="center">+9</td>
    <td><input type="text" name="up9" id="up9"></td>
  </tr>
  <tr>
    <td align="center">+10</td>
    <td><input type="text" name="up10" id="up10"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="KAYDET"></td>
    </tr>
</table></form></td>
    <td width="267" valign="top">
    <form method="post" action="default.asp?w8=upgrade2">
    <table width="200" border="1">
  <tr>
    <td colspan="2" align="center" class="style5">Takı Upgrade</td>
    </tr>
  <tr>
    <td width="43" align="center">+1</td>
    <td width="141">
      <input type="text" name="tup1" id="tup1">
   </td>
  </tr>
  <tr>
    <td align="center">+2</td>
    <td><input type="text" name="tup2" id="tup2"></td>
  </tr>
  <tr>
    <td align="center">+3</td>
    <td><input type="text" name="tup3" id="tup3"></td>
  </tr>
  <tr>
    <td align="center">+4</td>
    <td><input type="text" name="tup4" id="tup4"></td>
  </tr>
  <tr>
    <td align="center">+5</td>
    <td><input type="text" name="tup5" id="tup5"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="KAYDET"></td>
    </tr>
</table> </form>
</td>
  </tr>
</table>




<%Else
Response.Redirect "default.asp"
End If
Case "upgrade1"
If Session("durum")="esp" Then
up1=trim(Request.Form("up1"))
up2=trim(Request.Form("up2"))
up3=trim(Request.Form("up3"))
up4=trim(Request.Form("up4"))
up5=trim(Request.Form("up5"))
up6=trim(Request.Form("up6"))
up7=trim(Request.Form("up7"))
up8=trim(Request.Form("up8"))
up9=trim(Request.Form("up9"))
up10=trim(Request.Form("up10"))

If not up1="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up1&" WHERE nOriginItem LIKE '%1'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up1&" WHERE nOriginItem LIKE '%1'AND nReqItem2 = 379021000")
End If
If not up2="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up2&" WHERE nOriginItem LIKE '%2'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up2&" WHERE nOriginItem LIKE '%2'AND nReqItem2 = 379021000")
End If
If not up3="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up3&" WHERE nOriginItem LIKE '%3'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up3&" WHERE nOriginItem LIKE '%3'AND nReqItem2 = 379021000")
End If
If not up4="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up4&" WHERE nOriginItem LIKE '%4'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up4&" WHERE nOriginItem LIKE '%4'AND nReqItem2 = 379021000")
End If
If not up5="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up5&" WHERE nOriginItem LIKE '%5'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up5&" WHERE nOriginItem LIKE '%5'AND nReqItem2 = 379021000")
End If
If not up6="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up6&" WHERE nOriginItem LIKE '%6'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up6&" WHERE nOriginItem LIKE '%6'AND nReqItem2 = 379021000")
End If
If not up7="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up7&" WHERE nOriginItem LIKE '%7'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up7&" WHERE nOriginItem LIKE '%7'AND nReqItem2 = 379021000")
End If
If not up8="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up8&" WHERE nOriginItem LIKE '%8'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up8&" WHERE nOriginItem LIKE '%8'AND nReqItem2 = 379021000")
End If
If not up9="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up9&" WHERE nOriginItem LIKE '%9'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up9&" WHERE nOriginItem LIKE '%9'AND nReqItem2 = 379021000")
End If
If not up10="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up10&" WHERE nOriginItem LIKE '%0'AND nReqItem1 = 379021000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up10&" WHERE nOriginItem LIKE '%0'AND nReqItem2 = 379021000")
End If
Response.Write "<br><center>Upgrade Oranı Başarıyla Ayarlandı..</center>"
Else
Response.Redirect "default.asp"
End If
Case "upgrade2"
If Session("durum")="esp" Then
up1=trim(Request.Form("tup1"))
up2=trim(Request.Form("tup2"))
up3=trim(Request.Form("tup3"))
up4=trim(Request.Form("tup4"))
up5=trim(Request.Form("tup5"))


If not up1="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up1&" WHERE nOriginItem LIKE '%1'AND nReqItem1 = 379159000")
End If
If not up2="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up2&" WHERE nOriginItem LIKE '%2'AND nReqItem1 = 379159000")
End If
If not up3="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up3&" WHERE nOriginItem LIKE '%3'AND nReqItem1 = 379159000")
End If
If not up4="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up4&" WHERE nOriginItem LIKE '%4'AND nReqItem1 = 379159000")
End If
If not up5="" Then
vt("UPDATE ITEM_UPGRADE SET nGenRate = "&up5&" WHERE nOriginItem LIKE '%5'AND nReqItem1 = 379159000")
End If

vt("UPDATE ITEM_UPGRADE SET nGenRate = 0 WHERE nOriginItem LIKE '%6'AND nReqItem1 = 379159000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = 0 WHERE nOriginItem LIKE '%7'AND nReqItem1 = 379159000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = 0 WHERE nOriginItem LIKE '%8'AND nReqItem1 = 379159000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = 0 WHERE nOriginItem LIKE '%9'AND nReqItem1 = 379159000")
vt("UPDATE ITEM_UPGRADE SET nGenRate = 0 WHERE nOriginItem LIKE '%0'AND nReqItem1 = 379159000")
Response.Write "<br><center>Upgrade Oranı Başarıyla Ayarlandı..</center>"
Else
Response.Redirect "default.asp"
End If
Case "level"
If Session("durum")="esp" Then
Set level=Conne.Execute("select * from level_up")%>
<div id="textListesi">
<p><a href="#" onClick="HepsiniSec('textListesi')">All 1 Exp</a>
| <a href="#" onClick="Temizle('textListesi')">Temizle</a></p>
<%Response.Write "<form action='default.asp?w8=level2' method='post'>"
do while not level.eof%>
<%=level("level")%> : &nbsp;<input type=text value=<%=level("exp")%> name='<%=level("level")%>' id='text_<%=level("level")%>'><br>
<%level.movenext
Loop
Response.Write "<br><input type='submit' value='GÜNCELLE' style=""height:40px;width:150px""></form></div>"
End If
Case "level2"
If Session("durum")="esp" Then
For Each ix In Request.Form
Set levelUpdate=Conne.Execute("Update level_up Set exp='"&secur(Request.Form(ix))&"' where level='"&ix&"'")

next
Response.Redirect "default.asp?w8=level"
End If
Case "00"
If Session("durum")="esp" Then
Dim MenuId
For Each MenuId in Request.Form
Conne.Execute("Update MenuAyar Set PSt="&Request.Form(MenuId)&" Where PId='"&MenuId&"'")
Next
Response.Redirect("default.asp?w8=0")
End If
Case "0"
If Session("durum")="esp" Then
dim menuayar
Set MenuAyar=Conne.Execute("Select * from Menuayar")
Response.Write("<form action=""default.asp?w8=00"" method=""post""><table width=""400"" border=""1"" cellspacing=""0"">")
Do While Not MenuAyar.Eof%>
  <tr>
    <td width="180"><%=MenuAyar("PName")%></td>
    <td width="100"><select name="<%=MenuAyar("PId")%>">
    <option value="1">Açık</option>
    <option value="0" <%If menuayar("PSt")="0" Then Response.Write"Selected"%>>Kapalı</option>
    </select></td>
  </tr>
<%MenuAyar.MoveNext
Loop
MenuAyar.Close
Set MenuAyar=Nothing
Response.Write("<tr><td height=""31"" colspan=""2"" align=""center""><input type=""submit"" class=""inputstyle"" value=""Güncelle"" style=""width:350;height:50px;font-size:11px""></td> </tr></table></form>")
Else
End If

Case "menusettings"
If Session("durum")="esp" Then
Dim islem,menu
islem=Request.Querystring("islem")
If islem="" Then
Set menu=Conne.Execute("select * from menu order by id asc")
%>
<script src="../js/jquery_002.js" language="javascript" type="text/javascript"></script>
<script src="../js/jqueryTableDnDArticle.js" language="javascript" type="text/javascript"></script>
<script src="_inc/jquery.js" language="javascript" type="text/javascript"></script>
<style type="text/css">
.tableDemo {
	background-color: white;
	border: 1px solid #666699;
	margin-right: 10px;
	padding: 6px;
}

.tableDemo table {
	border: 1px solid silver;
}

.tableDemo td {
	padding: 2px 6px
}
td.showDragHandle {
	background-image: url(http://www.isocra.com/images/updown2.gif);
	background-repeat: no-repeat;
	background-position: center center;
	cursor: move;
}
.tDnD_whileDrag {
	background-color: #eee;
}
</style>
<br>
    <a href="default.asp?w8=menusettings&islem=ekle"><font class="style5">Menü Ekle</font></a><br><br>
<form action="default.asp?w8=menusettings&islem=kaydet" method="post">

<table width="571" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td width="10" align="center"> </td>
	<td width="138" align="center">Menü Adı</td>
	<td width="136" align="center">Adres</td>
	<td width="136" align="center">OnClick (Ajax için)</td>
	<td width="136" align="center">Durum</td>
</tr>
</table>
<table width="571" border="0" cellpadding="1" cellspacing="1" id="table-3">
<% If not menu.eof Then
do while not menu.eof %>
<tr id="<%=menu("menuid")%>">
<td align="center" class="draghandle">&nbsp;&nbsp;&nbsp; </td>
<td align="center"><input type="hidden" value="<%=menu("id")%>" name="id">
<input type="text" size="30" value="<%=menu("menuname")%>" name="menuname"></td>
<td align="center"><input type="text" size="30" value="<%=menu("url")%>" name="url"></td>
<td align="center"><input type="text" size="35" value="<%=menu("click")%>" name="click"></td>
<td align="center"><select name="durum">
<option value="1" <%If menu("durum")=1 Then Response.Write "selected"%>>Açık</option>
<option value="0" <%If menu("durum")=0 Then Response.Write "selected"%>>Kapalı</option>
</select></td>
<td width="138" align="center"><a href="default.asp?w8=menusettings&islem=sil&id=<%=menu("menuid")%>">Sil</a></td>
</tr>
<%menu.movenext
Loop
End If %>
</table>
<table width="571" border="0" cellpadding="0" cellspacing="0" >
<tr><td>&nbsp;</td></tr>
<tr><td colspan="5"><input type="submit" value="KAYDET" style="width:600px" class="inputstyle"></td></tr>
</table>
</form>
<form><input type="hidden" name="siralama" id="siralama"></form>
<%ElseIf islem="kaydet" Then
Dim menuup,strSQL,menuname,url,click,durum,ids,msira
Set menuup = Server.CreateObject("ADODB.Recordset")
strSQL ="SELECT * FROM menu"
menuup.open strSQL,Conne,1,3
menuname = Split(request.Form("menuname"),", ")
url = Split(request.Form("url"),", ")
click = Split(request.Form("click"),", ")
durum = Split(request.Form("durum"),", ")
Conne.Execute("truncate table menu")
msira=0
for each ids in menuname
menuup.addnew
menuup("id")=msira
menuup("menuid")=msira
menuup("menuname")=menuname(msira)
menuup("url")=url(msira)
menuup("click")=click(msira)
menuup("durum")=durum(msira)
menuup.Update
msira=msira+1
next
Response.Redirect("default.asp?w8=menusettings")
ElseIf islem="ekle" Then
Set toplammenu=Conne.Execute("select count(*) toplam from menu")%>
<form action="default.asp?w8=menusettings&islem=ekle2" method="post">
<table width="547">
	<tr>
    <td width="27" align="center">Sıra</td>
	<td width="131" align="center">Menü Adı</td>
    <td width="132" align="center">Adres</td>
    <td width="141" align="center">OnClick ( Ajax içindir!  )</td>
	</tr>
<tr>
	<td><input name="sira" type="text" size="3" value="<%=toplammenu("toplam")%>"></td>
    <td align="center"><input type="text" name="menuisim"></td>
    <td align="center"><input type="text" name="url" size="35"></td>
    <td align="center"><input type="text" name="onclick" size="45"></td>
</tr>
<tr>
<td colspan="3" align="center"><input type="submit" class="inputstyle" value="Ekle" ></td></tr>
</table>
</form>
<% ElseIf islem="ekle2" Then
sira=request.Form("sira")
menuisim=request.Form("menuisim")
url=request.Form("url")
onclick=request.Form("onclick")
Set menuekle = Server.CreateObject("ADODB.Recordset")
sql = "Select * From menu"
menuekle.open sql,conne,1,3
menuekle.addnew
menuekle("id")=sira
menuekle("menuid")=sira
menuekle("menuname")=menuisim
menuekle("url")=url
menuekle("click")=onclick
menuekle.Update
Response.Redirect("default.asp?w8=menusettings")
menuekle.close
Set menukekle=nothing
ElseIf islem="sil" Then
id=Request.Querystring("id")
Set menusil=Conne.Execute("delete menu where menuid='"&id&"'")
Response.Redirect ("default.asp?w8=menusettings")
End If
End If

Case "sitesettings" 

If Session("durum")="esp" Then 
Dim sitesql,notice
Set sitesettings = Server.CreateObject("ADODB.Recordset")
sitesql = "Select * From siteayar"
sitesettings.open sitesql,conne,1,3

notice=Request.Form("notice")
If not notice="" Then
application("notice")=notice&"|"&now()
application("noticeuser")=""
End If
sitesettings("sitebaslik") = Request.Form("sitebaslik")
sitesettings("banner") = Request.Form("banner")
sitesettings("bannercolor") = Request.Form("bannercolor")
sitesettings("SunucuAdi") =Request.Form("SunucuAdi")
sitesettings("IP") = Request.Form("IP")
sitesettings("Droprate") =  Request.Form("droprate")
sitesettings("Expt") = Request.Form("Expt")
sitesettings("kulsiralama") = Request.Form("ksira")
sitesettings("clansiralama") = Request.Form("csira")
sitesettings("war") = Request.Form("war")
sitesettings("wartime") = Request.Form("wartime")
sitesettings("radyo") = Request.Form("radyo")
sitesettings("radyokonum") = Request.Form("radyokonum")
sitesettings("gmail") = Request.Form("mail")
sitesettings("sifre") = Request.Form("gmailsifresi")
sitesettings("duyuru") = Request.Form("duyuru")
sitesettings("icerik") = Request.Form("icerik")
sitesettings("gmyetki") = trim(Request.Form("gmyetki"))
sitesettings("Effect") = Request.Form("effectstyle1")&","&Request.Form("effectspeed")&","&Request.Form("pageloading")
sitesettings.Update

Response.Redirect "default.asp?w8=siteayarlari"
End If
Case "siteayarlari"
If Session("durum")="esp" Then 
dim sitesettings
Set sitesettings = Conne.Execute("Select * From siteayar")
Dim Effect
Effect=Split(sitesettings("Effect"),",")
%>

<table width="750" height="200" border="0">
<form action="default.asp?w8=sitesettings" method="post">
  <tr>
    <td ><strong>NOTICE GEÇ : </strong></td>
    <td><input type="text" name="notice" size="50" > <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Sitede Üstten Geçen USKO Tarzı Notice (Oyun İçinden Atılan Notice Değildir.)</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
   <tr>
    <td><strong>Site başlığı: </strong></td>
    <td><input type="text" size="40" name="sitebaslik" value="<% =sitesettings("sitebaslik") %>" /></td>
   </tr>
   <tr>
    <td><strong>Banner Yazısı: </strong></td>
    <td ><input type="text" size="40" name="banner" value="<% =sitesettings("banner") %>" /></td>
   </tr>
   <tr>
    <td><strong>Banner Renk: (Önerilen: #F9EED8)</strong></td>
    <td><input type="text" size="40" name="bannercolor" value="<% =sitesettings("bannercolor") %>" /></td>
   </tr>
  <tr>
    <td ><strong>Kayan Duyuru : </strong></td>
    <td><input type="text" name="duyuru" size="50" value="<%=sitesettings("duyuru")%>"></td>
    </tr>

   <tr>
    <td><strong>Sunucu Adı: </strong></td>
    <td><input name="sunucuadi" type="text" value="<% =sitesettings("SunucuAdi") %>" size="30" /></td>
    </tr>
   <tr>
    <td><strong>IP : </strong></td>
    <td><input name="ip" type="text" value="<% =sitesettings("IP") %>" size="30" /></td>
    </tr>
		 <tr>
    <td><strong>Sayfa Açılırken Değişme Efekti : </strong></td>
    <td><select name="effectstyle1">
	<option value="fadeIn" <%If Effect(0)="fadeIn" Then Response.Write("Selected")%>>Netleşme Efekti (Önerilen)</option>
	<option value="show" <%If Effect(0)="show" Then Response.Write("Selected")%>>Kaydırma Efekti</option>
	<option value="0" <%If Effect(0)="0" Then Response.Write("Selected")%>>Efekt Yok (Sayfa Çok Hızlı Açılır)</option>
	</select> <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Menü Kısmındaki Linklere Tıklayınca Sayfa Açılmadan Önce Oluşan Efekt. <br><br>Sayfa Kapanırken Değişme Efekti Ile Aynı Yaparsanız Daha Güzel Effect Olabilir.(Soluklaşma Ile Netleşme Veya Kaydırma Efekti)<br><br> Sayfa Efektsiz Çok Hızlı Açılsın İsterseniz Efekt Yok Ayarını Seçiniz(Bilmiyorsanız Önerileni Seçiniz)</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
	   <tr>
    <td><strong>Sayfa Efekt Hızı : </strong></td>
    <td><select name="effectspeed">
	<option value="fast" <%If Effect(1)="fast" Then Response.Write("Selected")%>>Hızlı</option>
	<option value="normal" <%If Effect(1)="normal" Then Response.Write("Selected")%>>Normal</option>
	<option value="slow" <%If Effect(1)="slow" Then Response.Write("Selected")%>>Yavaş (Önerilen)</option>
	</select> <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Menü Kısmındaki Linklere Tıklayınca Efecktin Oluşma Hızı.(Sayfa Açılma Hızı)<br><br>Yavaş Yavarsanız Daha Belirgin Şekilde Görebilirsiniz. (Bilmiyorsanız Önerileni Seçiniz)</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
		<tr>
    <td><strong>Efekt Arası Sayfa Yükleniyor Mesajı: </strong></td>
    <td><select name="pageloading">
	<option value="1" <%If Effect(2)="1" Then Response.Write("Selected")%>>Açık (Önerilen)</option>
	<option value="0" <%If Effect(2)="0" Then Response.Write("Selected")%>>Kapalı</option>
	</select> <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Sayfa Değişiken Arada Çıkan Sayfa Yükleniyor Efekti (Bilmiyorsanız Önerileni Seçiniz)</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
   <tr>
    <td><strong>Exp Rate : </strong></td>
    <td><input type="text" name="expt" value="<% =sitesettings("Expt") %>" /></td>
    </tr>
       <tr>
    <td><strong>Drop Rate : </strong></td>
    <td><input type="text" name="droprate" value="<% =sitesettings("Droprate") %>" /></td>
    </tr>
    <tr>
    <td><strong>Kullanıcı sıralamasında çıkacak toplam user : </strong></td>
    <td><input type="text" name="ksira" value="<% =sitesettings("kulsiralama") %>" /></td>
    </tr>
    <tr>
    <td><strong>Clan sıralamasında çıkacak toplam clan: </strong></td>
    <td><input type="text" name="csira" value="<% =sitesettings("clansiralama") %>" /></td>
    </tr>
    <tr>
    <td><strong>Savaş Açık/Kapalı: </strong></td>
    <td><select name="war">
    <option value="on" <% If sitesettings("war")="on" Then
	Response.Write "selected"
	End If%>>On</option>
    <option value="off" <% If sitesettings("war")="off" Then
	Response.Write "selected"
	End If%>>Off</option>
    </option></select> <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Açık Yapılırsa Site Üzerine Savaş Açılmıştır Uyarısı Gelir.</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
    <tr>
    <td><strong>Savaş Bitiş Zamanı:<br> (Örn: <%=time%> ) </strong></td>
    <td><input type="text" name="wartime" value="<% =sitesettings("wartime") %>" /></td>
    </tr>
	    <tr>
    <td><strong>Müzik Çalar(Radyo)</strong></td>
    <td><select name="radyo">
    <option value="1" <% If sitesettings("radyo")="1" Then
	Response.Write "selected"
	End If%>>Açık</option>
    <option value="0" <% If sitesettings("radyo")="0" Then
	Response.Write "selected"
	End If%>>Kapalı</option>
    </option></select></td>
    </tr>
	    <tr>
    <td><strong>Müzik Çalar(Radyo) Konumu</strong></td>
    <td><select name="radyokonum">
    <option value="ust" <% If sitesettings("radyokonum")="ust" Then
	Response.Write "selected"
	End If%>>Sayfanın Üstünde</option>
    <option value="alt" <% If sitesettings("radyokonum")="alt" Then
	Response.Write "selected"
	End If%>>Sayfanın Altında</option>
    </option></select></td>
    </tr>
  <tr>
    <td><strong>Gmail Adres : </strong></td>
    <td><input name="mail" type="text" value="<% =sitesettings("gmail") %>" size="40" /> <span onMouseOver="return overlib('<font style=\'color:#ffffff;font-weight:bold\'>Şifremi Unuttum Kısmının Çalışması Içın Gereklidir. Gmail Hesabınız Yoksa Alınız.</font>', RIGHT, WIDTH, 200,CELLPAD, 5, 5, 5)" onMouseOut="return nd();">Bu Nedir ?</span></td>
    </tr>
  <tr>
    <td><strong>Mail Şifre : </strong></td>
    <td><input name="gmailsifresi" type="text" value="<% =sitesettings("sifre") %>" /></td>
    </tr>
  <tr>
    <td><strong>Gmlerin Web Site Üzerinden Kullanabileceği Komutlar<br><br>Eklemek istediğiniz komutları bir alt satıra geçerek ekleyebilirsiniz. </strong></td>
    <td width="500"><%dim yetki,x
	yetki=Split(sitesettings("gmyetki"),",")
	Response.Write "<textarea name=""gmyetki"" id="""&yetki(x)&""" rows=""25"" cols=""30"">"
	For x=0 To UBound(yetki)
	If x=ubound(yetki) Then
	Response.Write trim(yetki(x))
	Else
	Response.Write trim(yetki(x))&vbcrlf
	End If
	Next
	Response.Write "</textarea>"%></td>
    </tr>
  <tr>
    <td height="95" colspan="2" valign="top" ><strong>Anasayfa İçerik</strong><br><br>
      <textarea cols="150" rows="15" name="icerik" id="icerik"><%=sitesettings("icerik")%></textarea>
       <script language="JavaScript">
  generate_wysiwyg('icerik');
</script>  </td>
    </tr>
  <tr>
    <td colspan="2"><input type="submit" class="inputstyle" style=" width:500px; font-size:13px; font-weight:bold" value="Güncelle"></td>
    </tr>
    </form>
</table>

<% 
End If
Case "1" 
If Session("durum")="esp" Then
Set users=Conne.Execute("select * from userdata")%>
<script language="javascript">


function userara(){
$.ajax({
   type: 'post',
   url: 'userbul.asp',
   data: 'struserid='+encodeURI( document.getElementById("struserid").value ),
   success: function(ajaxCevap) {
      $('#chrs').html(ajaxCevap);
   }
});
$('#chrs').fadeIn(0);
}

function userbul(userid){
document.getElementById('struserid').value=userid
}

function kontrolet(olay){
olay = olay || event;
if (olay.keyCode==13){
icerikal();
$('#chrs').fadeOut(0);
}
else{
userara();
$('#chrs').fadeIn(0);
}
}


</script>
<form action="default.asp?w8=1xsearch" method="POST" name="userbul" id="userbul">
<font face="Verdana" style="font-size:9pt;"><b>Kullanıcı Adını Yazın: </b></font>
<input  type="text" name="struserid" id="struserid" size="30"><a href="#" onClick="userara();return false">Karakter Adı Bul</a>
<input type="submit" value="Ara"></form>
<div id="chrs" style="width:190px;height:100px;overflow:auto;position:relative;left:138px;top:-12px;background-color:lightgrey;padding-left:5px;padding-top:5px;padding-bottom:5px;border-left:2px;border-left-style: groove;border-right-style: groove; z-index:2"></div>

<% ElseIf Session("durum")="sup" Then
End If 

Case "1xsearch"
If Session("durum")="esp" Then
strUserID=Request.Form("strUserID")
Set statnp = Server.CreateObject("ADODB.Recordset")
statSQL = "Select * From [USERDATA] Where strUserID='"&strUserID&"'"
statnp.open statSQL,Conne,1,3

If not statnp.eof Then%>
<br />
<style>
.gn{
	width:150px;
	
	}
</style>
<form action="default.asp?w8=1xUpdate" method="POST">
<table width="295" border="0" cellpadding="0" align="left">
<tr width="120"><td><b>Kullanıcı Adı</b></td>
	<td width="175"><b><% =struserid %></b></td>
</tr>
  <tr><td><b>NP  </b></td>
<td><input type="text" name="Loyalty" value="<%=statnp("Loyalty")%>" class="gn"/></td>
</tr>
	<tr>
    <td><b>Aylık Np </b></td>
    <td><input type="text" name="LoyaltyMonthly" value="<%=statnp("LoyaltyMonthly")%>" class="gn"/></td>
    </tr>
 <tr>
    <td><b>Level  </b></td>
    <td><input type="text" name="level" value="<%=statnp("Level")%>" class="gn"></td>
    </tr>
 <tr>
    <td><b>Exp  </b></td>
    <td><input type="text" name="Exp" value="<%=statnp("Exp")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Para  </b></td>
    <td><input type="text" name="Gold" value="<%=statnp("Gold")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Clan No  </b></td>
    <td><input type="text" name="knights" value="<%=statnp("knights")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Fame  </b></td>
    <td><input type="text" name="fame" value="<%=statnp("fame")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Rank  </b></td>
    <td><input type="text" name="rank" value="<%=statnp("Rank")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Title  </b></td>
    <td><input type="text" name="title" value="<%=statnp("title")%>" class="gn"></td>
    </tr>

  <tr>
    <td><b>Race  </b></td>
    <td><input type="text" name="race" value="<%=statnp("race")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Class  </b></td>
    <td><input type="text" name="class" value="<%=statnp("class")%>" class="gn"> </td>
    </tr>
  <tr>
    <td><b>Face  </b></td>
    <td><input type="text" name="face" value="<%=statnp("face")%>" class="gn"> </td>
    </tr>
<tr>
    <td><b>Durum  </b></td>
    <td><select name="Authority" class="gn">
      <option value="0" <% If statnp("Authority")="0" Then%>selected<%Else%><%End If%>>GM</option>
      <option value="1" <% If statnp("Authority")="1" Then%>selected<%Else%><%End If%>>Normal Kullanıcı</option>
      <option value="11" <% If statnp("Authority")="2" or statnp("Authority")="11" Then%>selected<%Else%><%End If%>>Mute</option>
      <option value="255" <% If statnp("Authority")="255" Then%>selected<%Else%><%End If%>>Yasaklı Kullanıcı</option>
    </select></td>
    </tr>
    <tr>
    <td><b>Şehir  </b></td>
    <td>
	<select name="zone">
<% Set zonename=Conne.Execute("select bz,zoneno from zone_info ")
do while not zonename.eof
Response.Write "<option value="""&zonename("zoneno")&""""
If zonename("zoneno")=statnp("zone") Then
Response.Write "selected"
End If
Response.Write ">"&zonename("bz")&"</option>"
zonename.movenext
Loop %>
	</select></td>
    </tr>
  <tr>
    <td><b>Gm gün  </b></td>
    <td><input type="text" name="gmgun" value="<%=statnp("gm_gun")%>" class="gn"></td>
    </tr>
      <tr>
    <td><b>Mute Sayısı  </b></td>
    <td><input type="text" name="mutecount" value="<%=statnp("mutecount")%>" class="gn"></td>
    </tr>
  <tr>
    <td><b>Ban Sayısı  </b></td>
    <td><input type="text" name="bancount" value="<%=statnp("bancount")%>" class="gn"></td>
    </tr>
  <tr>
  <tr>
    <td><b>Yasak Süresi(Gün)</b></td>
    <td><input type="text" name="bangun" value="<%=statnp("yasakgun")%>" class="gn"></td>
    </tr>
      <tr>
    <td><b>Yasaklanma Sebebi  </b></td>
    <td><textarea name="bansebep" class="gn" ><%=statnp("yasaksebep")%></textarea></td>
    </tr>
  
</table>
<table width="350" border="0" cellpadding="0" style="position:relative;left:0px">
    <tr>
    <td width="50"><b>Str :</b></td>
    <td width="189"><input type="text" name="Strong" value="<%=statnp("Strong")%>" /></td>
    </tr>
  <tr>
    <td><b>HP:</b></td>
    <td><input type="text" name="Sta" value="<%=statnp("Sta")%>" /></td>
    </tr>
  <tr>
    <td><b>Dex:</b></td>
    <td><input type="text" name="Dex" value="<%=statnp("Dex")%>" /></td>
    </tr>
  <tr>
    <td><b>Intel:</b></td>
    <td><input type="text" name="Intel" value="<%=statnp("Intel")%>" /></td>
    </tr>
  <tr>
    <td><b>MP:</b></td>
    <td><input type="text" name="Cha" value="<%=statnp("Cha")%>" /></td>
    </tr>
  <tr>
    <td><b>Points:</b></td>
    <td><input type="text" name="Points" value="<%=statnp("Points")%>" /></td>
    </tr>

<%Function BinaryToString(Binary)
  Dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    If cl3>300 Then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      If cl2>200 Then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  BinaryToString = pl1 & pl2 & pl3
End Function


skill1=(BinaryToString(midb(statnp("strskill"),11,2)))
If skill1<>"" Then
skill1=asc(skill1)
Else
skill1=0
End If
skill2=(BinaryToString(midb(statnp("strskill"),13,2)))
If skill2<>"" Then
skill2=asc(skill2)
Else
skill2=0
End If
skill3=(BinaryToString(midb(statnp("strskill"),15,2)))
If skill3<>"" Then
skill3=asc(skill3)
Else
skill3=0
End If
skill4=(BinaryToString(midb(statnp("strskill"),17,2)))
If skill4<>"" Then
skill4=asc(skill4)
Else
skill4=0
End If
point=(BinaryToString(midb(statnp("strskill"),1,2)))
If point<>"" Then
point=asc(point)
Else
point=0
End If

Function cla3(tur)

select Case tur
Case "101", "105", "106", "201", "205", "206"
cla3="Warrior"
Case "102", "107", "108", "202", "207", "208"
cla3="Rogue"
Case "103", "109", "110", "203", "209", "210"
cla3= "Mage"
Case "104", "111", "112", "204", "211", "212"
cla3= "Priest"
Case Else
cla3= "Unknown"
end select

end Function

clas=cla3(statnp("class"))

If clas="Warrior" Then
skillname1="Attack"
skillname2="Defense"
skillname3="Berserker"
ElseIf clas = "Mage" Then
skillname1 = "Flame"
skillname2 = "Glacier"
skillname3 = "Lightning"
ElseIf clas = "Priest" Then
skillname1 = "Heal"
skillname2 = "Buff"
skillname3 = "Debuff"
ElseIf clas = "Rogue" Then
skillname1 = "Archer"
skillname2 = "Assassin"
skillname3 = "Explore"
End If
%>
<tr><td colspan="2"><br>
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Skill Point</b></td></tr>
<tr><td><b><%=skillname1%></b></td><td> <input type="text" value="<%=skill1%>" name="1"></td></tr>
<tr><td><b><%=skillname2%></b></td><td> <input type="text" value="<%=skill2%>" name="2"><br></td></tr>
<tr><td><b><%=skillname3%></b></td><td> <input type="text" value="<%=skill3%>" name="3"><br></td></tr>
<tr><td><b>Master</b></td><td> <input type="text" value="<%=skill4%>" name="4"><br></td></tr>
<tr><td><b>Point:</td><td> <input type="text" value="<%=point%>" name="5"></td></tr>
  <tr>
    <td colspan="2"><input type="hidden" value="<%=strUserID%>" name="StrUserID" />
      <input name="submit" type="submit" value="Güncelle" style="width:200px"></td>
    </tr>
       <tr>
     <td colspan="2"><a href="default.asp?w8=1xsil&char=<%=trim(strUserID)%>"><font face="tahoma" size="2">Kullanıcıyı sil</font></a></td>
    </tr>
    
</table>
</form>


<br>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Böyle Bir Kullanıcı Mevcut Değil.</font>
<% End If 
ElseIf Session("durum")="sup" Then
End If 
Case "1xUpdate" 

If Session("durum")="esp" Then

Loyalty=Request.Form("Loyalty")
LoyaltyMonthly=Request.Form("LoyaltyMonthly")
Strong=Request.Form("Strong")
Sta=Request.Form("Sta")
Dex=Request.Form("Dex")
Intel=Request.Form("Intel")
Cha=Request.Form("Cha")
points=Request.Form("Points")
Level=Request.Form("Level")
Authority=Request.Form("Authority")
strUserID=Request.Form("strUserID")
Gold=Request.Form("Gold")
Zone=Request.Form("Zone")
fame=Request.Form("fame")
face=Request.Form("face")
rank=Request.Form("rank")
title=Request.Form("title")
race=Request.Form("race")
clas=Request.Form("class")
gmgun=Request.Form("gmgun")
mutecount=Request.Form("mutecount")
bancount=Request.Form("bancount")
bangun=Request.Form("bangun")
bansebep=Request.Form("bansebep")
skill1=Request.Form("1")
skill2=Request.Form("2")
skill3=Request.Form("3")
skill4=Request.Form("4")
skill5=Request.Form("5")
skills= Chr(skill5)&Chr(0)&Chr(0)&Chr(0)&Chr(0)&Chr(skill1)&Chr(skill2)&Chr(skill3)&Chr(skill4)&Chr(0)
If trim(bangun)="" Then
bangun=Null
End If
If trim(bansebep)="" Then
bansebep=Null
End If
If Authority<>255 and Authority<>11 Then
bangun=null
bansebep=null
End If
Set statnp = Server.CreateObject("ADODB.Recordset")
statSQL = "Select * From USERDATA Where strUserID='"&strUserID&"'"
statnp.open statSQL,Conne,1,3

If not statnp.eof Then

statnp("Loyalty")=Loyalty
statnp("LoyaltyMonthly")=LoyaltyMonthly
statnp("Strong")=Strong
statnp("Sta")=Sta
statnp("Dex")=Dex
statnp("Intel")=Intel
statnp("Cha")=Cha
statnp("Points")=Points
statnp("level")=level
statnp("Authority")=Authority
statnp("Gold")=Gold
statnp("Zone")=Zone
statnp("fame")=fame
statnp("face")=face
statnp("rank")=rank
statnp("title")=title
statnp("race")=race
statnp("class")=clas
statnp("gm_gun")=gmgun
statnp("mutecount")=mutecount
statnp("bancount")=bancount
statnp("yasakgun")=bangun
statnp("yasaksebep")=bansebep
statnp("strskill")=skills
statnp.Update%>
<meta http-equiv="refresh" content="1;url=default.asp?w8=1">
<font face="Verdana" style="font-size:9pt;">Güncelleme Başarılı.
<br />
Yönlendiriliyorsunuz.<br />
<img src="../imgs/18-1.gif" />
<%Else %>
<font face="Verdana" style="font-size:9pt;">Kullanıcı Bulunamadı.</font>
<% End If
ElseIf Session("durum")="sup" Then
End If 
Case "1xsil" 
If Session("durum")="esp" Then 

char=Request.Querystring("char")
Set sil =Conne.Execute("delete from userdata where strUserID='"&char&"'")
End If
Case "2"
If Session("durum")="esp" Then
%>
<script language="javascript">
function userara(){
$.ajax({
   type: 'post',
   url: 'userbul.asp',
   data: $('#userbul').serialize(),
   success: function(ajaxCevap) {
      $('#userler').html(ajaxCevap);
   }
});
}
function userbul(userid){
document.getElementById('struserid').value=userid
}
</script>
<form action="default.asp?w8=2search" method="POST" name="userbul" id="userbul">
<font face="Verdana" style="font-size:9pt;"><b>Kullanıcı Adını Yazın: </b></font>
<input  type="text" name="strUserID" id="struserid" onKeyUp="userara()" autocomplete="off">
<input type="submit" value="Ara"></form>
<div name="userler" id="userler"></div>
<br></form>

<%ElseIf Session("durum")="sup" Then
End If 
Case "2search"
If Session("durum")="esp" Then


dim strUserID,irk,irksql
 strUserID=Request.Form("strUserID")
 If strUserID = "" Then
Response.Redirect "default.asp?w8=2"
End If 

Set irk = Server.CreateObject("ADODB.Recordset")
irkSQL = "Select * From [USERDATA] Where strUserID='"&strUserID&"'"
irk.open irkSQL,Conne,1,3

If not irk.eof Then%>
<form action="default.asp?w8=2xsearchxnationsecildi" method="POST">
<font face="Verdana" style="font-size:9pt;"><b>Irk Seçimi : </b></font>
<select name="Nation"><option value="1" <%If irk("Nation")="1" Then%>selected<%End If%>>Karus</option>
<option value="2" <%If irk("Nation")="2" Then%>selected<%End If%>>Human</option></select><br>
<input type="hidden" value="<%=irk("strUserID")%>" name="strUserID">
<input type="submit" value="Giriş"></form>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Böyle Bir Kullanıcı Mevcut Değil.</font>
<% End If 
ElseIf Session("durum")="sup" Then
End If
Case "2xsearchxnationsecildi" 

If Session("durum")="esp" Then
strUserID=Request.Form("strUserID")
dim natn
natn=Request.Form("nation") 

Set irk = Server.CreateObject("ADODB.Recordset")
irkSQL = "Select * From [USERDATA] Where strUserID='"&strUserID&"'"
irk.open irkSQL,Conne,1,3

If natn="1" Then %>
<form action="default.asp?w8=2xkarusUpdate&strUserID=<%=strUserID%>" method="POST">
Race :<select name="Rac">
			<option value="1"<%If irk("Race")="1" Then%> selected<%Else%><%End If %>>Büyük Warrior</option>
			<option value="2"<%If irk("Race")="2" Then%> selected<%Else%><%End If %>>Erkek oyuncu</option>
			<option value="3"<%If irk("Race")="3" Then%> selected<%Else%><%End If %>>Cüce Mage</option>
			<option value="4"<%If irk("Race")="4" Then%> selected<%Else%><%End If %>>Kız oyuncu</option>
</select><br>Tür :&nbsp;&nbsp;
<select name="Cla">
			<option value="106" <%If irk("class")="106" Then%> selected<%Else%><%End If %>>Warrior</option>
			<option value="108" <%If irk("class")="108" Then%> selected<%Else%><%End If %>>Rogue</option>
			<option value="110" <%If irk("class")="110" Then%> selected<%Else%><%End If %>>Mage</option>
			<option value="112" <%If irk("class")="112" Then%> selected<%Else%><%End If %>>Priest</option>
</select>
<input type="hidden" value="1" name="Nation">
<input type="submit" value="Kaydet"></form>
<%ElseIf natn="2" Then %>
<form action="default.asp?w8=2xhumanUpdate&strUserID=<%=strUserID%>" method="POST">
Race : <select name="Rac">
			<option value="11"<%If irk("Race")="11" Then%> selected<%Else%><%End If %>>Barbarian Warrior</option>
			<option value="12"<%If irk("Race")="12" Then%> selected<%Else%><%End If %>>Erkek oyuncu</option>
			<option value="13"<%If irk("Race")="13" Then%> selected<%Else%><%End If %>>Kız oyuncu</option>
</select><br>Tür :&nbsp;&nbsp;
<select name="Cla">
 			<option value="206" <%If irk("Class")="206" Then%> selected<%Else%><%End If %>>Warrior</option>
			<option value="208" <%If irk("Class")="208" Then%> selected<%Else%><%End If %>>Rogue </option>
			<option value="210" <%If irk("Class")="210" Then%> selected<%Else%><%End If %>>Mage</option>
			<option value="212" <%If irk("Class")="212" Then%> selected<%Else%><%End If %>>Priest</option>
</select>
<input type="hidden" value="2" name="Nation">
<input type="hidden" value="<%=strUserID%>" name="strUserID">&nbsp;&nbsp;&nbsp;
<input type="submit" value="Kaydet"></form>
<%End If 
ElseIf Session("durum")="sup" Then
 End If 
 Case "2xkarusUpdate" 
 If Session("durum")="esp" Then
 Naton=Request.Form("Nation")
 Rac=Request.Form("Rac") 
 Cl=Request.Form("Cla") 
 strUserID=Request.Querystring("strUserID") 
 If Naton="1" Then 
Set karusUpdate = Server.CreateObject("ADODB.Recordset")
karusSQL = "Select * From [USERDATA] Where strUserID='"&strUserID&"'"
karusUpdate.open karusSQL,Conne,1,3
If not karusUpdate.eof Then 
karusUpdate("Nation")=Naton
karusUpdate("Race")=Rac
karusUpdate("Class")=Cl
karusUpdate.Update
%>
<font face="Verdana" style="font-size:9pt;">Güncelleme Başarılı.</font>
<a href="default.asp?w8=9">Eğer Nation Transfer Yaptıysanız 
Geri Kalan İşlemi Tamamlamak için Buraya Tıklayın.</a>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Kullanıcı Bulunamadı.</font>
<% End If 
ElseIf Nation="2" Then
End If

End If 
Case "2xhumanUpdate"
If Session("durum")="esp" Then
dim Naton,Rac,Cl,humanUpdate,humanSQL
Naton=Request.Form("Nation")
Rac=Request.Form("Rac") 
Cl=Request.Form("Cla") 
strUserID=Request.Querystring("strUserID")
 If Naton="2" Then 

Set humanUpdate = Server.CreateObject("ADODB.Recordset")
humanSQL = "Select * From [USERDATA] Where strUserID='"&strUserID&"'"
humanUpdate.open humanSQL,Conne,1,3

 If not humanUpdate.eof Then 

humanUpdate("Nation")=Naton
humanUpdate("Race")=Rac
humanUpdate("Class")=Cl
humanUpdate.Update
%>
<a href="default.asp?w8=9">Eğer Nation Transfer Yaptıysanız Geri Kalan İşlemi Tamamlamak için Buraya Tıklayın.</a>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Kullanıcı Bulunamadı.</font>
<% End If 
ElseIf Naton="1" Then
End If 
ElseIf Session("durum")="sup" Then
 End If 
 Case "res" 
If Session("durum")="esp" Then%>
<table width="450" height="148" border="0">
		<tr>
		<td width="709" align="center" class="text4">
		Server Reset		</td>
		</tr>
  <tr>
    <td><strong>1. <a href="default.asp?w8=res&res=onlinedel" >Online Userler tablosunu Temizle (Currentuser)</a></strong></td>
    </tr>
	  <tr>
    <td><strong>2. <a href="default.asp?w8=res&res=sembolres" >Kareli Sembol Güncelleme (Genel Np Res)</a></strong></td>
	 </tr>
	  <tr>
    <td><strong>3. <a href="default.asp?w8=res&res=aylikres">Karesiz Sembol Güncelleme (Aylık Np Res)</a></strong></td>
  </tr>
  <tr>
    <td><strong>4. <a href="default.asp?w8=res&res=clanpointUpdate" >Clan Pointleri Nplere Göre Güncelle</a><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Veya<br>
    </strong><strong><a href="default.asp?w8=res&res=clannpUpdate" >Clan Pointleri Np Bağışlarına Göre Güncelle </a></strong></td>
    </tr>
	 
 <tr>
    <td><br>
      Not: İlk 4 maddeyi sırasıyla yapın aksi takdirde sorunlar çıkabilir. </td>
  </tr>
<ul><tr>
    <td><br>
       <li><a href="default.asp?w8=res&res=allflag2">Bütün Clanları Üst Clan yap</a></li></td>
  </tr>
<tr>
    <td>
      <li><a href="default.asp?w8=res&res=allg1">Bütün Clanları Grade1 yap</a></li></td>
  </tr>
  <tr>
    <td><li><a href="default.asp?w8=res&res=allcape">Pelerinsiz Clanlara Pelerin Ver</a></li></td>
	  </tr>
	  <tr>
    <td><li><a href="default.asp?w8=res&res=allnpresetreset" >Npleri Sıfırla</a></li></td>
	</tr>
	  <tr>
	<tr>
    <td><li><a href="default.asp?w8=res&res=ayliknpreset" >Aylık Npleri Sıfırla</a></li></td>
	  </tr>
	  	  <tr>
          	<tr>
    <td><li><a href="default.asp?w8=res&res=weeklynpreset" >Haftalık Npleri Sıfırla</a></li></td>
	  </tr>
	  	  <tr>
    <td><li><a href="default.asp?w8=res&res=gunluknpreset" >Günlük Npleri Sıfırla</a></li></td>
	</tr>
	</ul>
</table><br><br>

<%Dim res
res=Request.Querystring ("res")
If res="onlinedel" Then
Conne.Execute("truncate table currentuser")
Response.Write "Online tablosu temizlendi"
ElseIf res="sembolres" Then
Conne.Execute("EXEC UPDATE_USER_KNIGHTS_RANK")
Response.Write "Kareli np sembolleri güncellendi."
ElseIf res="aylikres" Then
Conne.Execute("EXEC UPDATE_USER_PERSONAL_RANK")
Response.Write "Karesiz np sembolleri güncellendi."
ElseIf res="clanpointUpdate" Then
Conne.Execute("Exec UPDATE_KNIGHTS_RATING")
Response.Write "Clanlar güncellendi."
ElseIf res="clannpUpdate" Then
Set res=Conne.Execute("select idnum,points from knights")
If not res.eof Then
do while not res.eof
dim point
Set point=Conne.Execute("Select sum(np) as np From npdonate where clan='"&res("idnum")&"' ")
If not point("np")="" Then
Conne.Execute("Update knights Set points='"&point("np")&"' where idnum='"&res("idnum")&"'")
Else
Conne.Execute("Update knights Set points='0' where idnum='"&res("idnum")&"'")
End If
res.movenext
Loop
Response.Write "Clanlar bağışlanan nplere göre güncellendi."
End If
Conne.Execute("TRUNCATE TABLE KNIGHTS_RATING")
Conne.Execute("INSERT INTO KNIGHTS_RATING SELECT IDNum, IDName, Points FROM KNIGHTS ORDER BY Points DESC")
Conne.Execute("UPDATE KNIGHTS SET RANKING=0")
Conne.Execute("UPDATE KNIGHTS SET Ranking = (SELECT nRank FROM KNIGHTS_RATING WHERE shIndex = IDNum AND nRank <=5) WHERE (SELECT nRank FROM KNIGHTS_RATING WHERE shIndex = IDNum AND nRank <= 5) <= 5")
ElseIf res="allflag2" Then
Conne.Execute("Update knights Set flag=2")
ElseIf res="allg1" Then
Conne.Execute("Update KNIGHTS Set Points =720000 where points<720000")
Response.Write "Bütün clanlar Grade 1 yapıldı."
ElseIf res="allcape" Then
Conne.Execute("Update KNIGHTS Set Scape =1 where scape=-1")
Response.Write "Pelerinsiz Clanlara pelerin verildi."
ElseIf res="allnpreset" Then
Conne.Execute("Update userdata Set loyalty =0")
Response.Write "National Pointler sıfırlandı."
ElseIf res="ayliknpreset" Then
Conne.Execute("Update userdata Set monthlyloyalty =0")
Response.Write "Aylık Npler Sıfırlandı."
ElseIf res="weeklynpreset" Then
Conne.Execute("Update userdata Set HaftalikNp1=0, HaftalikNp2=0")
ElseIf res="gunluknpreset" Then
Conne.Execute("Update userdata Set gunluknp1=0,gunluknp2=0")
Response.Write "Günlük Npler Sıfırlandı."
End If
End If

Case "pus"
If Session("durum")="esp" Then%>
<div align="center"><a href="default.asp?w8=pusitemadd"><font class="style5"><img src="../imgs/shutterstock.gif" border="0" width="48">İtem Ekle</font></a></div><br>

<%Dim power
Set power=Conne.Execute("select * from pus_itemleri order by type asc,itemismi asc")
If not power.eof Then%>
<form action="default.asp?w8=puskyt" method="post">
<table border="1" align="left" cellpadding="0" cellspacing="0">
  <tr>
    
    <td align="center" class="style5">TYPE</td>
    <td align="center" class="style5">ITEM ISMI</td>
    <td align="center" class="style5">ÜCRET</td>
    <td align="center" class="style5">ADET</td>
    <td align="center" class="style5">ITEM KODU</td>
    <td align="center" class="style5">RESİM</td>
    <td align="center" class="style5">HIT</td>
    <td align="center" class="style5">Düzenle</td>
    <td align="center" class="style5">&nbsp;&nbsp;Sil&nbsp;&nbsp;</td>
  </tr>
<%do while not power.eof%>

  <tr>
    <td align="center"><%=power(1)%></td>
    <td align="center"><%=power(2)%></td>
    <td align="center"><%=power(3)%></td>
    <td align="center"><%=power(4)%></td>
    <td align="center"><%=power(5)%></td>
    <td align="center"><%=power(6)%></td>
    <td align="center"><%=power(7)%></td>
    <td align="center" valign="middle"><a href="default.asp?w8=pusduzenle&id=<%=power(0)%>"><img src="../imgs/icons/pencil.gif" border="0" title="Düzenle"></a></td>
    <td align="center" valign="middle"><a href="default.asp?w8=pussil&id=<%=power(0)%>"><img src="../imgs/icons/sil.gif" border="0" title="Sil"></a></td>
  </tr>
<%
power.movenext
Loop%>
</table>
</form>
<%Else
Response.Write "İtem bulunamadı"
End If
End If

Case "pusduzenle"
If Session("durum")="esp" Then
id=Request.Querystring("id")
If  isnumeric(id)=false Then
Response.End
End If
Dim pusduzen,pus,itemtit
Set pusduzen = Server.CreateObject("ADODB.Recordset")
pus = "Select * From pus_itemleri Where id='"&id&"'"
pusduzen.open pus,Conne,1,3

If Request.Querystring("guncelle")="1" Then

dim itemid,itype,itemismi,ucret,adet,itemkodu,resm,hit,detay,sure,preitem,title
itemid=Request.Form("id")
itype=Request.Form("type")
itemismi=Request.Form("itemismi")
ucret=Request.Form("ucret")
adet=Request.Form("adet")
itemkodu=Request.Form("itemkodu")
resm=Request.Form("resim")
hit=Request.Form("hit")
detay=Request.Form("detay")
sure=Request.Form("sure")
preitem=Request.Form("preitem")
title=Request.Form("title")

If title="" Then
title=0
End If


Conne.Execute("Update item Set reqtitle="&title&" where num="&itemkodu)


If itemid="" or itype="" or itemismi="" or ucret="" or itemkodu="" or adet="" or resm="" or preitem="" or sure="" Then
Response.Write("Boş Alan Bırakmayınız.")
Response.End
End If

pusduzen("id")=itemid
pusduzen("type")=itype
pusduzen("itemismi")=itemismi
pusduzen("ucret")=ucret
pusduzen("adet")=adet
pusduzen("itemkodu")=itemkodu
pusduzen("resim")=resm
pusduzen("alindi")=hit
pusduzen("detay")=detay
pusduzen("premiumitem")=preitem
pusduzen("kullanimgunu")=sure
pusduzen.Update
Response.Write "Bilgiler Güncellendi."
End If %>


<form action="default.asp?w8=pusduzenle&id=<%=id%>&guncelle=1" method="post">
<table border="0">
<tr><td><a href="default.asp?w8=pus">Geri Dön</td></tr>
<tr>
<td>ID :</td>
<td><input type="text" name="id" value="<%=pusduzen("id")%>"></td>
</tr>
<tr>
<td>Type :</td>
<td><select name="type" id="type">
  <option value="scroll" <% If pusduzen("type")="scroll" Then Response.Write "selected"%>>Scroll</option>
  <option value="kiyafet" <% If pusduzen("type")="kiyafet" Then Response.Write "selected"%>>Kıyafet</option>
  <option value="taki" <% If pusduzen("type")="taki" Then Response.Write "selected"%>>Takı</option>
  <option value="silah" <% If pusduzen("type")="silah" Then Response.Write "selected"%>>Silah</option>
  <option value="kalkan" <% If pusduzen("type")="kalkan" Then Response.Write "selected"%>>Kalkan</option>
  <option value="anasayfa" <% If pusduzen("type")="anasayfa" Then Response.Write "selected"%>>Anasayfa</option>
</select></td>
</tr>
<tr>
<td>Item Kodu :</td>
<td><input type="text" name="itemkodu" id="itemn" value="<%=pusduzen("itemkodu")%>" onKeyUp="javascript:rsm(this.value)" onChange="javascript:rsm(this.value)"></td>
</tr>
<tr>
<td>Item Ismi :</td>
<td><input type="text" name="itemismi" id="itemismi" value="<%=pusduzen("itemismi")%>"></td>
</tr>
<tr>
<td>Resim :</td>
<td><input type="text" name="resim" id="resim" value="<%=pusduzen("resim")%>"></td>
</tr>
<tr>
<td>Ücret :</td>
<td><input type="text" name="ucret" value="<%=pusduzen("ucret")%>"></td>
</tr>
<tr>
<td>Adet :</td>
<td><input type="text" name="adet" value="<%=pusduzen("adet")%>"></td>
</tr>
<tr>
<td>Hit :</td>
<td><input type="text" name="hit" value="<%=pusduzen("alindi")%>"></td>
</tr>
<tr>
<td>Detay</td>
<td><textarea name="detay" cols="25" rows="3"><%=pusduzen("detay")%></textarea></td>
</tr>
<tr>
<td>Premium Item</td>
<td><select name="preitem">
<option value="1" <% If pusduzen("premiumitem")=1 Then Response.Write "selected"%>>Evet</option>"
<option value="0" <% If pusduzen("premiumitem")=0 Then Response.Write "selected"%>>Hayır</option>"
</select></td>
</tr>
<tr>
<td>Premium Item Süresi</td>
<td><input type="text" name="sure" size="4" value="<%=pusduzen("kullanimgunu")%>"> Gün</td>
</tr>
<tr>
<td>Item Title</td>
<td><input type="text" name="title" size="4" value="<%Set itemtit=Conne.Execute("select reqtitle from item where num="&pusduzen("itemkodu")&"")
Response.Write itemtit(0)%>"> (Özel Itemler Içın Kullanılır. Sadece Bu Title'a sahip olanlar bu itemi kullanabilir.)</td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="Güncelle" style="width:300px"></td>
</tr>
</table>
</form>

<script type="text/javascript">
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'keyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<form action="javascript:chng();"  method="post" id="search" name="search">
<table>
<tr valign="top">
    <td>
	<table >
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</select>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select>
	</tr></table>
      <input type="text" style="width:200px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr><td>
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
   </td>
	</tr> 
</table></form>
<%
pusduzen.close
Set pusduzen=nothing
End If
Case "pusitemadd"
If Session("durum")="esp" Then
dim itemadd
Set itemadd=Conne.Execute("select top 1 id from pus_itemleri order by id desc")
%>
<form action="default.asp?w8=pusitemadd2" method="post">
  <table border="0">
<tr><td><a href="default.asp?w8=pus">Geri Dön</a></td></tr>
    <tr>
      <td>ID :</td>
      <td><input type="text" name="id" value="<%=itemadd("id")+1%>"></td>
    </tr>
    <tr>
      <td>Type :</td>
      <td><select name="type">
      <option value="scroll">Scroll</option>
      <option value="kiyafet">Kıyafet</option>
      <option value="taki">Takı</option>
      <option value="silah">Silah</option>
      <option value="kalkan">Kalkan</option>
      <option value="anasayfa">Anasayfa</option>
      </select></td>
    </tr>
    <tr>
      <td>Item Ismi :</td>
      <td><input type="text" name="itemismi" id="itemismi"></td>
    </tr>
    <tr>
      <td>Ücret :</td>
      <td><input type="text" name="ucret"></td>
    </tr>
    <tr>
      <td>Adet :</td>
      <td><input type="text" name="adet" value="1"></td>
    </tr>
    <tr>
      <td>Item Kodu :</td>
      <td><input type="text" name="itemkodu" id="itemn" onKeyUp="javascript:rsm(this.value)"></td>
    </tr>
    <tr>
      <td>Resim :</td>
      <td><input type="text" name="resim" id="resim"></td>
    </tr>

    <tr>
      <td>Detay</td>
      <td><textarea name="detay" cols="25" rows="3"></textarea></td>
    </tr>
<tr>
<td>Premium Item</td>
<td><select name="preitem">
<option value="1" >Evet</option>
<option value="0" selected>Hayır</option>
</select></td>
</tr>
<tr>
<td>Premium Item Süresi</td>
<td><input type="text" name="sure" size="4"> Gün</td>
</tr>
<tr>
<td>Item Title</td>
<td><input type="text" name="title" size="4">(Özel Itemler Içindir.Sadece Bu Title'a sahip karakterler bu itemi kullanabilir.Bilmiyorsanız Boş Bırakın)</td>
</tr>
    <tr>
      <td colspan="2" align="center"><input type="submit" value="Ekle" style="width:300px"></td>
    </tr>
  </table>
</form>
<script type="text/javascript">
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'keyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
<form action="javascript:chng();"  method="post" id="search" name="search">
<table>
<tr valign="top">
    <td>
	<table>
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</select>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select></td>
	</tr></table>
      <input type="text" style="width:200px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr><td>
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
    </td>
	</tr> 
</table>
</form>
<%End If
Case "pusitemadd2"
If Session("durum")="esp" Then
id=Request.Form("id")
itype=Request.Form("type")
itemismi=Request.Form("itemismi")
ucret=Request.Form("ucret")
adet=Request.Form("adet")
itemkodu=Request.Form("itemkodu")
resm=Request.Form("resim")
hit=0
detay=Request.Form("detay")
sure=Request.Form("sure")
preitem=Request.Form("preitem")
title=Request.Form("title")
If sure="" Then
sure=0
End If
If title="" Then
title=0
End If

Conne.Execute("Update item Set reqtitle="&title&" where num="&itemkodu)

If id="" or itype="" or itemismi="" or ucret="" or adet="" or resm=""  Then
Response.Write "Boş alan bırakmayınız."
Response.End
End If

Dim pusitemekle
Set pusitemekle = Server.CreateObject("ADODB.Recordset")
pus = "Select * From pus_itemleri"
pusitemekle.open pus,Conne,1,3
pusitemekle.addnew
pusitemekle("id")=id
pusitemekle("type")=itype
pusitemekle("itemismi")=itemismi
pusitemekle("ucret")=ucret
pusitemekle("adet")=adet
pusitemekle("itemkodu")=itemkodu
pusitemekle("resim")=resm
pusitemekle("alindi")=hit
pusitemekle("detay")=detay
pusitemekle("premiumitem")=preitem
pusitemekle("kullanimgunu")=sure
pusitemekle.Update
pusitemekle.close
Set pusitemekle=nothing
Response.Redirect("default.asp?w8=pus")
End If
Case "pussil"
If Session("durum")="esp" Then
id=Request.Querystring("id")
Set pussil=Conne.Execute("delete pus_itemleri where id='"&id&"'")
Response.Redirect "default.asp?w8=pus"
End If
Case "20"
If Session("durum")="esp" Then %>

        <script src="_inc/jquery.hotkeys.js"></script>
<script language="javascript">
  function icerikal(){
  $.ajax({
  type: 'GET',
  url: 'bankaitemleribul.asp',
  data: 'charid='+ encodeURI( document.getElementById("charid").value ),
  start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
  success: function(ajaxCevap) {
$('#itemler').html(ajaxCevap);
}
});
}

function chng()
{
  $.ajax({
	type: 'POST',
	url: 'bkeyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
function itemozell(charid,inventoryslot){
  $.ajax({
	type: 'GET',
	url: 'bankaitemleri.asp?charid='+charid+'&inventoryslot='+inventoryslot,
	success: function(ajaxCevap) {
		$('div#itemoz').html(ajaxCevap);
	}
  });
}
 
    function loadpage(syf){
    $.ajax({
       type: 'GET',
       url: syf,
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
    }


	function itemekle(rsm,num,dur){
	var ids=document.getElementById('inventoryslot').value
	document.getElementById(ids).innerHTML='<img src="../item/'+rsm+'>';
	document.getElementById('num').value=num;
	document.getElementById('serial').value='0';
	document.getElementById('dur').value=dur;
	document.getElementById('stacksize').value='1';
	eval(itemkayt())
	}

	function itemsil(slot)
	{
	var ids=document.getElementById('inventoryslot').value
	document.getElementById(ids).innerHTML='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
	document.getElementById('num').value='0';
	document.getElementById('serial').value='0';
	document.getElementById('dur').value='0';
	document.getElementById('stacksize').value='0';
	eval(itemkayt())
	}


	function itemkayt()
	{
	  $.ajax({
		type: 'post',
		url: 'bankaitemlerikaydet.asp?kyt=one',
		data: $('#itemk').serialize(),
	  });
	document.getElementById('but').disabled=false
	}

	function itemkayt2()
	{
	  $.ajax({
		type: 'post',
		url: 'bankaitemlerikaydet.asp?kyt=all',
		data: $('#itemk').serialize(),
	
	  });
	document.getElementById('but').disabled=true
	}
	
function domo(){
jQuery(document).bind('keydown', 'Ctrl+s',function (evt){itemkayitall();document.getElementById('but').disabled=true; return false;});
jQuery(document).bind('keydown', 'Ctrl+c',function (evt){$cnum=$('#num').val();
$cserial=$('#serial').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cserial').val($cserial);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#cicon').val($cicon);
$('#citemmname').val($itemmname);
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+x',function (evt){$cnum=$('#num').val();
$cserial=$('#serial').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cserial').val($cserial);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#citemmname').val($itemmname);
$('#cicon').val($cicon);
$('#num').val('0');
$('#serial').val('0');
$('#dur').val('0');
$('#stacksize').val('0');
$('div#itemmname').html('');
itemkayit();
$('#'+$cinventoryslot).html('<img height="45" width="45" src="../imgs/blank.gif">');
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+v',function (evt){$cnum=$('#cnum').val();
$cserial=$('#cserial').val();
$cdur=$('#cdur').val();
$cstacksize=$('#cstacksize').val();
$cinventoryslot=$('#cinventoryslot').val();
$inventoryslot=$('#inventoryslot').val();
$cicon=$('#cicon').val();
$itemmname=$('#citemmname').val();
if($cicon=='')$cicon='<img height="45" width="45" src="../imgs/blank.gif">';
document.getElementById('but').disabled=false;
$('#num').val($cnum);
$('#serial').val($cserial);
$('#dur').val($cdur);
$('#stacksize').val($cstacksize);
$('#itemmname').html($itemmname);
$('#'+$inventoryslot).html($cicon);
itemkayit();
return false;
});


jQuery(document).bind('keydown', 'del',function (evt){var ids=document.getElementById('inventoryslot').value;
	document.getElementById('but').disabled=false;
	itemsil(ids);
});


jQuery(document).bind('keydown', 'left',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==0){slt=42}
	selectedbox(slt-1);
	itemozell('',slt-1);
});

jQuery(document).bind('keydown', 'right',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==41){slt=-1}
	selectedbox(slt+1);
	itemozell('',slt+1);
});

jQuery(document).bind('keydown', 'up',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt<14&&slt>0){slt=slt-3}
	if (slt<21&&slt>13){slt=13}
	if (slt>20&&slt<42){slt=slt-7}
	selectedbox(slt);
	itemozell('',slt);
});

jQuery(document).bind('keydown', 'down',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt>17&&slt<24){
	slt==23;
	}else if(slt>41&&slt<48){
	slt==47;
	}else if(slt>65&&slt<72){
	slt==71;
	}else if(slt>113&&slt<120){
	slt==119;
	}else if(slt>137&&slt<144){
	slt==143;
	}else if(slt>161&&slt<168){
	slt==167;
	}else if(slt>185&&slt<192){
	slt==191;
	}
	
	
	selectedbox(slt);
	itemozell('',slt);
});

}
            
            
    jQuery(document).ready(domo);
    </script>
<%Set users=Conne.Execute("select * from account_char") %>
<form action="javascript:icerikal(); " method="get"  name="itemform" id="itemform" >
Kullanıcı Adı Seçin :
	<select name="charid" id="charid">
<% If not users.eof Then
do while not users.eof %>
	<option value="<% =trim(users("straccountid"))%>"><% =trim(users("straccountid"))%></option>
    <% users.movenext
	Loop
	Else %><option value="" disabled="disabled" selected="selected">Kullanıcı Bulunamadı</option>
	<% End If %>
</select>
<input type="submit" value="Itemleri Bul" />
</form><span name="itemler" id="itemler"></span>
<%
Else 
End If 
Case "10"
If Session("durum")="esp" Then %>

<form onSubmit="icerikal();$('#chrs').fadeOut(0);return false">
<b>Karakter Adını Yazın: </b><input  type="text" name="charid" id="charid" size="30"><a href="#" onClick="userara();return false">Karakter Adı Bul</a>
<input type="submit" value="ITEMLERI BUL" id="arama" style="width:200px"><br>
<br><div id="chrs" style="width:190px;height:100px;overflow:auto;position:relative;left:138px;top:-12px;background-color:lightgrey;padding-left:5px;padding-top:5px;padding-bottom:5px;border-left:2px;border-left-style: groove;border-right-style: groove; z-index:2"></div></form>
<span name="itemler" id="itemler" style="position:relative; z-index:1"></span>
<script src="_inc/jquery.hotkeys.js"></script>
<script language="javascript">
function icerikal(){
if(document.getElementById('but')!=null&&document.getElementById('but').disabled==false){
if(confirm('\n\n\n\n\n\n\n\n\n\n\n                 ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! \n>>>>>>>>> INVENTORYDEKI ITEMLERI KAYIT ETMEDEN ÇIKIYORSUNUZ. <<<<<<<<<<<<<<<<<\n>>>>>EĞER KAYIT ETMEDEN ÇIKARSANIZ YAPTIĞINIZ DEĞİŞİKLİKLER GEÇERLİ OLMAZ<<<<<   \n>>>>>>>>>>>> ITEMLERI KAYIT ETMEK IÇIN VAZGEÇE TIKLAYINIZ<<<<<<<<<<<<<<<<<\n\n\n\n\n\n\n\n\n\n\n')){
$.ajax({
       type: 'GET',
       url: 'itembul.asp',
       data: 'charid='+ encodeURI( document.getElementById("charid").value ),
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
       $('#itemler').html(ajaxCevap);
       }
    });
document.getElementById("charid").blur();
}
else{
return false;
}
}
$.ajax({
       type: 'GET',
       url: 'itembul.asp',
       data: 'charid='+ encodeURI( document.getElementById("charid").value ),
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
}

function loadpage(syf){
 $.ajax({
       type: 'GET',
       url: syf,
       start: $('#itemler').html('<img src="../imgs/38-1.gif"><br>Itemler Yükleniyor...'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
$('#chrs').fadeOut(0);
    }

function userara(){
$.ajax({
   type: 'post',
   url: 'userbul.asp',
   data: 'struserid='+encodeURI( document.getElementById("charid").value ),
   start:$('#chrs').html("<center><img src='../imgs/38-1.gif'><br>Aranıyor. Lütfen Bekleyin.</center>"),
   success: function(ajaxCevap) {
      $('#chrs').html(ajaxCevap);
   }
});
$('#chrs').fadeIn(0);
}
function userbul(userid){
document.getElementById('charid').value=userid;
}

function kontrolet(olay){
olay = olay || event;
if (olay.keyCode==13){
icerikal();
$('#chrs').fadeOut(0);
}
else{
userara();
$('#chrs').fadeIn(0);
}
}


 function domo(){
jQuery(document).bind('keydown', 'Ctrl+s',function (evt){itemkayitall();document.getElementById('but').disabled=true; return false;});
jQuery(document).bind('keydown', 'Ctrl+c',function (evt){$cnum=$('#num').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cserial').val(0);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#cicon').val($cicon);
$('#citemmname').val($itemmname);
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+x',function (evt){$cnum=$('#num').val();
$cserial=$('#serial').val();
$cdur=$('#dur').val();
$cstacksize=$('#stacksize').val();
$cinventoryslot=$('#inventoryslot').val();
$cicon=$('#'+$cinventoryslot).html();
$itemmname=$('div#itemmname').html();
$('#cnum').val($cnum);
$('#cserial').val($cserial);
$('#cdur').val($cdur);
$('#cstacksize').val($cstacksize);
$('#cinventoryslot').val($cinventoryslot);
$('#citemmname').val($itemmname);
$('#cicon').val($cicon);
$('#num').val('0');
$('#serial').val('0');
$('#dur').val('0');
$('#stacksize').val('0');
$('div#itemmname').html('');
itemkayit();
$('#'+$cinventoryslot).html('<img height="45" width="45" src="../imgs/blank.gif">');
return false;
 });
 
jQuery(document).bind('keydown', 'Ctrl+v',function (evt){$cnum=$('#cnum').val();
$cserial=$('#cserial').val();
$cdur=$('#cdur').val();
$cstacksize=$('#cstacksize').val();
$cinventoryslot=$('#cinventoryslot').val();
$inventoryslot=$('#inventoryslot').val();
$cicon=$('#cicon').val();
$itemmname=$('#citemmname').val();
if($cicon=='')$cicon='<img height="45" width="45" src="../imgs/blank.gif">';
document.getElementById('but').disabled=false;
$('#num').val($cnum);
$('#serial').val($cserial);
$('#dur').val($cdur);
$('#stacksize').val($cstacksize);
$('#itemmname').html($itemmname);
$('#'+$inventoryslot).html($cicon);
itemkayit();
return false;
});


jQuery(document).bind('keydown', 'del',function (evt){var ids=document.getElementById('inventoryslot').value;
	document.getElementById('but').disabled=false;
	itemsil(ids);
});


jQuery(document).bind('keydown', 'left',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==0){slt=42}
	selectedbox(slt-1);
	itemozell('',slt-1);
});

jQuery(document).bind('keydown', 'right',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt==41){slt=-1}
	selectedbox(slt+1);
	itemozell('',slt+1);
});

jQuery(document).bind('keydown', 'up',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt<14&&slt>0){slt=slt-3}
	if (slt<21&&slt>13){slt=13}
	if (slt>20&&slt<42){slt=slt-7}
	selectedbox(slt);
	itemozell('',slt);
});

jQuery(document).bind('keydown', 'down',function (evt){var slt=parseInt($('#inventoryslot').val());
	if (slt<11&&slt>=0){slt=slt+3}
	else if(slt==11){slt=13} 
	else if (slt==13||slt==12){slt=14}
	else if (slt>13&&slt<35){slt=slt+7}
	else if (slt>34&&slt<42){slt=0}
	selectedbox(slt);
	itemozell('',slt);
});

}
            
            
    jQuery(document).ready(domo);

	function itemsil(slot){
	var ids=document.getElementById('inventoryslot').value
	document.getElementById(ids).innerHTML='<img src="../imgs/blank.gif">';
	document.getElementById('num').value='0';
	document.getElementById('serial').value='0';
	document.getElementById('dur').value='0';
	document.getElementById('stacksize').value='0';
	cal('ui_button2.mp3');
	itemkayit();
	document.getElementById('itemmname').innerHTML='&nbsp;';
	document.getElementById('but').disabled=false;
	if (slot>13){
	document.getElementById('adet'+slot).innerHTML='&nbsp;';
		}
	}


function selectedbox(slot){
 for (var i=0; i<42; i++){
  var spn = document.getElementById(i);
  spn.style.border = "1px solid #4C4B36";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
	}



function itemkayit(){
  $.ajax({
	type: 'post',
	url:'invitemkaydet.asp?islem=one',
	data:'num='+$('#num').val()+'&charid='+$('#charid').val()+'&serial='+$('#serial').val()+'&dur='+$('#dur').val()+'&stacksize='+$('#stacksize').val()+'&inventoryslot='+$('#inventoryslot').val() ,
	success:function(ajaxCevap) {
    $('div#itemkayitsonuc').html(ajaxCevap);
    }
  });
}




function itemkayitall(){
  $.ajax({
	type: 'post',
	url: 'invitemkaydet.asp?islem=all',
	data: 'num='+$('#num').val()+'&charid='+$('#charid').val()+'&serial='+$('#serial').val()+'&dur='+$('#dur').val()+'&stacksize='+$('#stacksize').val()+'&inventoryslot='+$('#inventoryslot').val() ,
	success:function(ajaxCevap) {
    $('div#itemkayitsonuc').html(ajaxCevap);
    }
  });
}



 
function loadpage(syf){
   $.ajax({
   type: 'GET',
   url: syf,
   start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
   success: function(ajaxCevap) {
   $('#itemler').html(ajaxCevap);
    }
    });
    }

function itemekle(rsm,num,dur,countable,name,detay,slot){

if (slot==5||slot==6||slot==7||slot==8||slot==9){
cal('item_armor.mp3')
}
else if(slot==1||slot==2||slot==3||slot==4){
cal('item_weapon.mp3')
}
else{
cal('ui_button2.mp3')
}

var ids=document.getElementById('inventoryslot').value;
document.getElementById(ids).innerHTML="<img src=\"../item/"+rsm+"\" onMouseOver=\"return overlib('"+detay+"', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10)\" onMouseOut=\"return nd();\">";
document.getElementById('num').value=num;
document.getElementById('serial').value='0';
document.getElementById('dur').value=dur;
document.getElementById('stacksize').value=countable;

if (ids>13){
document.getElementById('adet'+ids).innerHTML='&nbsp;';
}
document.getElementById('itemmname').innerHTML='<b>'+name+'</b>';
itemkayit()
document.getElementById('but').disabled=false
}


function dis(){
document.getElementById('but').disabled=true
}

window.onbeforeunload = cikis_yap


function cikis_yap(){
if(document.getElementById('but')!=null){
if (document.getElementById('but').disabled==false){
return ('\n\n\n\n\n\n\n\n\n\n\n                 ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! \n>>>>>>>>> INVENTORYDEKI ITEMLERI KAYIT ETMEDEN ÇIKIYORSUNUZ. <<<<<<<<<<<<<<<<<\n>>>>>EĞER KAYIT ETMEDEN ÇIKARSANIZ YAPTIĞINIZ DEĞİŞİKLİKLER GEÇERLİ OLMAZ<<<<<   \n>>>>>>>>>>>> ITEMLERI KAYIT ETMEK IÇIN VAZGEÇE TIKLAYINIZ<<<<<<<<<<<<<<<<<\n\n\n\n\n\n\n\n\n\n\n')

}
}
}


function stacksizeupdate(slot,no){
if (slot>13){
if (no==''||no==0||no=='0'){
no='&nbsp;'
}
document.getElementById('adet'+slot).innerHTML=no
}
}

function dragStart(ev) {
    ev.dataTransfer.effectAllowed='move';
    //ev.dataTransfer.dropEffect='move';
    ev.dataTransfer.setData("Text", ev.target.getAttribute('id'));
    ev.dataTransfer.setDragImage(ev.target,0,0);
    return true;
}

function dragEnter(ev) {
    var idelt = ev.dataTransfer.getData("Text");
    return true;
}

function dragOver(ev) {
    var idelt = ev.dataTransfer.getData("Text");
    var id = ev.target.getAttribute('id');
    if( (id =='b0' || id =='b2'|| id =='b1') && (idelt=='i1'|| idelt=='i2'|| idelt=='i3'))
        return false;
    else
        return true;
}

function dragEnd(ev) {
    ev.dataTransfer.clearData("Text");
    return true
}

function dragDrop(ev) {
    var idelt = ev.dataTransfer.getData("Text");
    ev.target.appendChild(document.getElementById(idelt));
    ev.stopPropagation();
    return false; // return false so the event will not be propagated to the browser
}

</script>
<style type="text/css">
#b0, #b2, #b1 { 
  float:left; width:50px; height:55px; padding:10px; margin:10px; 
}

#b1 { background-color: #474747; }
#b2 { background-color: #474747; }
#b0 { background-color: #474747; }


</style>
<%
Else 
End If 
Case "ticket"
If Session("durum")="esp" Then
dim tick
Set tick = Server.CreateObject("ADODB.Recordset")
strSQL = "Select * From tickets ORDER BY durum asc ,date DESC "
tick.open strSQL,Conne,1,3%>
<table width="577" border="0">
  <tr bgcolor="#FD6602" >
	<td width="67" align="center" bgcolor="#FFCC00" class="style5">Durum</td>
    <td width="125" align="center" bgcolor="#FFCC00" class="style5">Yollayan</td>
    <td width="149" align="center" bgcolor="#FFCC00" class="style5">Konu</td>
    <td width="76" align="center" bgcolor="#FFCC00" class="style5">Tarih</td>
    <td width="138" align="center" bgcolor="#FFCC00" class="style5">Ticket Oku</td>
  </tr>
  <%If not tick.eof Then
  do while not tick.eof%>
  <tr bgcolor="#E4E4E4" >
	<td align="center"><%If tick("durum")="1" Then
	Response.Write "<img src='../imgs/Mail_new.gif'><br>Okunmadı !"
	ElseIf tick("durum")="2" Then
	Response.Write "<img src='../imgs/Mail_verified.gif'><br>Okundu" 
	End If%></td>
    <td align="center" bgcolor="#E4E4E4" ><%=tick("charid")%></td>
    <td align="center" bgcolor="#E4E4E4" ><%=tick("subject")%></td>
    <td align="center" bgcolor="#E4E4E4" class="produces"><%=tick("date")%></td>
	<td width="138" align="center" bgcolor="#E4E4E4"><a href="?w8=ticketoku&id=<%=tick("id")%>"><font style="font-size:11px">Oku</font></a></td>
  <% tick.movenext
  Loop
  tick.Close
  Set tick=Nothing%>
</table>
<% Else
Response.Write "Ticket Yok" 
End If
End If


Case "ticketoku"
If Session("durum")="esp" Then
olay = trim(Request.Querystring("olay"))
If olay="" Then
olay="0"
End If
select Case olay
Case "ticketcevap"
ticketcevap=trim(request.Form("ticketcevap"))
id=Request.Querystring("id")
If ticketcevap="" Then 
Response.Redirect "default.asp?w8=ticketoku&id="&id
Response.End
End If
Set ticketUpdate = Server.CreateObject("ADODB.Recordset")
sSQL = ("select * from tickets where id='"&id&"'")
ticketUpdate.open sSQL,Conne,1,3
ticketUpdate("durum")=2
ticketUpdate("cevap")=ticketcevap
ticketUpdate.Update
Response.Write "<script>alert('Ticket Cevaplandı !')</script>"
Case "0"
Case Else

end select 

id=Request.Querystring("id")
Set tci = Server.CreateObject("ADODB.Recordset")
strSQL = "Select * From tickets where id='"&id&"'"
tci.open strSQL,Conne,1,3
If not tci.eof Then
tci("durum")=2
tci.Update
End If
%>
<table width="747" border="0" align="left">
  <tr>
    <td width="120" bgcolor="#FD6602"><div align="center" class="text6">Account name</div></td>
     <td width="120" bgcolor="#FD6602"><div align="center" class="text6">Gönderen</div></td>
    <td width="137" bgcolor="#FD6602"><div align="center" class="text6">Email</div></td>
    <td width="130" bgcolor="#FD6602"><div align="center" class="text6">Tarih</div></td>
  </tr>
  <tr>
    <td bgcolor="#E4E4E4"><div align="center" class="style7"><%=tci("charid")%></div></td>
    <td bgcolor="#E4E4E4"><div align="center" class="style7"><%=tci("name")%></div></td>
    <td bgcolor="#E4E4E4"><div align="center" class="style7"><%=tci("email")%></div></td>
     <td bgcolor="#E4E4E4"><div align="center" class="style7"><%=tci("date")%></div></td>
    </tr>
<tr>
    <td width="130" bgcolor="#FD6602" colspan="3"><div align="center" class="text6">Konu</div></td>
</tr>
<tr>
    <td bgcolor="#E4E4E4" colspan="3"><div align="center" class="style7"><%=tci("subject")%></div></td></tr>
    <tr><td colspan="4" bgcolor="#FD6602" width="500"><div align="center" class="text6">Mesaj</div></td>
        <td width="75" bgcolor="#FD6602"><div align="center" class="text6">Ticket Sil</div></td>
    </tr>
  <tr>
    <td colspan="4" align="center" bgcolor="#E4E4E4" class="style5"><br><%=tci("message")%></td>
    <td align="center" bgcolor="#E4E4E4"><a href="default.asp?w8=ticketsil&sid=<%=tci("id")%>" title="Ticketi Sil"><img src="../imgs/Mail_delete.gif" border="0" alt="sil"></a></td>
  </tr>
  <tr><td>&nbsp;</td></tr>
<tr>
<td colspan="4" align="center" bgcolor="#FD6602" class="text6">Cevap Yaz</td></tr>
<td colspan="5">
<form action="default.asp?w8=ticketoku&id=<%=tci("id")%>&olay=ticketcevap" method="post">
<textarea name="ticketcevap" id="ticketcevap" cols="100" rows="12" style="font-family:Verdana; font-size:10px"><%=tci("cevap")%></textarea>
       <script language="JavaScript">
  generate_wysiwyg('ticketcevap');
</script>
<input type="submit" value="Ticketı Cevaplandır." class="inputstyle"  style="width:600; font-size:11px;">
</form></td>
</table>
<%
ElseIf Session("durum")="sup" Then

End If
Case "ticketsil" 
If Session("durum")="esp" Then
sid = trim(Request.Querystring("sid"))

Set dtci = Server.CreateObject("ADODB.Recordset")
strSQL = "Delete From tickets Where id="&sid&""
dtci.open strSQL,Conne,1,3

Response.Redirect "default.asp?w8=ticket"
Else 
Response.Redirect "default.asp"
End If 



Case "7" 

If Session("durum")="esp" Then
Set accounts=Conne.Execute("Select * from TB_USER")%>
<form action="default.asp?w8=7xkulldetay" method="POST">
<font face="Verdana" style="font-size:9pt;"><b>Kullanıcı Adı : </b></font>
	<select name="strAccountID" id="strAccountID">
	<% If not accounts.eof Then
	do while not accounts.eof %>
	<option value="<% =accounts("strAccountID")%>"><% =accounts("strAccountID")%></option>
    <% accounts.movenext
	Loop
	Else %><option value="" disabled="disabled" selected="selected">Kullanıcı Bulunamadı</option>
	<% End If %>
	</select> 
<input type="submit" value="Ara">
</form>
<%Else
End If 

Case "7xkulldetay"
If Session("durum")="esp" Then

strAccountID=Request.Form("strAccountID")
If Session("durum")="esp" Then 
Set kull = Server.CreateObject("ADODB.Recordset")
strSQL = "Select * From [TB_USER] Where strAccountID='"&strAccountID&"'"
kull.open strSQL,Conne,1,3

Set chars=Conne.Execute("select * from account_char where straccountid='"&strAccountID&"'")

If not kull.eof Then %>
<form action="default.asp?w8=7xUpdatekull" method="POST">
<table width="250">
  <tr>
    <td><b>Giriş Id :</b></td>
    <td><input type="text" name="strgirisid" value="<%=kull("strAccountID")%>"/></td>
  </tr>
  <tr>
    <td><b>Şifre : </b></td>
    <td><input type="text" name="strPasswd" value="<%=kull("strPasswd")%>" maxlength="13"></td>
  </tr>
  <tr>
    <td><b>Account Durumu : </b></td>
    <td><select name="aut"><option value="0" <%If kull("strauthority")="0" Then Response.Write "selected"%>>Yönetici</option>
<option value="1" <%If kull("strauthority")="1" Then Response.Write "selected"%>>Normal Kullanıcı</option>
<option value="255" <%If kull("strauthority")="255" Then Response.Write "selected"%>>Banlı</option></select></td>
  </tr>
  <% If not chars.eof Then %>
  <tr>
    <td><b>1. Char : </b></td>
    <td> <input type="text" value="<%=trim(chars("strcharid1"))%>" name="char1"></td>
  </tr>
  <tr>
    <td><b>2. Char :</b></td>
    <td><input type="text" value="<%=trim(chars("strcharid2"))%>" name="char2"></td>
  </tr>
  <tr>
    <td><b>3. Char :</b></td>
    <td><input type="text" value="<%=trim(chars("strcharid3"))%>" name="char3"></td>
  </tr>
    <tr>
    <td><b>Char Sayısı :</b></td>
    <td><input type="text" value="<%=trim(chars("bcharnum"))%>" name="num"></td>
  </tr>
  <% End If %>
</table>


<br>
<br>

<input type="hidden" value="<%=kull("strAccountID")%>" name="strAccountID">
<input type="submit" value="Kaydet"></form>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Kullanıcı Adı Bulunamadı</font>
<% End If
Else %>
<font face="Verdana" style="font-size:9pt;">Şifre Girilmemiş</font>
<% End If
ElseIf Session("durum")="sup" Then
End If
Case "7xUpdatekull" 
If Session("durum")="esp" Then 

strgirisid=Request.Form("strgirisid")
char1=trim(Request.Form("char1"))
char2=trim(Request.Form("char2"))
char3=trim(Request.Form("char3"))
num=trim(Request.Form("num"))
strAccountID=Request.Form("strAccountID")
strPasswd=Request.Form("strPasswd")
straut=Request.Form("aut")

If char1="" Then
char1=null
End If
If char2="" Then
char2=null
End If
If char3="" Then
char3=null
End If

Set kulls = Server.CreateObject("ADODB.Recordset")
strSQL = "Select * From [TB_USER] Where strAccountID='"&strAccountID&"'"
kulls.open strSQL,Conne,1,3
Set kulls2 = Server.CreateObject("ADODB.Recordset")
strSQL = "Select * From [ACCOUNT_CHAR] Where strAccountID='"&strAccountID&"'"
kulls2.open strSQL,Conne,1,3

If not kulls.eof Then
If not kulls2.eof Then
kulls2("strcharid1")=char1
kulls2("strcharid2")=char2
kulls2("strcharid3")=char3
kulls2("bcharnum")=num
kulls2.Update
End If
kulls("strAccountID")=strgirisid
kulls2("strAccountID")=strgirisid
kulls("strPasswd")=strPasswd
kulls("strauthority")=straut
kulls.Update
kulls2.Update
%>
<meta http-equiv="refresh" content="1;url=default.asp?w8=7">
<font face="Verdana" style="font-size:9pt;">
Güncelleme Başarılı !</font>
<br />
Yönlendiriliyorsunuz.<br />
<img src="../imgs/18-1.gif" />
<% Else %>

<meta http-equiv="refresh" content="1;url=default.asp?w8=7">
<font face="Verdana" style="font-size:9pt;">
Böyle bir Kayıt Bulunamadı
<br />
Yönlendiriliyorsunuz.</font><br />
<img src="../imgs/18-1.gif" />
<% End If 

ElseIf Session("durum")="sup" Then
End If
Case "premium"
If Session("durum")="esp" Then
Dim accounts
Set accounts=Conne.Execute("select * from tb_user")%>
<form action="default.asp?w8=premium2" method="post"><b>Premium & Cash Editor</b><br><br>
	Kullancı seçin : <select name="strAccountID" id="strAccountID">
	<% If not accounts.eof Then
	do while not accounts.eof %>
	<option value="<% =accounts("strAccountID")%>"><% =accounts("strAccountID")%></option>
    <% accounts.movenext
	Loop
	Else %><option value="" disabled="disabled" selected="selected">Kullanıcı Bulunamadı</option>
	<% End If %>
    </select><input type="submit" value="Sonraki>>">
</form>
<br>
<a href="default.asp?w8=premium&sayfa=cashtablo" class="e">Cash Tablosu</a><br>
<a href="default.asp?w8=premium&sayfa=cash" class="e">Cash Kodu Üret</a>
<%If Request.Querystring("sayfa")="cash" Then
sub coderepeat
randomize
Dim Kod
Kod=int(rnd*8999)+1000&"-"&int(rnd*8999)+1000&"-"&int(rnd*8999)+1000&"-"&int(rnd*8999)+1000&"-"&int(rnd*8999)+1000
Response.Write kod
end sub
%>
<form action="default.asp?w8=premium&sayfa=cash2" method="post">
<table>
<tr><td>Cash Kodu :</td>
<td><input type="text" name="kod" value="<%coderepeat%>" size="30"></td>
</tr>
<tr>
<td>Cash Miktari :</td>
<td><input type="text" name="miktar"></td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="Databaseye Ekle" class="inputstyle"></td></tr>
</table>
</form>
<%ElseIf Request.Querystring("sayfa")="cashtablo" Then
Dim cashs
Set cashs=Conne.Execute("select * from cash_table order by durum asc")
Response.Write "<br><br><table border=""1"" cellspacing=""0"" width=""450""><tr><td align=""center""><b>Cash Code</b></td><td align=""center""><b>Cash Miktarı</b></td><td align=""center""><b>Durum</b></td><td align=""center""><b>Alan Char</b></td></tr>"
do while not cashs.eof
Response.Write "<tr><td>"&cashs("cashcode")&"</td><td>"&cashs("cashmiktar")&"</td><td> "&cashs("durum")&"</td><td>&nbsp;"&cashs("alanchar")&"<td><td><a href=""default.asp?w8=premium&sayfa=sil&id="&cashs("cashid")&""">Sil</a></td></tr>"
cashs.movenext
Loop

ElseIf Request.Querystring("sayfa")="cash2" Then
Dim kod,miktar,cashadd
kod=Request.Form("kod")
miktar=Request.Form("miktar")

Set cashadd = Server.CreateObject("ADODB.Recordset")
SQL ="SELECT * FROM cash_table"
cashadd.open SQL,Conne,1,3

cashadd.addnew
cashadd("cashcode")=kod
cashadd("cashmiktar")=miktar
cashadd.Update
cashadd.close
Set cashadd=nothing
Response.Write "<br><br>"&miktar&" Cash Eklendi"

ElseIf Request.Querystring("sayfa")="sil" Then
id=Request.Querystring("id")
Conne.Execute("delete cash_table where cashid="&id&"")
Response.Redirect("default.asp?w8=premium&sayfa=cashtablo")
End If
Else
End If
Case "premium2"
If Session("durum")="esp" Then
dim charid,pre
charid=Request.Form("straccountid")
Set pre=Conne.Execute("select * from tb_user where straccountid='"&charid&"'")
If not pre.eof Then
%><form action="default.asp?w8=premium3" method="post"><table border="0">
  <tr>
    <td>Premium :</td>
    <td><select name="premium">
      <option value="0" <% If pre("premiumtype")="0" Then
Response.Write "selected"
End If%>>Yok</option>
      <option value="1" <% If pre("premiumtype")="1" Then
Response.Write "selected"
End If%> >Var</option>
    </select></td>
  </tr>
  <tr>
    <td>Premium Gün :</td>
    <td><input type="text" name="pregun" value="<%=round(pre("premiumdays")-now)%>"></td>
  </tr>
  <tr>
    <td>Cash Point :</td>
    <td><input type="text" name="cash" value="<%=pre("cashpoint")%>"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="hidden" value="<%=charid%>" name="charid"><input type="submit" value="Güncelle"></td>
    </tr>
</table>
</form>

<%
Else
Response.Redirect("default.asp?w8=premium")
End If
Else
End If
Case "premium3"
If Session("durum")="esp" Then
Dim prem,pregun,cash,premium,stracSQL
charid=Request.Form("charid")
prem=Request.Form("premium")
pregun=Request.Form("pregun")
cash=Request.Form("cash")

Set premium = Server.CreateObject("ADODB.Recordset")
stracSQL = "Select * From tb_user Where strAccountID='"&charid&"'"
premium.open stracSQL,Conne,1,3

premium("premiumtype")=prem
premium("premiumdays")=now+pregun
premium("cashpoint")=cash
premium.Update

premium.close
Set premium=nothing
Response.Redirect("default.asp?w8=premium")
Else
End If
Case "12"
If Session("durum")="esp" Then%>
<script type="text/javascript">
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'keyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}

</script>
<table>
	<tr valign="top">
    <td><form action="default.asp?w8=13" method="post"  name="itemn">İtem No Girin: 
	<input type="text" name="itemno" id="itemn"><input type="submit" value="İtem Düzenle">
	</form></td>
</tr>
    <tr valign="top">
    <td>
	<form action="javascript:chng();"  method="post" id="search" name="search">
	<table>
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</select>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select></td>
	</tr>
	<tr><td colspan="5">
	<input type="text" style="width:200px;"  name="keyw" id="keyw" />
    <input name="submit" type="submit" value="Item Ara">
	</td></tr></table></form>
	</td>
	</tr>
	<tr><td>
      <div id="sonuc" style="width:300px; background-color:silver;">.. Item İsmini Yazın</div>
   </td>
	</tr>
</table>

<%Else
Response.Redirect("default.asp")
End If
Case "13"
If Session("durum")="esp" Then
itemno=Request.Form("itemno")
Set itemkontrol=Conne.Execute("select * from item where num='"&itemno&"'")
If itemkontrol.eof Then
Response.Write "İtem Bulunamadı!!!"
Response.End
End If
Set cno=Conne.Execute("SELECT * FROM sysobjects where name='item'")
Set columns=Conne.Execute("select name from syscolumns where id="&cno("id")&" order by colorder")%>
<form action="default.asp?w8=14&itemnum=<%=itemno%>" method="post">
<table><tr>
<%
Set itemdetay=Conne.Execute("select * from item where num='"&itemno&"'")
x=1
do while not columns.eof
If not itemdetay.eof Then
If x mod 2=0 Then
Response.Write "</tr>"&vbcrlf&"<tr>"
End If
Response.Write "<td>"&columns("name")&"</td>"
Response.Write "<td><input type=""text"" name="""&columns("name")&""" value="""&trim(itemdetay(""&columns("name"))&"")&"""></td>"


Else
Response.Write("İtem Bulunamadı !")
End If
x=x+1
columns.movenext
Loop%>
<tr><td><a href="default.asp?w8=14&sil=<%=itemno%>"><font class="style5">İtemi Sil</font></a></td>
</tr>
<tr><td colspan="2"><input type="submit" value="KAYDET" style="width:300PX; height:50px"></td></tr>
</table></form>
<%Else
Response.Redirect("default.asp")
End If
Case "14"
If Session("durum")="esp" Then
If not Request.Querystring("sil")="" Then
Conne.Execute("delete item where num="&Request.Querystring("sil")&"")
Else
itemnum=Request.Querystring("itemnum")
For Each col In Request.Form
Conne.Execute("Update item Set "&col&"='"&Request.Form(col)&"' where num='"&itemnum&"'")
next
on error resume next
Response.Redirect("default.asp?w8=12")
End If
Else
Response.Redirect("default.asp")
End If
Case "9"
If Session("durum")="esp" Then
Set accounts=Conne.Execute("select * from tb_user")%>

<form action="default.asp?w8=9search" method="POST">
<font face="Verdana" style="font-size:9pt;"><b>Kullanıcı Adı (Login) : </b></font><br>
	<select name="strAccountID" id="strAccountID">
	<% If not accounts.eof Then
	do while not accounts.eof %>
	<option value="<% =accounts("strAccountID")%>"><% =accounts("strAccountID")%></option>
    <% accounts.movenext
	Loop
	Else %><option value="" disabled="disabled" selected="selected">Kullanıcı Bulunamadı</option>
	<% End If %>
    </select>
<br>
<input type="submit" value="Ara"></form>
<br><center><a href="default.asp?w8=idbul"><font face="Verdana" style="font-size:9pt;">Eğer Kullanıcı Adını Bilmiyorsanız Buraya Tıklayın ve 
Karakter Adını Girerek Kullanıcı Adı Ortaya Çıksın . </font></a></center>
<%ElseIf Session("durum")="sup" Then

End If 
Case "9search"
If Session("durum")="esp" Then

strAccountID=Request.Form("strAccountID")
Set ac = Server.CreateObject("ADODB.Recordset")
stracSQL = "Select * From ACCOUNT_CHAR Where strAccountID='"&strAccountID&"'"
ac.open stracSQL,Conne,1,3

 If not ac.eof Then %>
<form action="default.asp?w8=9xsearchxnationsecildi" method="POST">
<font face="Verdana" style="font-size:9pt;"><b>Irk Seçimi : </b>
<select name="bNation"><option value="1" <%If ac("bNation")="1" Then%>selected<%ElseIf ac("bNation")="2" Then
End If%>>Karus</option>
<option value="2" <%If ac("bNation")="2" Then%>selected<%ElseIf ac("bNation")="1" Then
End If%>>Human</option></select><br>
<input type="hidden" value="<%=ac("strAccountID")%>" name="strAccountID">
<input type="submit" value="Giriş"></font></form>
<% Else %>
<font face="Verdana" style="font-size:9pt;">Böyle Bir Kullanıcı Mevcut Değil.</font>
<% End If 
ElseIf Session("durum")="sup" Then
End If 
Case "9xsearchxnationsecildi"
If Session("durum")="esp" Then

 strAccountID=Request.Form("strAccountID")
 bNation=Request.Form("bNation")

Set acs = Server.CreateObject("ADODB.Recordset")
stracsSQL = "Select * From [ACCOUNT_CHAR] Where strAccountID='"&strAccountID&"'"
acs.open stracsSQL,Conne,1,3

 If not acs.eof Then 

acs("bNation")=bNation
acs.Update
%>
<meta http-equiv="refresh" content="1;url=default.asp?w8=9">
<font face="Verdana" style="font-size:9pt;">Güncelleme Başarılı.
<br />
Yönlendiriliyorsunuz.<br />
<img src="../imgs/18-1.gif" />
<% Else %>

<font face="Verdana" style="font-size:9pt;">Böyle Bir Kullanıcı Mevcut Değil.
<br />
Yönlendiriliyorsunuz.<br />
<img src="../imgs/18-1.gif" />
<% End If 
 ElseIf Session("durum")="sup" Then
 End If 
 Case "idbul" %>

<br><i>Kişinin Karakter Adını Giriniz </i>
<form action="default.asp?w8=idbulsearch" method="POST">
<br>
<font face="Verdana" style="font-size:9pt;"><b>Karakter Adı  : </b></font><br>
<input type="text" name="strChar" onKeyUp="karakterfiltre(this);"><br>
<input type="submit" value="Ara"></form>
<% Case "idbulsearch" 
If Session("durum")="esp" Then

strChar=request("strChar")

If strChar="" Then
Response.Redirect "default.asp?w8=idbul"
End If

Set acs = Server.CreateObject("ADODB.Recordset")
stracsSQL = "Select * From [ACCOUNT_CHAR] Where strCharID1='"&strChar&"' or strCharID2='"&strChar&"' or strCharID3='"&strChar&"'"
acs.open stracsSQL,Conne,1,3

If not acs.eof Then %>
<font face="Verdana" style="font-size:9pt;">
<b>Login ID : </b> <%=acs("strAccountID")%><br><br>
<b>Kullanıcının Karakterleri : </b><br>

<b>1.</b> <% 
If acs("strCharID1")="" Then
Response.Write "Karakter Yok"
Else
Response.Write acs("strCharID1")&"<a href=default.asp?w8=idbulsearch&strchar="&strchar&"&strChar1="&acs("strCharID1")&">&nbsp;Sil</a>"
End If %>
<br>
<b>2.</b> <% If acs("strCharID2")="" Then
Response.Write "Karakter Yok"
Else
Response.Write acs("strCharID2")&"<a href=default.asp?w8=idbulsearch&strchar="&strchar&"&strChar2="&acs("strCharID2")&">&nbsp;Sil</a>"
End If %>
<br>
<b>3.</b>
<% 
If acs("strCharID3")="" Then
Response.Write "Karakter Yok"
Else
Response.Write acs("strCharID3")&"<a href=default.asp?w8=idbulsearch&strchar="&strchar&"&strChar3="&acs("strCharID3")&">&nbsp;Sil</a>"
End If 

strChar1=Request.Querystring("strChar1")
strChar2=Request.Querystring("strChar2")
strChar3=Request.Querystring("strChar3")

If not strChar1="" or strChar2="" or strChar3="" Then

Set charsil = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From [ACCOUNT_CHAR] Where strCharID1='"&strChar1&"' or strCharID2='"&strChar2&"' or strCharID3='"&strChar3&"'"
charsil.open SQL,Conne,1,3

If not charsil.eof Then

If not strChar1=""  Then

charsil("strcharid1")=NULL
charsil("bcharnum")=charsil("bcharnum")-1
charsil.Update

ElseIf not strChar2="" Then

charsil("strcharid2")=NULL
charsil("bcharnum")=charsil("bcharnum")-1
charsil.Update

ElseIf not strChar3=""  Then
charsil("strcharid3")=NULL
charsil("bcharnum")=charsil("bcharnum")-1
charsil.Update
End If

Else
End If
Else
Response.Write "Karakter silinemedi."
End If
Else 
Response.Write ("<font face='Verdana' style='font-size:9pt;'>Böyle Bir Kullanıcı Mevcut Değil. "&strchar &"</font>")
End If
End If

End Select
%></span></td>
      </tr>
</table>
  </tr>
</table>
<%Response.Write "<center><font face='Verdana' style='font-size:9pt;' color='white'>Copyright &copy; 2007 WebMsn<br>Edited By Asi Beşiktaşlı 2011</font></center>" %>
</table>
<div id="player" ></div>
<script type="text/javascript" src="Flash/swfobject.js"></script>
<script type="text/javascript">
function cal(url){
var so = new SWFObject('Flash/player.swf','mpl','100','100','9');
so.addVariable('file', 'http://localhost/sounds/'+url);
so.addVariable('title', 'My Video');
so.addVariable('flashvars','&autostart=true&');
so.write('player');
}
</script>
</body>
</html>