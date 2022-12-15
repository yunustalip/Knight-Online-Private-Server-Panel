<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%Dim MenuAyar,ksira,accid,itemid,islem,sira,itemnum,wishadd,sql,wish,items,itemler,i,pusitem,wishupdate,x
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then
if Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")

elseif Session("login")="ok" and Session("username")<>"" Then
accid=Session("username")

else
server.execute("hata.asp")
Response.End
End If

itemid=secur(Request.Querystring("itemid"))
islem=secur(Request.Querystring("islem"))
sira=secur(Request.Querystring("id"))
itemnum=secur(Request.Querystring("itemnum"))

if not itemid="" Then
Set wishadd = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&accid&"'"
wishadd.open sql,conne,1,3
wishadd("sepet")=wishadd("sepet")+itemid&","
wishadd.update
wishadd.close
set wishadd=nothing
End If
if islem="1" Then
Dim deletewish
Set deletewish = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where straccountid='"&accid&"'"
deletewish.open sql,conne,1,3
deletewish("sepet")=""
deletewish.update
deletewish.close
set deletewish=nothing
End If

	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>

<title>sepet</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-9">
<link href="images/style.css" rel="stylesheet" type="text/css">
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

function buy_go(num) {
	location.href="buy.asp?itemid="+num;
}
function wish_go(num) {
	location.href="cart.asp?itemid="+num;
}
function detail_go(num) {
	location.href="detay.asp?itemid="+num;
}

//-->
</script>
<style type="text/css">
body {background-color: transparent}
</style>
</head><body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="MM_preloadImages('images/butt/butt_buy_on.gif','images/butt/butt_all_cencle_on.gif')">
<table width="288" border="0" cellspacing="0" cellpadding="0">
  <tbody><tr>
    <td height="31">
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

<table style="position: relative; top: -5px;" border="0" cellpadding="0" cellspacing="0">
  <tbody><tr>
        <td width="8"></td>
		<td style="padding-top: 0px;"><a href="cart.asp"><img src="images/right_sub_memu_02b.gif" name="butt_cart" border="0" width="100" height="29"></a></td>
        <td width="1"></td>
		<td style="padding-top: 0px;"><a href="hit.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_detail','','images/right_sub_memu_03b.gif',1)"><img src="images/right_sub_memu_03.gif" name="butt_detail" border="0" width="59" height="29"></a></td>
		<td width="1"></td>

    </tr>
</tbody></table>

    </td>
  </tr>
  <tr>
    <td width="288" height="473" align="center" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" width="230">
        <tbody><tr>
          <td height="27">&nbsp;</td>
        </tr>
        <tr>
         <td class="12_yellow140_shadow" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" align="center">* Favori Listeniz </td>
        </tr>
        <tr>
          <td class="12_white" align="center" valign="bottom"> <div id="item_list" style="border: 0px none ; overflow: auto; height: 400px;">
              <table border="0" cellpadding="0" cellspacing="0" width="212">
              <tbody>
              <%	
set wish=Conne.Execute("select * from tb_user where straccountid='"&accid&"'")
 if islem="2" Then
items=wish("sepet")
itemler=split(items,",")
Set wishupdate = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&accid&"'"
wishupdate.open sql,conne,1,3
wishupdate("sepet")=""
wishupdate.update
wishupdate.close
set wishupdate=nothing
for x=0 to ubound(itemler)-1
if x=cint(sira) Then
itemler(sira)=""
else
dim wishupdt
Set wishupdt = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&accid&"'"
wishupdt.open sql,conne,1,3
wishupdt("sepet")=wishupdt("sepet")+itemler(x)&","
wishupdt.update

End If
next
Response.Redirect ("cart.asp")
End If
			  if not wish.eof  Then
			  items=wish("sepet")
			  itemler=split(items,",")
			  for i=0 to ubound(itemler)
			  if not itemler(i)="" Then
			  set pusitem=Conne.Execute("select * from pus_itemleri where itemkodu='"&itemler(i)&"'")
			  if not pusitem.eof Then%><tr>
                  <td valign="bottom" height="89">
                  <table border="0" cellpadding="0" cellspacing="0" width="212">
                      <tbody><tr>
                        <td width="74"><table border="0" cellpadding="0" cellspacing="0" width="72" height="72">
                            <tbody><tr>
                              <td align="center" background="images/table1_01.gif"><a href="javascript:detail_go(<%=itemler(i)%>);"><img src="item/<%=resim(pusitem("resim"))%>" border="0" width="64" height="64"></a><a href="javascript:detail_go(<%=itemler(i)%>);"></a></td>
                            </tr>
                          </tbody></table></td>
                          <td><a href="cart.asp?islem=2&id=<%=i%>&itemnum=<%=pusitem("itemkodu")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cencle<%=i%>','','images/butt/butt_cart_cencle_on.gif',1)"><img src="images/butt/butt_cart_cencle.gif" name="butt_cencle<%=i%>" id="butt_cencle<%=i%>" border="0" width="22" height="65"></a></td>
                          <td class="12_yellow140_shadow" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;">* <%=pusitem("itemismi")%><br>* Ücret : <%=pusitem("ucret")%><br>* Adet : <%=pusitem("adet")%><br></td>
                      </tr>
                    </tbody></table></td></tr>
<%End If
End If
next
End If %>
                  
								
              </tbody></table>
         <table border="0" cellpadding="0" cellspacing="0" width="96%" height="40">
                <tbody><tr align="center" valign="bottom">
                  <td style="display:{disptemizle}"><a href="cart.asp?islem=1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_all_cencle1','','images/butt/butt_all_cencle_on.gif',1)"><img src="images/butt/butt_all_cencle.gif" name="butt_all_cencle1" id="butt_all_cencle1" border="0" width="89" height="30"></a></td>
                </tr>
              </tbody></table>
            
         
            </div></td>
        </tr>
      </tbody></table>
    </td>
  </tr>
  <tr>
    <td style="padding-top: 10px;" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="252" height="26"><tbody><tr><td style="background-image: url(images/n_table2_02.gif); background-repeat: no-repeat;" height="26">
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr>
  <td style="font-size: 11px; color: rgb(255, 238, 47); font-weight: bold;" align="right"><strong><%=wish("cashpoint")%> P</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr></tbody></table>
</td></tr></tbody></table>
    </td>
  </tr>
  <tr> 
    <td align="center" height="44"> 
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

<table border="0" cellpadding="0" cellspacing="0" width="254">
  <tbody><tr> 
    <td align="center"><a href="my_item.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_item','','images/butt/butt_item_on.gif',1)"><img src="images/butt/butt_item.gif" name="butt_item" border="0" width="119" height="30"></a></td>
    
  </tr>
</tbody></table>

    </td>
  </tr>
</tbody></table>
</body></html>
<%

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>