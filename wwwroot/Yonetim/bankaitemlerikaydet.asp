<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<%
charid=request.form("charid")
num=request.form("num")
serial=request.form("serial")
dur=request.form("dur")
stacksize=request.form("stacksize")
inventoryslot=request.form("inventoryslot")


if charid="" Then
charid=""
End If
if num="" Then
num="0"
End If
if serial="" Then
serial="0"
End If
if dur="" Then
dur="0"
End If
if stacksize="" Then
stacksize="0"
End If

if Request.Querystring("kyt")="one" Then
if not charid="" Then
set userara=Conne.Execute("select StrAccountID from banka_check where StrAccountID='"&charid&"'")
if userara.eof Then
Conne.Execute("exec banka_item_decode '"&charid&"'")
End If
Conne.Execute("update banka_check set dwid="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and straccountid='"&charid&"'")
End If

elseif Request.Querystring("kyt")="all" Then

if not charid="" Then
set userara=Conne.Execute("select StrAccountID from banka_check where StrAccountID='"&charid&"'")
if userara.eof Then
Conne.Execute("exec banka_item_decode '"&charid&"'")
End If
Conne.Execute("update banka_check set dwid="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and straccountid='"&charid&"'")
Conne.Execute("exec banka_item_encode '"&charid&"'")
Conne.Execute("delete banka_check where straccountid='"&charid&"'")
End If
End If


End If
%>