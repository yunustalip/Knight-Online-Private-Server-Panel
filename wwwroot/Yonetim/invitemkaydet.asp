<!--#include file="../_inc/conn.asp"-->
<%
if Session("durum")="esp" Then
dim charid,num,serial,dur,stacksize,inventoryslot,userara
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

if Request.Querystring("islem")="one" Then

set userara=Conne.Execute("select struserid from inventory_edit where struserid='"&charid&"'")
if userara.eof Then
Conne.Execute("exec item_decode2 '"&charid&"'")
End If
Conne.Execute("update inventory_edit set num="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and struserid='"&charid&"'")
elseif Request.Querystring("islem")="all" Then

if not charid="" Then
set userara=Conne.Execute("select struserid from inventory_edit where struserid='"&charid&"'")
if userara.eof Then
Conne.Execute("exec item_decode2 '"&charid&"'")
End If
Conne.Execute("update inventory_edit set num="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and struserid='"&charid&"'")
Conne.Execute("exec item_encode2 '"&charid&"'")
End If

End If
End If

%>