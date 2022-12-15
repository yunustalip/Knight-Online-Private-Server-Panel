<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<%
num=request.form("num")
serial=request.form("serial")
dur=request.form("dur")
stacksize=request.form("stacksize")
inventoryslot=request.form("inventoryslot")


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
set userara=Conne.Execute("select StrAccountID from banka_check where StrAccountID='baslangic-item' ")
if userara.eof Then
Conne.Execute("exec banka_baslangicitem_decode")
End If
Conne.Execute("update banka_check set dwid="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and straccountid='baslangic-item'")

elseif Request.Querystring("kyt")="all" Then
set userara=Conne.Execute("select StrAccountID from banka_check where StrAccountID='baslangic-item' ")
if userara.eof Then
Conne.Execute("exec banka_baslangicitem_decode")
End If
Conne.Execute("update banka_check set dwid="&num&",strserial="&serial&",durability="&dur&",stacksize="&stacksize&" where inventoryslot="&inventoryslot&" and straccountid='baslangic-item'")
Conne.Execute("exec banka_baslangicitem_encode ")

End If


End If
%>