<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<%
charid=request.form("charid")
num=request.form("num")
dur=request.form("dur")
stacksize=request.form("stacksize")
inventoryslot=request.form("inventoryslot")


if charid="" Then
charid=""
End If
if num="" Then
num="0"
End If

if dur="" Then
dur="0"
End If
if stacksize="" Then
stacksize="0"
End If

if Request.Querystring("islem")="one" Then

set userara=Conne.Execute("select struserid from baslangic_duzenle where struserid='"&charid&"'")
if userara.eof Then
Conne.Execute("exec baslangicitemleri_bul '"&charid&"'")
End If
Conne.Execute("update baslangic_duzenle set dwid="&num&",durability="&dur&",stacksize="&stacksize&" where sira="&inventoryslot&" and struserid='"&charid&"'")

elseif Request.Querystring("islem")="all" Then
if not charid="" Then
set userara=Conne.Execute("select struserid from baslangic_duzenle where struserid='"&charid&"'")
if userara.eof Then
Conne.Execute("exec baslangicitemleri_bul '"&charid&"'")
End If
Conne.Execute("update baslangic_duzenle set dwid="&num&",durability="&dur&",stacksize="&stacksize&" where sira="&inventoryslot&" and struserid='"&charid&"'")
End If
Conne.Execute("exec baslangicitemleri_kaydet '"&charid&"'")
Conne.Execute("delete from baslangic_duzenle where struserid='"&charid&"'")
End If
End If

%>