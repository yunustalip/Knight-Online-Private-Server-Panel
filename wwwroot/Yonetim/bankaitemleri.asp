<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<% Response.Charset = "iso-8859-9"
charid=secur(Request.Querystring("charid"))
inventoryslot=Request.Querystring("inventoryslot")

if isnull(num)=true Then
num="0"
End If


set itemozel=Conne.Execute("select * from banka_check where inventoryslot="&inventoryslot&" and straccountid='"&charid&"'")


if not itemozel.eof Then
dwid=itemozel("dwid")
strserial=itemozel("strserial")
durability=itemozel("durability")
adet=itemozel("stacksize")

if isnull(itemozel("dwid"))=true Then
dwid=0
End If
if isnull(itemozel("strserial")) Then
strserial=0
End If
if isnull(itemozel("durability")) Then
durability=0
End If
if isnull(itemozel("stacksize")) Then
adet=0
End If
set itemad=Conne.Execute("select strname from item where num="&dwid&"")
if not itemad.eof Then
itema=secur(itemad("strname"))
else
itema=""
End If

Response.Write "<form action='javascript:itemkayt();' name=""itemk"" id=""itemk"" onchange=""itemkayt();"">"
Response.Write "<b>"&itema
Response.Write "<input type=""hidden"" value="""&charid&""" name=""charid"" >"
Response.Write "<br>Item num: <input type=text value="&dwid&" name=""num"" id=""num"" >"
Response.Write "<br>Item Serial: <input type=text value="&strserial&" name=""serial"" id=""serial"" >"
Response.Write "<br>Item Durability: <input type=text value="&durability&" name=""dur"" id=""dur"" >"
Response.Write "<br>Item Adet: <input type=text value="&adet&" name=""stacksize"" id=""stacksize"" >"
Response.Write "<br>Item Slot: <input type=text value="&inventoryslot&" name=""inventoryslot"" id=""inventoryslot"" >"
Response.Write "<br><input type=""submit"" value=""KAYDET"" id=""formbtn"">"
Response.Write "<br><a href=""javascript:itemsil('"&inventoryslot&"')"">ITEMI SIL</a></form>"
else
Conne.Execute("exec banka_item_decode '"&charid&"'")
Response.Redirect("bankaitemleri.asp?is=sil&charid="&charid&"&num="&num&"&serial="&serial&"&dur="&dur&"&stacksize="&stacksize&"&inventoryslot="&inventoryslot)
End If
End If%>