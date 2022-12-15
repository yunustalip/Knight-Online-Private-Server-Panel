<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<% if Session("durum")="esp" Then
Response.Expires=0
Response.Charset = "iso-8859-9"
Dim charid,inventoryslot,itemozel,itemad,itema
charid=secur(Request.Querystring("charid"))
inventoryslot=Request.Querystring("inventoryslot")

if charid="" or inventoryslot="" Then
Response.End 
End If

set itemozel=Conne.Execute("select * from inventory_edit where inventoryslot="&inventoryslot&" and StrUserId='"&charid&"'")
if not itemozel.eof Then
set itemad=Conne.Execute("select strname from item where num="&itemozel("num")&"")
if not itemad.eof Then
itema=secur(itemad("strname"))
else
itema="&nbsp;"
End If

Response.Write "<form action='javascript:itemkayit();' name=""itemk"" id=""itemk"">"
Response.Write "<div id=""itemmname""><b>"&itema&"</b></div>"
Response.Write "<input type=""hidden"" value="""&charid&""" name=""charid"" >"
Response.Write "<br>Item num: <input type=text value="&itemozel("num")&" name=""num"" id=""num"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"
Response.Write "<br>Item Serial: <input type=text value="""&itemozel("strserial")&""" name=""serial"" id=""serial"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"
Response.Write "<br>Item Durability: <input type=text value="&itemozel("durability")&" name=""dur"" id=""dur"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"
Response.Write "<br>Item Adet: <input type=text value="&itemozel("stacksize")&" name=""stacksize"" id=""stacksize"" onkeyup=""stacksizeupdate('"&inventoryslot&"',this.value)"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"
Response.Write "<br>Item Slot: <input type=text value="&inventoryslot&" name=""inventoryslot"" id=""inventoryslot"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"

Response.Write "<br><br><a href=""#"" onclick=""itemsil('"&inventoryslot&"');return false;"">ITEMI SIL</a></form>"
else
Conne.Execute("exec item_decode2 '"&charid&"'")
Response.Redirect("invitem.asp?charid="&charid&"&inventoryslot="&inventoryslot)
End If
End If%>