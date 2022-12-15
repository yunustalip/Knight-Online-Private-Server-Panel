<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->

<% Response.Charset = "iso-8859-9"
ssid=secur(Request.Querystring("ssid"))
slot=secur(Request.Querystring("slot"))
num=secur(Request.Querystring("itemno"))


set itemozel=Conne.Execute("select * from k_monster_item where sindex="&ssid)
if not itemozel.eof Then
set itemad=Conne.Execute("select strname from item where num="&itemozel("iItem0"&slot&"")&"")


Response.Write "<form action='javascript:itemkayt();' name=""itemk"" id=""itemk"">"
Response.Write "<br>Ýtem No: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""text"" value="""&itemozel("iItem0"&slot&"")&""" onblur=""itemkayit()"" name=""num"" id=""num"" >"
Response.Write "<br>Drop Yüzdesi: <input type=""text"" value="""&itemozel("sPersent0"&slot&"")/100&""" onblur=""itemkayit()"" name=""dropyuzde"" id=""dropyuzde"" >"
Response.Write "<br><input type=""hidden"" value="&slot&" name=""inventoryslot"" id=""inventoryslot"" >"
Response.Write "<br><input type=""hidden"" value="&ssid&" name=""ssid"" id=""ssid"" >"
Response.Write "<br><a href=""javascript:itemsil('"&inventoryslot&"')"">ITEMI SIL</a></form>"
else
End If
End If%>