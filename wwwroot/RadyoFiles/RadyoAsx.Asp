<!--#include file="../_inc/conn.asp"-->
<%
if mid(Request.ServerVariables("HTTP_USER_AGENT"),1,20)<>"Windows-Media-Player" Then
Response.Redirect("../")
Response.End
End If
dim REFERER_URL,REFERER_DOMAIN
REFERER_URL = Request.ServerVariables("HTTP_REFERER")

if InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
else
REFERER_DOMAIN = left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If

if REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then

else
Response.Redirect("../")
End If

Response.expires=0 

function veritemizle(ayikla)
ayikla = Replace (ayikla ,"`","")
ayikla = Replace (ayikla ,"'","")
ayikla = Replace (ayikla ,"=","")
ayikla = Replace (ayikla ,"&","")
ayikla = Replace (ayikla ,"%","")
ayikla = Replace (ayikla ,"!","")
ayikla = Replace (ayikla ,"#","")
ayikla = Replace (ayikla ,"<","")
ayikla = Replace (ayikla ,">","")
ayikla = Replace (ayikla ,"*","")
ayikla = Replace (ayikla ,"And","")
ayikla = Replace (ayikla ,"’","")
ayikla = Replace (ayikla ,"Chr(34)","")
ayikla = Replace (ayikla ,"Chr(39)","")
veritemizle = ayikla
End Function
Dim radyoid,con
radyoid=veritemizle(Request.Querystring("radyoid"))
set con=Conne.execute("Select * from RadyoList where id='"&radyoid&"'")
If Not con.Eof Then%>
<ASX version = "3.0">
<repeat>
   <Entry>  
<title><%=con("radyo")%> - Radyolarim.Net</title>
      <Ref href = "<%=con("url")%>" />
   </Entry>
</repeat>
</ASX>
<%else
response.clear
Response.Write "Radyo Bulunamadý!"
End If%>