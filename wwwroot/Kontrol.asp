<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="guvenlik.asp"-->
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<%
Response.Charset = "iso-8859-9"
Dim id
id=secur(tr(Lcase(Trim(Request.Querystring("username")))))
If id="" Then
Response.Write "<font color=""red""><script>var dbl=document.getElementById('kayitbutton').disabled=true;eval (dbl);var hata = (document.getElementById('usernam').focus()); eval(hata);</script><b>Bo� b�rakmay�n�z.</b></font>"

Else
If Len(id)>20 Then
Response.Write "<font color=""red""><script>var dbl=document.getElementById('kayitbutton').disabled=true;eval (dbl);var hata = (document.getElementById('usernam').focus()); eval(hata);</script><b>Kullan�c� Ad�n�z 20 Karakterden fazla olamaz.</b></font>"
else
Dim Kontrol
Set kontrol=Conne.Execute("select straccountid from tb_user where straccountid='"&id&"'")
If Not kontrol.Eof  Then
Response.Write "<font color=""red""><script>var dbl=document.getElementById('kayitbutton').disabled=true;eval (dbl);var hata = (document.getElementById('usernam').focus()); eval(hata);</script><b>Bu kullan�c� ad� kay�tl�d�r.</b></font>"
Else
Response.Write "<font color=""green""><script>var dbl=document.getElementById('kayitbutton').disabled=false;eval (dbl);</script><b>Ge�erli kullan�c� ad� ("&secur(id)&")</b></font>"
End If
kontrol.Close
Set kontrol=Nothing
End If
End If
%>