<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<br><img src="imgs/haberler.gif" /><br><br><br><% response.expires=0
Response.Charset = "iso-8859-9"
Dim MenuAyar,ksira,haber,news
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='News'")
If MenuAyar("PSt")=1 Then
link = Session("Sayfa")
gelenlink_bol = split(link, "/")
tp=ubound(gelenlink_bol)

If tp=4 Then
d1=gelenlink_bol(4)
end if
haber=secur(d1)
if haber="" Then %><br>

<table width="500" bgcolor="#e7dbc3">
<tr bgcolor="#25120b">
<td width="328" align="center" bgcolor="#25120b"><span style="color: #FFFFFF">Konu</span></td>
<td width="160" align="center" bgcolor="#25120b"><span style="color: #FFFFFF">Tarih</span></td>
</tr>
<%
Set News=Conne.Execute("select * from haberler")
If not news.eof Then
Do While Not news.eof
Response.Write "<tr ><td><a href=""News/"&news("id")&""" onclick=""pageload('News/"&news("id")&"');return false"" class='link1'>"&news("baslik")&"</a></td><td>"&news("tarih")&"</td></tr>"
News.MoveNext
Loop
End If
News.close
set News=Nothing%>
</table>
<%ElseIf haber<>"" Then
Dim Id
Id=d1
If isnumeric(id)=false Then
Response.End
End If
Dim hbr
Set hbr=Conne.Execute("Select * From haberler where id='"&id&"'")
If Not hbr.eof Then%><br><br>
<table width="557" height="282" cellpadding="0" cellspacing="0" bgcolor="#e7dbc3">
<tr  bgcolor="#25120b" style="color: #FFFFFF; font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; font-style: normal; font-weight: normal; ">
<td width="250" height="42" style="padding-left:20px;color: #FFFFFF;"><% Response.Write "Yazan: "& hbr("gonderen")&"<br>Tarih: "& hbr("tarih")%></td>
<td bgcolor="#25120b" style="color: #FFFFFF;">Konu: 
  <% =hbr("baslik") %></td>
</tr>
<tr>
<td height="150" colspan="2" valign="top" style="padding-left:30px;padding-top:20px;color: #900; font-size: 13px; font-family: Verdana, Arial, Helvetica, sans-serif; font-style: normal; font-weight: normal;"><% =hbr("haber") %></td>
</tr>
<tr>
<td colspan="2" align="right" style="color: #900; font-size: 13px; font-family: Verdana, Arial, Helvetica, sans-serif; font-style: normal; font-weight: normal;">&nbsp;</td>
</tr>
</table>
<%End If
End If
else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>