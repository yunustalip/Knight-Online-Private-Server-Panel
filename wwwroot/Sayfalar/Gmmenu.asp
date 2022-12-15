<%Response.expires=0
response.charset="iso-8859-9"
if Session("yetki")="1" Then %>
<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<script type="text/javascript">
function selectCode(a)
{
   var e = a;
   if (window.getSelection)
   {
      var s = window.getSelection();
       if (s.setBaseAndExtent)
      {
         s.setBaseAndExtent(e, 0, e, e.innerText.length - 1);
      }
      else
      {
         var r = document.createRange();
         r.selectNodeContents(e);
         s.removeAllRanges();
         s.addRange(r);
      }
   }
   else if (document.getSelection)
   {
      var s = document.getSelection();
      var r = document.createRange();
      r.selectNodeContents(e);
      s.removeAllRanges();
      s.addRange(r);
   }
   else if (document.selection)
   {
      var r = document.body.createTextRange();
      r.moveToElementText(e);
      r.select();
   }
}
</script>
<br><br><br />
<br />
<br><div align="left" style="position:relative;left:30px">
<b><u>Gmlerin Site Üzerinde Kullanabileceði Komutlar</u></b><br>
<%Set gmytk=Conne.Execute("select gmyetki from siteayar")
yetkiler=Split(gmytk(0),vBcrlf)
For x=0 To UBound(yetkiler)
Response.Write "<a class=""link1"" href=""javascript:komutgir('"&yetkiler(x)&"')"">"&yetkiler(x)&"</a>&nbsp;|&nbsp;"
If x Mod 5=4 Then
Response.Write "<br>"
End If
Next%></div>
<div style="font-size:12px;font-weight:bold">Oyuncularýn Program Listesi</div>
<%Set programs=Conne.Execute("select * from myst_check order by updatetime")
If Not programs.eof Then
Response.Write "<br><table width=""600"" border=""1"" cellspacing=""3"" cellpadding=""3""><tr style=""font-weight:bold""><td width=""70"">Hesap Adý</td><td width=""80"">Karakter Adý</td><td>Programlar</td><td>Tarih</td></tr>"
Do While Not programs.Eof
Response.Write "<tr><td valign=""top"">"&programs("straccountid")&"</td><td valign=""top"">"&programs("struserid")&"</td>"
Response.Write "<td>"&Replace(programs("strprogram"),"|","<br>")&"</td>"
Response.Write "<td valign=""top"">"&programs("updatetime")&"</td></tr>"
programs.Movenext
Loop
Response.Write "</table>"
Else
Response.Write "<br><b>Kayýt Bulunamadý.</b>"
End If
End If %>