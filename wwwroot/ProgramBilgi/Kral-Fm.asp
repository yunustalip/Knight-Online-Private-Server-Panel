<!--#include File="../common.asp"-->
<%Session.CodePage=65001
response.charset="utf-8"

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
Dosyaismi = lcase(Request.ServerVariables("Script_Name"))

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://fmradyodinle.net" or  REFERER_DOMAIN="http://www.fmradyodinle.net" or dosyaismi="/default.asp" or dosyaismi="/404.asp"  Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

If Instr(Request.ServerVariables("ALL_HTTP"),"HTTP_X_REQUESTED_WITH:")>0  or dosyaismi="/default.asp" or dosyaismi="/404.asp" Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

Saat = Hour(Time)
Dakika = Minute(Time)

sSaat=Saat

For sSaat=Saat to Saat-23 step -1
If sSaat<0 Then
aSaat=sSaat+24
Else
aSaat=sSaat
End If

Set DjInf = Adocon.Execute("SELECT * FROM DJBilgi WHERE Radyo_id = 'KralFm' and day='"&Weekday(now)&"' and DATEDIFF(hh, CONVERT(varchar, StartTime, 8), '"&aSaat&":00') = 0")
If Not DjInf.Eof Then

Start = DjInf("StartTime")
Finish = DjInf("EndTime")

psaati=datediff("h",start,finish)

If psaati<0 then
psaati=psaati+24
end if
If not psaati<=Saat-aSaat Then

DjImage=DjInf("DjImage")
ProgramName=djinf("ProgramName") 
DjLink=DjInf("DjLink")
DjName=DjInf("djname")

Start = Left(Start,Len(Start)-3)
Finish = Left(Finish,Len(Finish)-3)
Response.Write "<table ><tr><td rowspan=""8"" valign=""top"">"
If Len(djimage)>0 Then
Response.Write "<img src="""&djimage&""" width=""110""  style=""border: 3px solid rgb(221, 0, 0);"">"
End If
Response.Write "</td></tr>"
If Len(programname)>0 Then
Response.Write "<tr><td class=""Text""><strong>Program: </strong><a href="""&DjLink&""" target=""_blank"">"&ProgramName&"</a></td></tr>"
End If
Response.Write "<tr><td class=""Text""><strong>DJ: </strong>"
If Len(DjLink)>0 Then
Response.Write "<a href="""&DjLink&""" target=""_blank"">"&DjName&"</a>"
Else
Response.Write DjName
End If
Response.Write "<blink><font color=""red""><b> Yayında !</b></font></blink><br>&nbsp;&nbsp;&nbsp;&nbsp;("&Start&" - "&Finish&")"
%>
  <tr>
     <td align="center"><a href="http://www.kralfm.com.tr/yayin_akisi.asp" target="_blank">Yayın Akışı</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.kralfm.com.tr/frekans.asp" target="_blank"><img src="../images/frekans.gif" border="0" width="16" height="16" align="absmiddle"> Frekanslar</a></td>
  </tr>
  <tr>
    <td align="center" class="Text"><strong>Canlı Yayın Tel: </strong>0212 304 0 333</td>
  </tr>
<tr>
    <td align="center" ><strong><a href="ProgramBilgi/KralFmCanliMesaj.asp" target="_blank">Canlı Yayına Mesaj Gönder</strong></td>
  </tr>
<tr>
    <td align="center"><strong><a href="ProgramBilgi/KralFmProgramciMesaj.asp" target="_blank">Programcıya Mesaj Gönder</strong></td>
  </tr>
<%
Response.Write"</table></td>"
Response.Write "  </tr>"
End If

End If

If Not DjInf.Eof Then
Exit For
End If
Next
%>