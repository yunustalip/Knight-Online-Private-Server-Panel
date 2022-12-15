<!--#include file="DbSettings.Asp"-->
<%
Class Connection

Private Sub Class_Initialize
If SiteyiKapat="1" Then
Response.Redirect("../SiteKapali.html")
Response.End
End If
End Sub

Private Sub Class_Terminate
Connect.Close
Set Connect=Nothing
End Sub

Public Function Connect
On Error Resume Next

Set Connect = Server.CreateObject("ADODB.Connection")
Connect.open= "driver={SQL Server};server=" & Sunucu & ";database=" & VeriTabani & ";uid=" & kullanici & ";pwd="&Sifre

If Err.Number<>0 then
Response.Clear
Response.Write "Hata Oluþtu ! Veri Tabaný Bilgileri Yanlýþ Olabilir. Lütfen Kontrol Ediniz."
Response.End
End If

End Function

End Class

Set Bag = New Connection
Set Conne = Bag.Connect

%>