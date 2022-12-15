<%
Class Baglanti

Private Sub Class_Initialize
If SiteyiKapat="1" Then
Response.Redirect("../SiteKapali.html")
Response.End
End If
End Sub

Private Sub Class_Terminate

End Sub
Public Function Connect(Sunucu,VeriTabani,Kullanici,Sifre)
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.open= "driver={SQL Server};server=" & Sunucu & ";database=" & VeriTabani & ";uid=" & Kullanici & ";pwd="&Sifre

End Function

End Class
%>