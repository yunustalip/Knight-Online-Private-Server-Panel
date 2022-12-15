<%
Function Connect(Sunucu,VeriTabani,Kullanici,Sifre,SiteyiKapat)
If SiteyiKapat=1 Then
Response.Redirect("../SiteKapali.html")
Response.End
End If
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.open= "driver={SQL Server};server=" & Sunucu & ";database=" & VeriTabani & ";uid=" & kullanici & ";pwd="&Sifre
End Function


Set Conne=Connect(Sunucu,VeriTabani,Kullanici,Sifre,SiteKapali)
%>

