<%
Function UniqueSecurity(SecurityStatus,IpBlock)
Response.Charset = "iso-8859-9"

If lcase(Request.ServerVariables("SCRIPT_NAME"))="/guvenlik.asp" Then Response.Redirect("Default.Asp"):Response.End


Dim bancook
Dim OrjIP
Dim baglan
Dim floodget
Dim FloodKayit
Dim floodsql

If Session("SimdikiZaman") = Now() And Session("izin")="" Then
Session("FloodSayi") = Session("FloodSayi") + 1

IF Session("FloodSayi")>=IpBlock Then
Set FloodKayit = Server.CreateObject("ADODB.RecordSet")
floodSQL="SELECT * From FloodBan"
FloodKayit.Open floodsql,Conne,1,3
FloodKayit.AddNew
FloodKayit("floodip") = Request.ServerVariables("REMOTE_ADDR")
FloodKayit("floodzaman") = Now()   
FloodKayit.Update
FloodKayit.Close
Set FloodKayit = Nothing
Response.Write ("<base href=""http://"&Request.ServerVariables("Server_Name")&""">")
Response.Write("<br><br><br><br><br><div align=""center""><img border=""0"" src=""imgs/ban.gif""></div>")
Session("FloodSayi")=0
Response.End
End If

Response.Write "<base href=""http://"&Request.ServerVariables("Server_Name")&"""><br><br><br><br><br><div align=""center"" style=""font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif;""><img src=""imgs/Warning2.gif""><br><br>Lütfen Sisteme Flood Yapmayýnýz. Sayfalarý Çok Hýzlý Deðiþtirmek Siteden Çýkarýlmanýza Neden Olabilir."
Response.Write "<br><b>Flood Sayýnýz:"&Session("FloodSayi")&"<br>"&ipblock-Session("FloodSayi")&" Kez Daha Flood Yapmanýz Halinde Siteye Eriþiminiz Engellenecektir.</div>"
Response.End
End If

IF Session("FloodSayi")>=IpBlock Then
Set FloodKayit = Server.CreateObject("ADODB.RecordSet")
floodSQL="SELECT * From FloodBan"
FloodKayit.Open floodsql,Conne,1,3
FloodKayit.AddNew
FloodKayit("floodip") = Request.ServerVariables("REMOTE_ADDR")
FloodKayit("floodzaman") = Now()   
FloodKayit.Update
FloodKayit.Close
Set FloodKayit = Nothing
End If

OrjIP = Request.ServerVariables("REMOTE_ADDR")
Set FloodGet = Conne.Execute("SELECT * From FloodBan Where floodip='"&OrjIP&"'")
IF Not FloodGet.Eof Then
Response.Clear
Response.Write ("<base href=""http://"&Request.ServerVariables("Server_Name")&""">")
Response.Write("<br><br><br><br><br><div align=""center""><img border=""0"" src=""imgs/ban.gif""></div>")
FloodGet.Close
Set FloodGet = Nothing
Session("FloodSayi")=0
Response.End
End If



Session("SimdikiZaman") = Now()

If Session("izin")="yok" Then
Session("izin")=""
End If

End Function

Call UniqueSecurity(1,6)

%>