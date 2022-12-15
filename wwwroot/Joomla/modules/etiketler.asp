<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top">
           <% 
tag=request("tag")
secenek=request("secenek")
deste = 10
if Request("s")="" Then
Sayfa=1
else
Sayfa=Cint(Request("s"))  
End if


if tag="" then
response.write"<center>Geçerli Etikete Týklayýnýz...</center></td></tr></table>"
else
%>
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">        
<%
set sc=server.CreateObject("adodb.recordset") 
  sql =  "select * from gop_veriler where vetiket like '%"&tag&"%' order by vhit desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste 
 
 Set SQLToplam = baglanti.Execute("select count(vid) from gop_veriler where vetiket like '%"&tag&"%'") 
TopKayit = SQLToplam(0)

sc.open sql,baglanti,1,3 



if sc.eof Then
response.write "<td><center><br><br><img src=""img/hata.gif"" /><br><br><b><font color=""red"" size=2>Geçerli Etikete Týklayýnýz...</font></b></center></td><br>"
end if 

if Not sc.eof Then 
For zayfa= 1 to 999 
 %>
          <tr>
            <td width="200"><span>
<%
Response.Write "<b> <a href="""&session("siteadres")&"?islem=oku&vid="&sc("vid")&""">"&Replace(sc("vbaslik"),""&tag&"","<font style=""background-color:#FFCC00"" color=""#000000"">"&tag&"</font>")&"</a></b><br>" 
Response.Write Replace(Left(sc("vicerik"),150),""&tag&"","<font style=""background-color:#FFCC00"" color=""#000000"">"&tag&"</font>")&"<br>"
Response.Write "<i><a href="""&session("siteadres")&"/?islem=oku&vid="&sc("vid")&""">"&session("siteadres")&"/?islem=oku&vid="&sc("vid")&"</i>"

%>
            </span></td>
            </tr>
          <tr>
            <td height="25"><hr size="1"></td>
          </tr>
                    <%
sc.MoveNext
if sc.eof Then exit For
Next
end if
%>
        </table>
            <div align="center">
<%

Set rs = baglanti.Execute("SELECT COUNT(vid) AS toplam FROM gop_veriler where vetiket like '%"&tag&"%'")
%>

   </div></td>
  </tr>
      </table>
<div align="center">
<%

If CInt(TopKayit) > CInt(deste) Then 
SayfaSayisi = CInt(TopKayit) / CInt(deste) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a class=pagenav href=""?islem=etiket&tag="&tag&"&s=1""><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a class=pagenav href=""?islem=etiket&tag="&tag&"&s="&Sayfa-1&""">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a class=pagenav href=""?islem=etiket&tag="&tag&"&s=" & t & """><b>" & t & "</b></a> "
end if
next 

For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span class=pagenav><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href=""?islem=etiket&tag="&tag&"&s=" & d & """><b>" & d & "</b></a> " 
end if
Next


if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href=""?islem=etiket&tag="&tag&"&s="&Sayfa+1&"""> Sonraki ></a> "
Response.Write " <a class=pagenav href=""?islem=etiket&tag="&tag&"&s="&SayfaSayisi&""">Son >></a>"
end if


End If 
End If 

%>

</div>
<%
sc.close
Set sc=Nothing
end if
%>