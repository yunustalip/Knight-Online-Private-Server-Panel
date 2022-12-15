<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr align="left">
                  <td width="45" height="25"></td>
    <td><span><a href="?islem=uyeler&amp;uye_adi=asc"><img src="images/asagi.gif" width="13" height="13" border="0"></a> <strong><%= username %></strong> <a href="?islem=uyeler&amp;uye_adi=desc"><img src="images/yukari.gif" width="13" height="13" border="0"></a></span></td>
    <td width="150" align="center"><span><a href="?islem=uyeler&amp;uye_kayit=asc"><img src="images/asagi.gif" width="13" height="13" border="0"></a> <strong><%= register_date %></strong> <a href="?islem=uyeler&amp;uye_kayit=desc"><img src="images/yukari.gif" width="13" height="13" border="0"></a></span></td>
    <td width="150" align="center"><span><a href="?islem=uyeler&amp;uye_son_giris=asc"><img src="images/asagi.gif" width="13" height="13" border="0"></a> <strong><%= last_login %></strong> <a href="?islem=uyeler&amp;uye_son_giris=desc"><img src="images/yukari.gif" width="13" height="13" border="0"></a></span></td>
    <td width="75" align="center">&nbsp;</td>
  </tr>
<%

deste = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

tarih2 = Year(date)&"-"&Month(date)&"-"&Day(date)&" "&Hour(now)&":"&Minute(now)-1&":"&second(now)

if Request.QueryString("uye_son_giris") = "desc" then
set rs = baglanti.Execute("select * from gop_uyeler order by uye_son_tarih desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
session("sayfam")="uye_giris_son=desc"
elseif Request.QueryString("uye_son_giris") = "asc" then
session("sayfam")="uye_giris_son=asc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_son_tarih asc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
elseif Request.QueryString("uye_kayit") = "desc" then
session("sayfam")="uye_kayit=desc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_tarih desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
elseif Request.QueryString("uye_kayit") = "asc" then
session("sayfam")="uye_kayit=asc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_tarih asc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
elseif Request.QueryString("uye_adi") = "asc" then
session("sayfam")="uye_adi=asc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_adi asc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
elseif Request.QueryString("uye_adi") = "desc" then
session("sayfam")="uye_adi=desc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_adi desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
elseif Request.QueryString("durum") = "on" then
session("sayfam")="durum=on"
set rs = baglanti.Execute("Select * from gop_uyeler where uye_son_tarih >= '"&tarih2&"';")
elseif Request.QueryString("durum") = "off" then
session("sayfam")="durum=off"
set rs = baglanti.Execute("Select * from gop_uyeler where uye_son_tarih < '"&tarih2&"' LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
else
session("sayfam")="uye_adi=asc"
set rs = baglanti.Execute("select * from gop_uyeler order by uye_adi asc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)
end if

sayfam = "?islem=uyeler&"&session("sayfam")

Set SQLToplam = baglanti.Execute("select count(uye_id) from gop_uyeler") 
TopKayit = SQLToplam(0)


if rs.eof or rs.bof then
response.Write ""
else
for k=1 to CInt(TopKayit)
if rs.eof then exit for
%>
<tr align="left" bgcolor="#ffffff" onmouseover="this.className='tablorenk1';" onmouseout="this.className='tablorenk2';">
                  <td align="center"><strong><%= deste*Sayfa+k-deste%></strong></td>
                  <td ><a href="?islem=bilgi_uye&uye_id=<%= rs("uye_id")%>"><%=rs("uye_adi")%></a></td>
                  <td align="center"><%=rs("uye_tarih") %></td>
                  <td align="center"><%=rs("uye_son_tarih") %></td>
                  <td align="center"><% 
Set online = baglanti.Execute("Select * from gop_uyeler where uye_id='"& rs("uye_id") &"' and uye_son_tarih >= '"&tarih2&"' ;") 
if online.eof or online.bof then
Response.Write "<img src=""images/pasif.png"" alt=""On-line"">"
else
Response.write "<img src=""images/aktif.png"" alt=""Off-line"">"
online.close
		end if%></td>
  </tr>
                <%
rs.movenext
next
%>
              </table>
             
<div align="center">
<%
if Request.QueryString("durum") = "on" then
Response.Write ""
else


If CInt(TopKayit) > CInt(deste) Then 
SayfaSayisi = CInt(TopKayit) / CInt(deste) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a class=pagenav href="&sayfam&"&s=1><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a class=pagenav href="&sayfam&"&s="&Sayfa-1&">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a class=pagenav href="&sayfam&"&s=" & t & "><b>" & t & "</b></a> "
end if
next 

For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span class=pagenav><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href="&sayfam&"&s=" & d & "><b>" & d & "</b></a> " 
end if
Next


if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href="&sayfam&"&s="&Sayfa+1&"> Sonraki ></a> "
Response.Write " <a class=pagenav href="&sayfam&"&s="&SayfaSayisi&">Son >></a>"
end if


End If 
End If 

end if
%>

</div>
<%
rs.close
set rs = nothing
end if
%>