<% 
if Session("durum")="giris_yapmis" then
uye_id = Session("uye_id")

deste = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1
End If 
%>


<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="50%" height="25"><div align="center" class="current" id="t11"><a href="#" onclick="showtab(1,1); return false;"><h3><%= inbox %></h3></a></div></td>
    <td><div align="center" id="t12"><a href="#" onclick="showtab(1,2); return false;"><h3><%= sent_items %></h3></a></div></td>
  </tr>
</table>

<div id="tab11">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20"><strong>&nbsp;<%= heading %></strong></td>
    <td width="200" align="center"><strong><%= recipients %></strong></td>
    <td width="150" align="center"><strong><%= date_in %></strong></td>
    <td width="50" align="center"><strong><%= deleting %></strong></td>
  </tr>
  <%
  
deste = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1
End If 

Set topla = baglanti.Execute("select count(mesaj_id) AS toplam from gop_mesajlar where alici='"& uye_id &"' and mesaj_sil=0")
TopKayit = topla(0)

set mesaj =Baglanti.Execute("Select * from gop_mesajlar where alici= "& uye_id &" and mesaj_sil=0 order by mesaj_id desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)

if mesaj.eof or mesaj.bof then
Response.Write "  <tr>    <td colspan=""4"" align=""center"">"&no_message&"</td>  </tr>"
else
for k=1 to "25"
if mesaj.eof then exit for

if mesaj("mesaj_okundu") = 0 then
 %>
<tr bgcolor="#ffffff" onmouseover="this.className='tablorenk4';" onmouseout="this.className='tablorenk3';" class="tablorenk3">
<% else %>
<tr bgcolor="#ffffff" onmouseover="this.className='tablorenk1';" onmouseout="this.className='tablorenk2';">
<% end if %>
<td height="20">
&nbsp;
<%
if mesaj("mesaj_okundu") = 0 then

if mesaj("mesaj_baslik") = "" then
Response.Write "<a href=""default.asp?islem=mesaj_oku&mid="&mesaj("mesaj_id")&""">"&unsubjected&"</a> ("&unread&")"
else
Response.Write "<a href=""default.asp?islem=mesaj_oku&mid="&mesaj("mesaj_id")&""">"& mesaj("mesaj_baslik") &"</a> ("&unread&")" 
end if

else

if mesaj("mesaj_baslik") = "" then
Response.Write "<a href=""default.asp?islem=mesaj_oku&mid="&mesaj("mesaj_id")&""">"&unsubjected&"</a>"
else
Response.Write "<a href=""default.asp?islem=mesaj_oku&mid="&mesaj("mesaj_id")&""">"& mesaj("mesaj_baslik") &"</a>" 
end if

end if

%></td>
    <td><div align="center">
<%
set uyebul = baglanti.execute("select * from gop_uyeler where uye_id = "& mesaj("yollayan") &";")
if uyebul.eof or uyebul.bof then
Response.Write ""
else
Response.Write "<a href=""default.asp?islem=bilgi_uye&uye_id="&mesaj("yollayan")&""">"& uyebul("uye_adi") & "</a>"
end if

%></div></td>
    <td><%= mesaj("mesaj_tarih") %></td>
    <td><div align="center"><a href="default.asp?islem=mesaj_sil&mid=<%= mesaj("mesaj_id")%>"><%= deleting %></a></div></td>
  </tr>
  <%
mesaj.movenext
next
mesaj.close
set mesaj = nothing
end if %>
</table>
<div align="center">

<%
If CInt(TopKayit) > CInt(deste) Then 
SayfaSayisi = CInt(TopKayit) / CInt(deste) 

If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a class=""pagenav"" href=default.asp?islem=mesajlarim&s=1><< "&start_page&"</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a class=""pagenav"" href=default.asp?islem=mesajlarim&s="&Sayfa-1&">< "&prev_page&"</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-10 and t > Sayfa-10  and t > 0 then
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim&s=" & t & "><b>" & t & "</b></a> "
end if
next 

For d=Sayfa To Sayfa+9
if d = CInt(Sayfa) then
Response.Write "<span class=""pagenav""><b>" & d & "</b></span>"
elseif d > CInt(SayfaSayisi+1) then
Response.Write ""
else
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim&s=" & d & "><b>" & d & "</b></a> " 
end if
Next


if Sayfa = CInt(SayfaSayisi+1) then
Response.Write ""
else
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim&s="&Sayfa+1&"> "&next_page&" ></a> "
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim&s="&CInt(SayfaSayisi+1)&">"&end_page&" >></a>"
end if


End If 
End If 
%></div></div>

<div id="tab12" style="display: none;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20"><strong>&nbsp;<%= heading %></strong></td>
    <td width="200" align="center"><strong><%= recipients %></strong></td>
    <td width="150" align="center"><strong><%= date_in %></strong></td>
    <td width="50" align="center"><strong><%= deleting %></strong></td>
  </tr>
  <%
  
deste = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1
End If 

Set topla = baglanti.Execute("select count(mesaj_id) AS toplam from gop_mesajlar where yollayan='"& uye_id &"' and mesaj_gsil=0")
TopKayit = topla(0)

set mesaj =Baglanti.Execute("Select * from gop_mesajlar where yollayan= "& uye_id &" and mesaj_gsil=0 order by mesaj_id desc LIMIT "& (deste*Sayfa)-(deste) & "," & deste)

if mesaj.eof or mesaj.bof then
Response.Write "  <tr>    <td colspan=""4"" align=""center"">"&no_message&"</td>  </tr>"
else
for k=1 to "25"
if mesaj.eof then exit for

if mesaj("mesaj_okundu") = 0 then
 %>
<tr bgcolor="#ffffff" onmouseover="this.className='tablorenk4';" onmouseout="this.className='tablorenk3';" class="tablorenk3">
<% else %>
<tr bgcolor="#ffffff" onmouseover="this.className='tablorenk1';" onmouseout="this.className='tablorenk2';">
<% end if %>
<td height="20">
&nbsp;
<%
if mesaj("mesaj_okundu") = 0 then

if mesaj("mesaj_baslik") = "" then
Response.Write "<a href=""default.asp?islem=mesaj_oku_giden&mid="&mesaj("mesaj_id")&""">"&unsubjected&"</a> ("&unread&")"
else
Response.Write "<a href=""default.asp?islem=mesaj_oku_giden&mid="&mesaj("mesaj_id")&""">"& mesaj("mesaj_baslik") &"</a> ("&unread&")" 
end if

else

if mesaj("mesaj_baslik") = "" then
Response.Write "<a href=""default.asp?islem=mesaj_oku_giden&mid="&mesaj("mesaj_id")&""">"&unsubjected&"</a>"
else
Response.Write "<a href=""default.asp?islem=mesaj_oku_giden&mid="&mesaj("mesaj_id")&""">"& mesaj("mesaj_baslik") &"</a>" 
end if

end if

%></td>
    <td><div align="center">
<%
set uyebul = baglanti.execute("select * from gop_uyeler where uye_id = "& mesaj("yollayan") &";")
if uyebul.eof or uyebul.bof then
Response.Write ""
else
Response.Write "<a href=""default.asp?islem=bilgi_uye&uye_id="&mesaj("yollayan")&""">"& uyebul("uye_adi") & "</a>"
end if

%></div></td>
    <td><%= mesaj("mesaj_tarih") %></td>
    <td><div align="center"><a href="default.asp?islem=mesaj_giden_sil&mid=<%= mesaj("mesaj_id")%>"><%= deleting %></a></div></td>
  </tr>
  <%
mesaj.movenext
next
mesaj.close
set mesaj = nothing
end if %>
</table>
<div align="center">

<%
If CInt(TopKayit) > CInt(deste) Then 
SayfaSayisi = CInt(TopKayit) / CInt(deste) 

If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s=1><< "&start_page&"</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s="&Sayfa-1&">< "&prev_page&"</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-10 and t > Sayfa-10  and t > 0 then
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s=" & t & "><b>" & t & "</b></a> "
end if
next 

For d=Sayfa To Sayfa+9
if d = CInt(Sayfa) then
Response.Write "<span class=""pagenav""><b>" & d & "</b></span>"
elseif d > CInt(SayfaSayisi+1) then
Response.Write ""
else
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s=" & d & "><b>" & d & "</b></a> " 
end if
Next


if Sayfa = CInt(SayfaSayisi+1) then
Response.Write ""
else
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s="&Sayfa+1&"> "&next_page&" ></a> "
Response.Write " <a class=""pagenav"" href=default.asp?islem=mesajlarim_giden&s="&CInt(SayfaSayisi+1)&">"&end_page&" >></a>"
end if


End If 
End If 
%></div>
</div>


<%
else
Response.Write "<center>"&notice4&"</center>"
end if
 %>