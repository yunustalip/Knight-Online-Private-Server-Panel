<% 
if Session("durum")="giris_yapmis" then
uye_id=session("uye_id")
midi = guvenlik(request.querystring("mid"))

session("mesajtopla") = ""
session("mesajtopla2") = ""
set mesaj =Baglanti.Execute("Select * from gop_mesajlar where alici= " & uye_id & " and mesaj_id= " & midi )
if mesaj.eof or mesaj.bof then 
Response.Write ""
else

session("mesajtopla") = guvenmesajoku(mesaj("mesaj_icerik")) + session("mesajtopla")
session("mesajtopla2") = mesaj("mesaj_baslik") + session("mesajtopla2")

Response.Write "<b>"&mesaj("mesaj_baslik") & "</b><br>" & session("mesajtopla")
baglanti.Execute("UPDATE gop_mesajlar set mesaj_okundu=1 where mesaj_id=" & midi)

Response.Write "<br><br><br><center><b><a href=""default.asp?islem=bilgi_uye&uye_id="&mesaj("yollayan")&""">"&reply&"</a></b></center>"
end if



else
Response.Write "<center>"&notice4&"</center>"
end if
%>


