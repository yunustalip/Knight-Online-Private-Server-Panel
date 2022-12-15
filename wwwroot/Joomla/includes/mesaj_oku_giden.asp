<% 
if Session("durum")="giris_yapmis" then
uid=session("uye_id")
midi = guvenlik(request.querystring("mid"))


set mesaj =Baglanti.Execute("Select * from gop_mesajlar where yollayan= "&uid&" and mesaj_id= "&midi&";" )
if mesaj.eof or mesaj.bof then 
Response.Write ""
else


Response.Write "<b>"&mesaj("mesaj_baslik") & "</b><br>" & guvenmesajoku(mesaj("mesaj_icerik"))

end if



else
Response.Write "<center>"&notice4&"</center>"
end if
%>


