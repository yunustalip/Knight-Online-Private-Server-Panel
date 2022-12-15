<%
if Session("durum")="giris_yapmis" then
uye_id = Session("uye_id")
midi = guvenlik(request.querystring("mid"))

baglanti.Execute("UPDATE gop_mesajlar set mesaj_gsil=1 where mesaj_id='" & midi & "' and yollayan='"& uye_id &"';")

Response.Write "<br><br><center>"& removed_message &"<br><a href=""" & request.ServerVariables("HTTP_REFERER") & """> "& return &" </a>"

else
Response.Write "<center>"&notice4&"</center>"
end if
%>

