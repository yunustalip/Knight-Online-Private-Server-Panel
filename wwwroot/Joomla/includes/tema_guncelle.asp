<% 
tema_id = request.Form("tema")
if Session("durum")="giris_yapmis" then
uye_id = Session("uye_id")
baglanti.execute("update gop_uyeler set tema_id='"& tema_id &"' where uye_id= " & uye_id)
response.Redirect "default.asp"
end if
%>