<div align="center">
<%
If CInt(TopKayit) > CInt(listele) Then 
SayfaSayisi = CInt(TopKayit) / CInt(listele) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a href=?s=1><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a href=?s="&Sayfa-1&">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a href=?s=" & t & "><b>" & t & "</b></a> "
end if
next 
For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s=" & d & "><b>" & d & "</b></a> " 
end if
Next

if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s="&Sayfa+1&"> Sonraki ></a> "
Response.Write " <a href=?s="&SayfaSayisi&">Son >></a>"
end if
end if
end if%>
</div>