<%
Class Linkver
Public link_metin
Public link_url
Public link

sub linkiyaz(a,b)
link="<a href=" & chr(34) & a& chr(34) & ">"&b&"</a>"
response.write link
end sub

End Class
%>

<%
set linkler = new linkver
linkler.link_metin = "Yasal Eðitim Sitesine Gider"
linkler.link_url = "http://www.yasalegitim.com"
linkler.linkiyaz "Yasal Eðitim Sitesine Gider" , "http://www.yasalegitim.com"

%>