<%
Class Linkver
Public link_metin
Public link_url
Public link

sub linkiyaz
link="<a href=" & chr(34) & link_url & chr(34) & ">"&link_metin&"</a>"
response.write link
end sub

End Class
%>

<%
set linkler = new linkver
linkler.link_metin = "Yasal Eğitim Sitesine Gider"
linkler.link_url = "http://www.yasalegitim.com"
linkler.linkiyaz()

%>