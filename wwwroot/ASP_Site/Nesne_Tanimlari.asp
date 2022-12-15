<%
Class Linkver
Public link_metin
Public link_url
Public link

sub linkiyazdir(parmetin,parlink)
link="<a href=" & chr(34) & parlink & chr(34) & ">"&parmetin&"</a>"
response.write link
end sub

End Class
%>

<%
sub sublink(parmetin,parlink)
response.write "<a href=" & chr(34) & parlink & chr(34) & ">"&parmetin&"</a>"

end sub
%>