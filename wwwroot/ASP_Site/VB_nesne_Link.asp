
<%
Class Linkver
Public link_metin
Public link_url
Public link

sub linkiyaz
link="<a href=" & chr(34) & link_url & chr(34) & ">"&link_metin&"</a>"
response.write link
end sub

sub linkiyazdir(parmetin,parlink)
link="<a href=" & chr(34) & parlink & chr(34) & ">"&parmetin&"</a>"
response.write link
end sub

End Class
%>

<%
set linkler = new linkver
linkler.link_metin = "Yasal E�itim Sitesine Gider"
linkler.link_url = "http://www.yasalegitim.com"
linkler.linkiyaz() : response.write "<br>"

linkler.linkiyazdir "Yasal E�itim","http://www.yasalegitim.com": response.write "<br>"

%>

Firmam�z <%linkler.linkiyazdir "Yasal E�itim","http://www.yasalegitim.com"%>, G�rsel ve ��itsel E�itim CD'leri �retmektedir.
erwerwe wer wer 
wer<%=linkler.link%>rwer
wer
we
rewrw werwer <%=linkler.link%>erwer rewrwer
wewerwer
<%=linkler.link%>wewr<%=linkler.link%>wer
ewerrwer