<%
dim aylar(12)
aylar(4)="Nisan"
aylar(5)="May�s"
aylar(6)="Haziran"
aylar(7)="Temmuz"
%>
<table border="1"><tr><td><%=aylar(5)%></td></tr></table>
<%
dim mevsimler
mevsimler=array("�lkbahar","Yaz","Sonbahar","K��")
%>
<table border="1"><tr><td><%=mevsimler(0)%></td></tr></table>
<% if isarray(mevsimler) then 
response.write ("dizidir")
end if%>
Alt = <%=Lbound(Aylar)%><br />
Ust = <%=ubound(Aylar)%><br />
<%
for i = Lbound(Aylar) to ubound(Aylar) 
response.Write(Aylar(i)) & "<br>"

next
%>