
<%
'FormatNumber(Deðer,Ondalýk,ön_sýfýr, parantez,Binlik grup)
' -1 : bu ayarý kullan
' 0  : bu ayarý kullanma
' -2 : sistem ayarlarýný kullan
%>
<% sayi = -1256.8426%>
<%= formatnumber(sayi) & "<br>"%>
<%= formatnumber(sayi,3) & "<br>"%>
<%= formatnumber(.75,3,-1) & "<br>"%>
<%= formatnumber(sayi,3,0,-1) & "<br>"%>
<%xx= formatnumber(sayi,3,0,0,-2) %>
firmamýza borcunuz <%=xx%> liradýr.<br />

<%
'FormatCurrency(Deðer,Ondalýk,ön_sýfýr, parantez,Binlik grup)
' -1 : bu ayarý kullan
' 0  : bu ayarý kullanma
' -2 : sistem ayarlarýný kullan
%>
<% sayi = -1256.8426%>
<%= FormatCurrency(sayi) & "<br>"%>
<%= FormatCurrency(sayi,3) & "<br>"%>
<%= FormatCurrency(.75,3,-1) & "<br>"%>
<%= FormatCurrency(sayi,3,0,-1) & "<br>"%>
<%xx= FormatCurrency(sayi,3,0,0,-2) %>
firmamýza borcunuz <%=xx%>.<br />

<%
'FormatPercent(Deðer,Ondalýk,ön_sýfýr, parantez,Binlik grup)
' -1 : bu ayarý kullan
' 0  : bu ayarý kullanma
' -2 : sistem ayarlarýný kullan
%>
<% sayi = 34.678/100%>
<%= FormatPercent(sayi) & "<br>"%>
<%= FormatPercent(sayi,2) & "<br>"%>
<%= FormatPercent(sayi,,-1) & "<br>"%>
<%= FormatPercent(sayi,,,-1) & "<br>"%>
<%xx= FormatPercent(sayi,,,0,-2) %>
string olarak : <%=xx%>.<br />
% <%= formatnumber(100*sayi,2) %>