
<%
'FormatNumber(De�er,Ondal�k,�n_s�f�r, parantez,Binlik grup)
' -1 : bu ayar� kullan
' 0  : bu ayar� kullanma
' -2 : sistem ayarlar�n� kullan
%>
<% sayi = -1256.8426%>
<%= formatnumber(sayi) & "<br>"%>
<%= formatnumber(sayi,3) & "<br>"%>
<%= formatnumber(.75,3,-1) & "<br>"%>
<%= formatnumber(sayi,3,0,-1) & "<br>"%>
<%xx= formatnumber(sayi,3,0,0,-2) %>
firmam�za borcunuz <%=xx%> lirad�r.<br />

<%
'FormatCurrency(De�er,Ondal�k,�n_s�f�r, parantez,Binlik grup)
' -1 : bu ayar� kullan
' 0  : bu ayar� kullanma
' -2 : sistem ayarlar�n� kullan
%>
<% sayi = -1256.8426%>
<%= FormatCurrency(sayi) & "<br>"%>
<%= FormatCurrency(sayi,3) & "<br>"%>
<%= FormatCurrency(.75,3,-1) & "<br>"%>
<%= FormatCurrency(sayi,3,0,-1) & "<br>"%>
<%xx= FormatCurrency(sayi,3,0,0,-2) %>
firmam�za borcunuz <%=xx%>.<br />

<%
'FormatPercent(De�er,Ondal�k,�n_s�f�r, parantez,Binlik grup)
' -1 : bu ayar� kullan
' 0  : bu ayar� kullanma
' -2 : sistem ayarlar�n� kullan
%>
<% sayi = 34.678/100%>
<%= FormatPercent(sayi) & "<br>"%>
<%= FormatPercent(sayi,2) & "<br>"%>
<%= FormatPercent(sayi,,-1) & "<br>"%>
<%= FormatPercent(sayi,,,-1) & "<br>"%>
<%xx= FormatPercent(sayi,,,0,-2) %>
string olarak : <%=xx%>.<br />
% <%= formatnumber(100*sayi,2) %>