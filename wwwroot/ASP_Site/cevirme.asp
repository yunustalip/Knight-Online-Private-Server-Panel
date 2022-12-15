<%

response.write CBool(x) & "<br>"
response.write CByte(224) & "<br>"
response.write CCur(5568.89) & "<br>"
response.write CCur(25640.878555649) & "<br>"
response.write typename(CCur(12500)) & "<br>"%>
<%= CDate("12/1/2006") %>
<%= typename("12/1/2006") %>
<%= typename(CDate("12/1/2006")) & "<br>"%>
<%= CDate("12 ocak 2006") %>
<%= CDate("12:12 pm")  & "<br>"%>

<%= CDbl(158.78)  & "<br>"%>
<%= Csng(999999999999999999)  & "<br>"%>
<%= Cint(158.78)  & "<br>"%>
<%= CLng(158.78)  & "<br>"%>
<%= CStr(158.78)  & "<br>"%>
<%= typename(CStr(158.78))  & "<br>"%>

