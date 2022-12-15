<%
Dim ASPHataNesnesi
et ASPHataNesnesi = Server.GetLastError 
%>
<%="ASPCode = " & ASPHataNesnesi.ASPCode & "<br>"%>
<%="ASPDescription = " & ASPHataNesnesi.ASPDescription & "<br>"%>
<%="Category = " & ASPHataNesnesi.Category & "<br>"%>
<%="Column = " & ASPHataNesnesi.Column & "<br>"%>
<%="Description = " & ASPHataNesnesi.Description & "<br>"%>
<%="File = " & ASPHataNesnesi.File & "<br>"%>
<%="Line = " & ASPHataNesnesi.Line & "<br>"%>
<%="Number = " & ASPHataNesnesi.Number & "<br>"%>
<%="Source = " & ASPHataNesnesi.Source & "<br>"%>

