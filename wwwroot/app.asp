<%
dim i
For Each i in Application.Contents
  Response.Write(i & "<br>")
Next
%>