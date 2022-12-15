<% response.Buffer=true

x=0
Do 
    x = x+1
    Response.Write x & "<br>"
Loop until x = 10
response.flush
y=0
Do 
    y = y+1
    Response.Write y & "<br>"
Loop until y = 10
'response.clear
%>