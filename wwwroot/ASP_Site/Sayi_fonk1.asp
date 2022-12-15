<%
response.write int(125.8567) & "<br>"
response.write int(-125.4567) & "<br>"
response.write fix(125.8567) & "<br>"
response.write fix(-125.4567) & "<br>"
response.write round(125.8567) & "<br>"
response.write round(125.4567) & "<br>"
response.write round(-125.8567) & "<br>"
response.write round(-125.4567) & "<br>"
response.write round(125.426,2) & "<br>"
response.write round(-125.8525826547,5) & "<br>"
a=fix(125.426) - round(125.426,2)
response.write round(a,2) & "<br>"

response.write sgn(125.8567) & "<br>"
response.write sgn(0) & "<br>"
response.write sgn(-125.8567) & "<br>" 

response.write hex(125) & "<br>"
response.write oct(125) & "<br>"

%>





