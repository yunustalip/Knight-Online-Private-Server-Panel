<?xml version="1.0" encoding="iso-8859-9" ?>
<!--#include file="functions/fonksiyonlar.asp"-->
<%
      Response.Buffer = True
      Response.ContentType = "text/xml" 
      Function Temizle(strInput)
          strInput = Replace(strInput,"&", "&amp;", 1, -1, 1)
          strInput = Replace(strInput,"'", "'", 1, -1, 1)
          strInput = Replace(strInput,"""", "", 1, -1, 1)
          strInput = Replace(strInput, ">", "&gt;", 1, -1, 1)
          strInput = Replace(strInput,"<","&lt;", 1, -1, 1)
      Temizle = strInput
      End Function
      %>
      
      <rss version="2.0">
      <channel>
      <title><%= siteadi%></title>
      <link>http://<%= session("siteadres")%></link>
      <description><%= siteadi%></description>
      <language>tr</language>
      <%
      Set rs = Server.CreateObject("Adodb.Recordset")
      SQL = "SELECT * FROM gop_veriler ORDER BY vid DESC"
      rs.Open SQL, Baglanti, 1, 3
	  
	  i = 0
   
      Do While i =< 50 and not rs.EOF 
      response.write("<item>")
      response.write("<title>" & Temizle(rs("vbaslik")) & "</title>")
      response.write("<link>http://"&session("siteadres")&"/default.asp?islem=oku&amp;vid="&rs("vid")&"</link>")
      response.write("<description>" & Temizle(rs("vicerik")) & "</description>")
      response.write("</item>")
	    
		i = i +1
      rs.MoveNext
      Loop
      rs.Close
      set rs = Nothing
      %>
      </channel>
      </rss>

