<%
response.write "-----application-----<br>"
for each eleman in application.StaticObjects
response.write eleman & "<br>"
next
response.write "-----session-----<br>"
for each eleman in session.StaticObjects
response.write eleman & "<br>"
next

sayfasayac.pagehit
response.Write sayfasayac.hits
response.write "<br>"
Sayac.set "faturano" , 10
Sayac.increment "faturano"
response.Write sayac.get("faturano")
response.write "<br>"
response.Write "þu anda sitemizde " & application("Sitedeki_Ziyaretci_Sayisi") & " kiþi var."
response.write "<br>"
%>