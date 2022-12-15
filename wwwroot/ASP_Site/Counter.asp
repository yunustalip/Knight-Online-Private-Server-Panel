<%

Set Sayac=Server.CreateObject("MSWC.Counters")
Sayac.set "siparisno", 0
Sayac.set "faturano" , 10
sayac.set "ziyaret", 5000

Sayac.Increment  "siparisno"
Sayac.Increment  "siparisno"
Sayac.Increment  "siparisno"
Sayac.Increment  "faturano" 
Sayac.Increment  "faturano" 
sayac.Increment  "ziyaret"

response.Write sayac.get("siparisno")
response.write "<br>"
response.Write sayac.get("faturano")
response.write "<br>"
response.Write sayac.get("ziyaret")
response.write "<br>-------------<br>"
sayac.remove "faturano"
sayac.set "ziyaret", 0

response.Write sayac.get("siparisno")
response.write "<br>"
response.Write sayac.get("faturano")
response.write "<br>"
response.Write sayac.get("ziyaret")
response.write "<br>"

%>