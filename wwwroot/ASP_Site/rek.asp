<%

Set reklamlar=Server.CreateObject("MSWC.Adrotator")
response.write reklamlar.GetAdvertisement("adrotator.txt")

%>
