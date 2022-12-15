<%
'______________________________________________
'### BU KISMI DEаноTнRMEYнN

set objayar = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from ayar"
objayar.open SQL,data,1,3
if not objayar.eof then
sitebaslik 	=objayar("sitebaslik")
strsite 	=objayar("site")
adsif 		=objayar("adsif")
adkull 		=objayar("adkull")
aciklama	=objayar("aciklama")
etiket 		=objayar("etiket")
hakkimda	=objayar("hakkimda")
end if
objayar.Close
Set objayar = Nothing
'______________________________________________
%>