<!-- #include file="VB_Nesne_recordset_Class.asp" -->

<%
Set VeriArama = New VeriOku
VeriArama.Aranacak_Kelime = Request.Form("Arama_Kriter")
VeriArama.Kayit_Bul()
response.write "Okunan Kayýt Sayýsý : " & VeriArama.Kayit_Sayisi & "<br>"
%>
