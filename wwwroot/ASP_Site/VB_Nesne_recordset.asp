<!-- #include file="VB_Nesne_recordset_Class.asp" -->

<%
Set VeriArama = New VeriOku
VeriArama.Aranacak_Kelime = Request.Form("Arama_Kriter")
VeriArama.Kayit_Bul()
response.write "Okunan Kay�t Say�s� : " & VeriArama.Kayit_Sayisi & "<br>"
%>
