<%%><% 
' NESNELER
' nesne �zellikleri bir bilgi d�nd�r�r.
' nesne.�zellik
' d_diski.diskad� = "aspegitim"
' d_diski.bo�alan
'
' sonuc=nesne.fonksiyon (parametre)
' sonuc=explorer.disk_d.klas�ra� ("asp_icin_acildi")
'metod, fonksiyon demektir.
%>
<%
set FSO = createobject("scripting.filesystemobject")
surucu="d:"
set surucunesne = FSO.getdrive(surucu)
set rootfolder = surucunesne.rootfolder 'd:\
set altklasorler=rootfolder.subfolders
set dosyalar=rootfolder.files
%>
disk s�r�c� : <%= surucunesne.driveletter%> <br />
disk s�r�c� ad�: <%= surucunesne.volumename%> <br />
k�k dizin : <%= rootfolder%> <br />
klas�r say�s� : <%=altklasorler.count %> <br />
dosya say�s� : <%= dosyalar.count%> <br />
<% for each klasor in altklasorler %>
klas�r : <%= klasor%> <br />
klas�r k�sa yol : <%= klasor.shortpath%> <br />
<%next%>

<%for each dosya in dosyalar%>
dosya  : <%= dosya%> <br />
dosya ad� : <%= dosya.name%> <br />
<%next%>
