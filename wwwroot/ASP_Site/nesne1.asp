<%%><% 
' NESNELER
' nesne özellikleri bir bilgi döndürür.
' nesne.özellik
' d_diski.diskadý = "aspegitim"
' d_diski.boþalan
'
' sonuc=nesne.fonksiyon (parametre)
' sonuc=explorer.disk_d.klasöraç ("asp_icin_acildi")
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
disk sürücü : <%= surucunesne.driveletter%> <br />
disk sürücü adý: <%= surucunesne.volumename%> <br />
kök dizin : <%= rootfolder%> <br />
klasör sayýsý : <%=altklasorler.count %> <br />
dosya sayýsý : <%= dosyalar.count%> <br />
<% for each klasor in altklasorler %>
klasör : <%= klasor%> <br />
klasör kýsa yol : <%= klasor.shortpath%> <br />
<%next%>

<%for each dosya in dosyalar%>
dosya  : <%= dosya%> <br />
dosya adý : <%= dosya.name%> <br />
<%next%>
