<% 
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
Dim FSO, Suruculer, surucu
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Suruculer = FSO.Drives
for each Surucu in Suruculer
response.Write "sürücü adý : " & surucu & "<br>"
if (surucu.isready) then
response.Write "sürücü adý : " & surucu.volumename & "<br>"
response.Write "sürücü adý : " & surucu.drivetype & "<br>"
else 
response.Write("sürücü hazýr deðil<br>")
end if
next
%>
<% set surucunesnesi = FSO.Getdrive("a:")
response.Write "a_diskinin adý :" & surucunesnesi.volumename & "<br>"
surucunesnesi.volumename = "disket"
response.Write "a_diskinin adý :" & surucunesnesi.volumename & "<br>"%>
%>