<% 
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
Dim FSO, Suruculer, surucu
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Suruculer = FSO.Drives
for each Surucu in Suruculer
response.Write "s�r�c� ad� : " & surucu & "<br>"
if (surucu.isready) then
response.Write "s�r�c� ad� : " & surucu.volumename & "<br>"
response.Write "s�r�c� ad� : " & surucu.drivetype & "<br>"
else 
response.Write("s�r�c� haz�r de�il<br>")
end if
next
%>
<% set surucunesnesi = FSO.Getdrive("a:")
response.Write "a_diskinin ad� :" & surucunesnesi.volumename & "<br>"
surucunesnesi.volumename = "disket"
response.Write "a_diskinin ad� :" & surucunesnesi.volumename & "<br>"%>
%>