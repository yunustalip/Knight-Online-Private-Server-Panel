<%

On Error Resume Next
Err.Raise 55
Err.number = 8000
Err.description = "�zel hata var"
hatanumarasi = Err.number
hataaciklama = Err.description
If hatanumarasi <> 0 Then
 Response.Write "Hata var! Hata Numaras� ve a��klamas� : " 
 Response.write hatanumarasi & " - " & hataaciklama
End If 
%>
