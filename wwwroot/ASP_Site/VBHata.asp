<%

On Error Resume Next
Err.Raise 55
Err.number = 8000
Err.description = "özel hata var"
hatanumarasi = Err.number
hataaciklama = Err.description
If hatanumarasi <> 0 Then
 Response.Write "Hata var! Hata Numarasý ve açýklamasý : " 
 Response.write hatanumarasi & " - " & hataaciklama
End If 
%>
