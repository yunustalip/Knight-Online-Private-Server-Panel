<%
Sub Sub_yazdir(SubParametre)
response.write Subparametre & "<br>"
End Sub

Function Fonk_yazdir(FonkParametre)
Fonk_yazdir=left(FonkParametre,10)
End Function

metin="subroutine metin yazdýrmasý"
Sub_Yazdir(metin)

fonkmetin="fonksiyon metin yazdýrmasý"
fonkdeger=Fonk_yazdir(fonkmetin)
response.write "dönen deðer : " & fonkdeger & "<br>"

%>