<%
Sub Sub_yazdir(SubParametre)
response.write Subparametre & "<br>"
End Sub

Function Fonk_yazdir(FonkParametre)
Fonk_yazdir=left(FonkParametre,10)
End Function

metin="subroutine metin yazd�rmas�"
Sub_Yazdir(metin)

fonkmetin="fonksiyon metin yazd�rmas�"
fonkdeger=Fonk_yazdir(fonkmetin)
response.write "d�nen de�er : " & fonkdeger & "<br>"

%>