<%
Class Ekrana_Yazdir

Sub CSub_yazdir(CSubParametre)
response.write CSubparametre & "<br>"
End Sub

Function CFonk_yazdir(CFonkParametre)
CFonk_yazdir=left(CFonkParametre,10)
End Function

End Class

Set YeniEkranaYazdir = New Ekrana_Yazdir
YeniEkranaYazdir.CSub_Yazdir("nesne1 deneme")
YeniEkranaYazdir.CSub_Yazdir("nesne2 deneme")
metin="nesne3 deneme"
YeniEkranaYazdir.CSub_Yazdir(metin)

fonkmetin="fonksiyon metin yazdýrmasý"
fonkdeger=YeniEkranaYazdir.CFonk_yazdir(fonkmetin)
response.write "dönen deðer : " & fonkdeger & "<br>"
%>