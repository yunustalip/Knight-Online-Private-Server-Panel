<%
Class Ekrana_Yazdir
Public yazdirilacakmetin
Public submetin

Public Function CFonk_yazdir
CFonk_submetin()
CFonk_yazdir=left(yazdirilacakmetin,10)
End Function

private Function CFonk_submetin
submetin=right(yazdirilacakmetin,10)
End Function

End Class

Set YeniNesne = New Ekrana_Yazdir
YeniNesne.yazdirilacakmetin = "fonksiyon metin yazdýrmasý"%>
<%="fonkdönen deðer : " & YeniNesne.CFonk_yazdir() & "<br>"%>
<%="subdönen deðer : " & YeniNesne.submetin & "<br>"%><br />
<%
Set BaskaNesne = New Ekrana_Yazdir
BaskaNesne.yazdirilacakmetin = "bu da baþka bir nesne"%>
<%="fonkdönen deðer : " & BaskaNesne.CFonk_yazdir() & "<br>"%>
<%="subdönen deðer : " & BaskaNesne.submetin & "<br>"%>
<%="subdönen deðer : " & YeniNesne.submetin & "<br>"%><br />
<%Set BaskaNesne1 = New Ekrana_Yazdir
BaskaNesne1.yazdirilacakmetin = "iyiler her zaman kazanýr"%>
<%="fonkdönen deðer : " & BaskaNesne1.CFonk_yazdir() & "<br>"%>
<%="subdönen deðer : " & BaskaNesne1.submetin & "<br>"%>
