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
YeniNesne.yazdirilacakmetin = "fonksiyon metin yazd�rmas�"%>
<%="fonkd�nen de�er : " & YeniNesne.CFonk_yazdir() & "<br>"%>
<%="subd�nen de�er : " & YeniNesne.submetin & "<br>"%><br />
<%
Set BaskaNesne = New Ekrana_Yazdir
BaskaNesne.yazdirilacakmetin = "bu da ba�ka bir nesne"%>
<%="fonkd�nen de�er : " & BaskaNesne.CFonk_yazdir() & "<br>"%>
<%="subd�nen de�er : " & BaskaNesne.submetin & "<br>"%>
<%="subd�nen de�er : " & YeniNesne.submetin & "<br>"%><br />
<%Set BaskaNesne1 = New Ekrana_Yazdir
BaskaNesne1.yazdirilacakmetin = "iyiler her zaman kazan�r"%>
<%="fonkd�nen de�er : " & BaskaNesne1.CFonk_yazdir() & "<br>"%>
<%="subd�nen de�er : " & BaskaNesne1.submetin & "<br>"%>
