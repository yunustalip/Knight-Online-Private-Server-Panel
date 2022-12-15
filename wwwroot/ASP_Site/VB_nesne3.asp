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
YeniNesne.yazdirilacakmetin = "fonksiyon metin yazdýrmasý"
%>
<%="fonkdönen deðer : " & YeniNesne.CFonk_yazdir() & "<br>"%>
<%="subdönen deðer : " & YeniNesne.submetin & "<br>"%>
