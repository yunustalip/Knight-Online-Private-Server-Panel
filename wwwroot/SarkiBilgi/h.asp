<%
response.charset="iso-8859-9"

sarkici="S'ONSUZ"

Function Buyut(strBaslik)
strBaslik = Replace(strBaslik, "a", "A")
strBaslik = Replace(strBaslik, "b", "B")
strBaslik = Replace(strBaslik, "c", "C")
strBaslik = Replace(strBaslik, "�", "�")
strBaslik = Replace(strBaslik, "d", "D")
strBaslik = Replace(strBaslik, "e", "E")
strBaslik = Replace(strBaslik, "f", "F")
strBaslik = Replace(strBaslik, "g", "G")
strBaslik = Replace(strBaslik, "�", "�")
strBaslik = Replace(strBaslik, "h", "H")
strBaslik = Replace(strBaslik, "�", "I")
strBaslik = Replace(strBaslik, "i", "�")
strBaslik = Replace(strBaslik, "j", "J")
strBaslik = Replace(strBaslik, "k", "K")
strBaslik = Replace(strBaslik, "l", "L")
strBaslik = Replace(strBaslik, "m", "M")
strBaslik = Replace(strBaslik, "n", "N")
strBaslik = Replace(strBaslik, "o", "O")
strBaslik = Replace(strBaslik, "�", "�")
strBaslik = Replace(strBaslik, "p", "P")
strBaslik = Replace(strBaslik, "q", "Q")
strBaslik = Replace(strBaslik, "r", "R")
strBaslik = Replace(strBaslik, "s", "S")
strBaslik = Replace(strBaslik, "�", "�")
strBaslik = Replace(strBaslik, "t", "T")
strBaslik = Replace(strBaslik, "u", "U")
strBaslik = Replace(strBaslik, "�", "�")
strBaslik = Replace(strBaslik, "v", "V")
strBaslik = Replace(strBaslik, "w", "W")
strBaslik = Replace(strBaslik, "x", "X")
strBaslik = Replace(strBaslik, "y", "Y")
strBaslik = Replace(strBaslik, "z", "Z")
strBaslik = Replace(strBaslik, "&", "&#38;")
strBaslik = Replace(strBaslik, "'", "&#39;")
strBaslik = Replace(strBaslik, "`", "&#39;")
Buyut = strBaslik
End Function

sarkici=Buyut(sarkici)
sarki=Buyut(sarki)
album=Buyut(album)

Response.Write sarkici
%>