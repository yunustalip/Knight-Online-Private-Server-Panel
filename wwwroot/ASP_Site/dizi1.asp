<%
'dim mevsim(3)
dim mevsim
redim mevsim(3)
mevsim(0)="�lkbahar"
mevsim(1)="Yaz"
mevsim(2)="Sonbahar"
mevsim(3)="K��"

'dim mevsim
'mevsim=array("�lkbahar","Yaz","Sonbahar","K��")

for i = Lbound(mevsim) to ubound(mevsim) 
%><%response.Write(mevsim(i)) & "<br>"
next

'for i = Lbound(mevsim) to ubound(mevsim) 
'mevsim(i)=0
'next
erase mevsim 
' string "",  say�=0, obje = nothing

for i = Lbound(mevsim) to ubound(mevsim) 
%>"<%response.Write(mevsim(i)) & """<br>"
next
%>

