<%
dim mevsim
mevsim=array("Ýlkbahar","Yaz","Sonbahar","Kýþ")
metin = "Ben seni sen de beni tanýyoruz."
sayi=5:say=empty:t=null:tarih=datevalue("01/05/2005")

if isarray(mevsim) then response.Write("dizi")&"<br>"
if isdate("01/01/2006") then response.Write("tarihtir")&"<br>"
if isempty(empty) then response.Write("boþ")&"<br>"
if isnull(t) then response.Write("null")&"<br>"
if isnumeric(sayi) then response.Write("sayi")&"<br>"
if isobject(mevsim) then response.Write("nesne")&"<br>"
response.write vartype(tarih) & "<br>"
response.write typename(mevsim) & "<br>"
dim a(3)

response.write typename(null) & "<br>"

%>




<!--
VARTYPE için sabitler ve anlamlarý

Sabit        Deðer  Açýklama 
-----------  -----  ---------------------------------
vbEmpty        0     Empty (uninitialized) 
vbNull         1     Null (no valid data) 
vbInteger      2     Integer 
vbLong         3     Long integer 
vbSingle       4     Single floating-point number 
vbDouble       5     Double-precision fp number 
vbCurrency     6     Currency 
vbDate         7     Date 
vbString       8     String 
vbObject       9     Automation object  
vbError       10     Error 
vbBoolean     11     Boolean 
vbVariant     12     Variant (diziler için) 
vbDataObject  13     A data-access object 
vbByte        17     Byte 
vbArray      8192    Array 


-->

