<%
script=Request.ServerVariables("SCRIPT_NAME")

if Instr(1,script,"index.asp",1) then
baslik="Anasayfa"

Elseif Instr(1,script,"galeri.asp",1) Then

kategori=Request.QueryString("kategori")
if isnumeric(kategori)=false or kategori="" then
baslik="Resim Galerisi"
else
set rs = data.execute("SELECT id,isim FROM galeri_kat where id="& kategori &"")
if rs.eof then
baslik="Resim Galerisi"
else
baslik=rs("isim") &" | Resim Galerisi"
end if
rs.close : set rs=nothing
end if

Elseif Instr(1,script,"iletisim.asp",1) Then
baslik="Ýletiþim"

Elseif Instr(1,script,"anketler.asp",1) Then
baslik="Tüm Anketler"

Elseif Instr(1,script,"paylas.asp",1) Then
baslik="Paylaþ"

Elseif Instr(1,script,"hakkimda.asp",1) Then
baslik="Hakkýmda"
id1="0"

Elseif Instr(1,script,"zd.asp",1) Then
if (Request.QueryString("zd"))="yaz" then
baslik="Ziyaretçi Defterine Yaz"
else
baslik="Ziyaretçi Defteri"
end if

Elseif Instr(1,script,"etiket.asp",1) Then
etiket=Filtre(Trim(AramaFiltre(Request.QueryString("etiket"))))
baslik="Etiket: "&etiket

Elseif Instr(1,script,"arac.asp",1) Then
if (Request.QueryString("son"))="yorumlar" then
baslik="Son Yorumlar"
else
baslik="Blog Arþivi"
end if
Elseif Instr(1,script,"kategori.asp",1) Then
id=Filtre(request.querystring("id"))
			if isnumeric(id)=false or id="" then
					response.redirect "index.asp"
			end if
set bslk = data.execute("SELECT id,ad FROM kategori where id="& id &"")
if bslk.eof then
baslik="Kategori Yok"
else
baslik=bslk("ad")
end if
bslk.close : set bslk = nothing
Elseif Instr(1,script,"ara.asp",1) Then
sonuc=" için Arama Sonuçlarý"
baslik=Request.QueryString("ara")
baslik=Trim(Filtre(AramaFiltre(baslik)))&sonuc

Elseif Instr(1,script,"takvim.asp",1) Then
gun=filtre(Request.QueryString("gun"))
ay=filtre(Request.QueryString("ay"))
yil=filtre(Request.QueryString("yil"))
Tarih = ay & "/" & gun & "/" & yil
baslik=Tarih &", Tarihli Bloglar"

Elseif Instr(1,script,"404.asp",1) Then
adres = Request.ServerVariables("QUERY_STRING")
if adres="" then
id1="0"
else
ayir = split(adres,"/")
no = ayir(Ubound(ayir))
if no="" then no="bosluk"
tire = split(no,"-")
id1=filtre(tire(0))
end if
if not isnumeric(id1)=false then
set bslk = data.execute("SELECT id,konu FROM blog where id="& id1 &"")
if not bslk.eof then
baslik=bslk("konu")
else
baslik="404 Not Found / Bulunamadý"
end if
bslk.close : set bslk = nothing
else
baslik="404 Not Found / Bulunamadý"
end if
elseif Instr(1,script,"blog.asp",1) Then
id1=Request.QueryString("id")
if id1="" or isnumeric(id1)=false then
baslik="Kayýt Yok"
else
set bslk = data.execute("SELECT id,konu FROM blog where id="& id1 &"")
if not bslk.eof then
baslik=bslk("konu")
else
baslik="Kayýt Bulunamadý"
end if
end if
else
baslik=script
end if
%>