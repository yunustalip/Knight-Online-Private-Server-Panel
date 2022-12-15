<%

SQLmod ="SELECT * from gop_modules where modul_izin = '"& 1 &"' AND modul_yer ='sol' order by modul_sira desc;"
set modul = server.createobject("ADODB.Recordset")
modul.open SQLmod , Baglanti
if modul.eof or modul.bof then 
Response.write " "
else


do while not modul.eof

%>
<%=modul("modul_icerik")%>
<%

modul.MoveNext
loop
modul.close
set modul=nothing
end if
%>