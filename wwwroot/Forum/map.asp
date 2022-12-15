

<%
'*******************************************************
' Kodlarýmý kullandýðýnýz için teþekkürler
' Kullandýðýnýz siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalarýmý ziyaret etmeyi unutmayýnýz  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vardýr ...
' LÜTFEN BU TÜR ÇALIÞMALARIN ÖNÜNÜ KESMEMEK ÝÇÝN TELÝF YAZILARINI SÝLMEYÝN
' EMEÐE SAYGI LÜTFEN 
' KÝÞÝSEL KULLANIM ÝÇÝN ÜCRETSÝZDÝR DÝÐER KULLANIMLARDA HAK TALEP EDÝLEBÝLÝR
'*******************************************************
%>


<!--#INCLUDE file="forumayar.asp"-->

<select name="menu" onChange="location=document.jump.menu.options[document.jump.menu.selectedIndex].value;" value="Git">
<option value="" selected>Hýzlý Menu</option>

<!-- ANA KATLAR -->
<%sor = "Select * from grup where altgrp=0 order  by grp asc"  
forum.Open sor,forumbag,1,3
if forum.eof or forum.bof then
Response.Write "<BR><BR><BR><center><B>Kayýt yok</B><P>"
Response.End
End If
do while not forum.eof  
%>

<option value="default.asp?part=grp&id=<%=forum("id")%>">+<%=Left(forum("grp"),15)%></option>
<!-- ALT KATLAR -->
<%sor = "Select * from grup where altgrp="&forum("id")&" order  by grp asc"  
forum1.Open sor,forumbag,1,3
do while not forum1.eof  %>

<option value="default.asp?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>">
--<%=Left(forum1("grp"),15)%></option>

<% 
forum1.movenext 
loop 
forum1.close
forum.movenext 
loop 
forum.close
%>

</select>



<% 
set forum =Nothing
set forum1 =nothing
set forum2 =nothing
%>