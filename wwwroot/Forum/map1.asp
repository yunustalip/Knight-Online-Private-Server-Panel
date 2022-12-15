<!--#INCLUDE file="forumayar.asp"-->
<div align="center">
<table  width="100%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" align="left" valign="top">
<BR><BR>
<B>FORUM KATEGORÝLERÝ</B><BR>
<%
'ANA KATLAR DÖK
sor = "Select * from grup where altgrp=0 order  by grp asc"  
forum.Open sor,forumbag,1,3
if forum.eof or forum.bof then
Response.Write "<BR><BR><BR><center><B>Kayýt yok</B><P>"
Response.End
End If
do while not forum.eof  
%>
<IMG SRC="images/cat.gif" WIDTH="20" HEIGHT="15" BORDER="0" ALT="">
<A HREF="default.asp?part=grp&id=<%=forum("id")%>"><B><%=forum("grp")%></B></A>
<BR>
<!-- ALT KATLAR -->
<%sor = "Select * from grup where altgrp="&forum("id")&" order  by grp asc"  
forum1.Open sor,forumbag,1,3
do while not forum1.eof  %>
<IMG SRC="images/dal.gif" WIDTH="15" HEIGHT="14" BORDER="0" ALT="">
<IMG SRC="images/file.gif" WIDTH="13" HEIGHT="13" BORDER="0" ALT="">
<A HREF="default.asp?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>"><%=forum1("grp")%></A>
<BR>
<% 
forum1.movenext 
loop 
forum1.close
forum.movenext 
loop 
forum.close

set forum =Nothing
set forum1 =nothing
set forum2 =nothing
%>

</td></tr></table>