<!--#include file="_inc/conn.asp"-->
<%
dpx=request.form("dpx")
dpy=request.form("dpy")
did=request.form("did")
mapid=request.form("mapid")
if mapid="21" Then
dpx=round(replace(dpx,".",","))
dpy=round(replace(dpy,".",","))
End If
did=mid(did,2,len(did))
if isnumeric(did) and isnumeric(dpy) and isnumeric(dpx) and isnumeric(mapid) Then
Conne.Execute("update k_npcpos set leftx="&dpx&",topz="&dpy&" where id="&did&" and zoneid="&mapid)
End If
%>