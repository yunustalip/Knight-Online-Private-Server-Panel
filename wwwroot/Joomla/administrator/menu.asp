<%
set adminmenu = baglanti.Execute("select * from gop_adminmenua order by amord")
if adminmenu.eof or adminmenu.bof then
response.Write ""
else
do while not adminmenu.eof
anamenu_resim = adminmenu("amar")
anamenu_adi = adminmenu("amadi")

Response.Write "<table width=""101%"" border=""0"" cellpadding=""1"" cellspacing=""1"" bgcolor=""#FFFFFF""><tr align=""left""><td height=""20"" background=""../images/menu_p.png"" bgcolor=""#333333""><font color=""#FFCC00"">&nbsp;<img src=""adm_img/"& anamenu_resim &""" height=""18"" align=""absmiddle"">&nbsp;"

Execute anamenu_adi
Response.Write "</font></td></tr>"

set adminmenu2 = baglanti.Execute("select * from gop_adminmenub where amka='"&adminmenu("amka")&"';")
if adminmenu2.eof or adminmenu2.bof then
response.Write ""
else
do while not adminmenu2.eof

admin_altmenu = adminmenu2("ambadi")
admin_altmenu_link = adminmenu2("amblink")

Response.Write "<tr><td><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td bgcolor=""#CCCCCC"">&nbsp;<a href="""& admin_altmenu_link &""">"
                  
Execute admin_altmenu

Response.Write "</a></td></tr></table></td></tr>"

adminmenu2.movenext
loop
adminmenu2.close
set adminmenu2 = nothing
end if
            
Response.Write "</table>"
 
adminmenu.movenext
loop
adminmenu.close
set adminmenu = nothing
end if
%>