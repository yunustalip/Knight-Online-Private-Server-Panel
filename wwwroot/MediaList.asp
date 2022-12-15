<!--#include file="_inc/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=1254">
<title><%dim sitesettings,sitebaslik
Set Sitesettings=Conne.Execute("select * from siteayar")
sitebaslik = sitesettings("sitebaslik")
Response.Write(sitebaslik&" Playlist")%></title>
<link rel="stylesheet" rev="stylesheet" type="text/css" href="RadyoFiles\skin.css">
<script language="javascript" type="text/javascript" src="RadyoFiles\skin.js"></script>
<script language="javascript">
function ListenNow(ids){
parent.window.document.getElementById('player').src="RadyoFiles/RadyoAsx.Asp?RadyoId="+ids;
}
</script>
</head>
<body scroll="no" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
				<td height="30" background="RadyoFiles/bg2.gif" id="pl_header"><table width="100%" border="0" cellpadding="0" cellspacing="0">
								<tr>
										<td height="30" class="padd_l3"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;">
														<tr>
																<td width="11"><img src="RadyoFiles/lcd2_left.gif" width="11" height="30"></td>
																<td nowrap background="RadyoFiles/lcd2_bg.gif" class="padd_t2 padd_l3 padding_r3"><div id="text_title" name="text_title" class="shadow bold">
                                                                  |<%Response.Write(sitebaslik&" Playlist")%></div></td>
																<td width="11"><img src="RadyoFiles/lcd2_right.gif" width="10" height="30"></td>
														</tr>
												</table></td>
										<td width="20" class="padd_r3"><img src="RadyoFiles/btn2_close.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_close','','RadyoFiles/btn2_close_c.gif',1)" alt="Kapat" name="img_close" width="21" height="30" border="0" align="absmiddle" style="cursor:pointer" onClick="window.close();"></td>
								</tr>
						</table></td>
		</tr>
		<tr>
			<td valign="top">
			<div id="playlist">
			<%Dim Radyo,Url,RadyoList
			Set RadyoList=Conne.Execute("Select * From RadyoList")
			Do While Not RadyoList.Eof%>
	<table width="100%" style="table-layout:fixed;" cellpadding="0" cellspacing="0">
	<tr onMouseOver="this.style.background='#E4E4E4'" onMouseOut="this.style.background=''">
	<td align="center" width="10%"><% =RadyoList("id")+1 %></td>
	<td title="<% =RadyoList("radyo") %>">
	<span style="cursor:pointer;" onClick="<%Response.Write("ListenNow('"&RadyoList("id")&"');ListenNow('"&RadyoList("id")&"');")%>"><% =RadyoList("Radyo") %></span>
	</td>
	<td align="right" >
	<img src="RadyoFiles/btn_listen.gif" style="cursor:pointer;" align="absmiddle" onClick="<%Response.Write("ListenNow('"&RadyoList("id")&"');ListenNow('"&RadyoList("id")&"');")%>"></td>
	</tr>
	</table>
	<%RadyoList.MoveNext
	Loop%>
	</div></td>
		</tr>
		<tr>
				<td height="30" background="RadyoFiles/bg2.gif" id="pl_footer" class="padd_l3 padd_r3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        		<tr>
        				<td><i>
                        <font size="1" style="font-style: normal; font-variant: normal; font-weight: normal; font-size: 8pt; font-family: Tahoma" color="#800000">
                        Copyright © 2011&nbsp; </font></i>
                        <font color="#800000">Asi Beşiktaşlı</font></td>
        				<td width="20" align="right"><span class="padd_r3"><img src="RadyoFiles/btn2_close.gif" alt="Kapat" name="img_close1" width="21" height="30" border="0" align="absmiddle" id="img_close1" onClick="window.close();" style="cursor:pointer" onMouseOver="MM_swapImage('img_close1','','RadyoFiles/btn2_close_c.gif',1)" onMouseOut="MM_swapImgRestore()"></span></td>
        				</tr>
        		</table>						</td>
		</tr>
</table>
</body>
</html>