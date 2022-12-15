<!--#include file="_inc/conn.asp"-->
<!--#include file="Function.asp"-->
<link rel="stylesheet" type="text/css" href="css/webstyle.css" >
<style>
input{
font-size:10px
}
</style>
<%npcid=trim(Request.Querystring("npcid"))
If not npcid="" Then
Set knpc=Conne.Execute("select * from k_npc where ssid="&npcid&"")%>
<form action="npcedit.asp?is=kaydet" method="post" >
<input type="hidden" value="<%=knpc("ssid")%>" name="ssid">
<table width="350" border="0">
  <tr>
    <td width="91">Npc Adý:</td>
    <td width="150"><input type="text" value="<%=trim(knpc("strname"))%>" style="width:200px" name="strname"></td>
    </tr>
  <tr>
    <td width="91">Exp:</td>
    <td width="150"><input type="text" value="<%=knpc("iexp")%>" name="iexp"></td>
  </tr>
  <tr>
    <td>Np:</td>
    <td><input type="text" value="<%=knpc("iloyalty")%>" name="iloyalty"></td>
  </tr>
  <tr>
    <td>Hp:</td>
    <td><input type="text" value="<%=knpc("ihppoint")%>" name="ihppoint"></td>
  </tr>
  <tr>
    <td>Mp:</td>
    <td><input type="text" value="<%=knpc("smppoint")%>" name="smppoint"></td>
  </tr>
  <tr>
    <td>Para:</td>
    <td><input type="text" value="<%=knpc("imoney")%>" name="imoney"></td>
  </tr>
  <tr>
    <td>Sað El Silah No:</td>
    <td><input type="text" value="<%=knpc("iweapon1")%>" name="iweapon1"></td>
  </tr>
  <tr>
    <td>Sol El Silah No:</td>
    <td><input type="text" value="<%=knpc("iweapon2")%>" name="iweapon2"></td>
  </tr>
    <tr>
    <td>Defans:</td>
    <td><input type="text" value="<%=knpc("sac")%>" name="sac"></td>
  </tr>
  <tr>
  <td colspan="2" align="center"><input type="submit" value="Kaydet">&nbsp;<a href="default.asp?w8=npc&is=sil&id=<%=npcid%>" ><font class="style5">Npcyi sil</font></a></td>
  </tr>
</table>
</form>
<%End If
If Request.Querystring("is")="kaydet" Then
Conne.Execute("Update k_npc Set strname='"&Request.Form("strname")&"', iexp="&Request.Form("iexp")&",iloyalty="&Request.Form("iloyalty")&",ihppoint="&Request.Form("ihppoint")&",smppoint="&Request.Form("smppoint")&",imoney="&Request.Form("imoney")&",iweapon1="&Request.Form("iweapon1")&",iweapon2="&Request.Form("iweapon2")&",sac="&Request.Form("sac")&" where ssid="&Request.Form("ssid")&"")
Response.Write "<script>alert('Kayýt Edildi');window.close()</script>"
End If%>