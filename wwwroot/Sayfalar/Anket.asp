<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<%response.charset="iso-8859-9"
sy=secur(Request.Querystring("sy"))%>
<script language="javascript">
function vote(fid){
$.ajax({
   type: 'POST',
   url: 'sayfalar/anket.asp?sy=reg',
   data: $('#'+fid).serialize() ,
   start:  $('#voteload').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>Oyunuz Kayit Ediliyor. Lütfen Bekleyin.</center>'),
   success: function(ajaxCevap) {
      $('#voteload').html(ajaxCevap);
   }
});
}
function loads(){
$.ajax({
   url: 'sayfalar/anket.asp',
   start:  $('#voteload').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>Yükleniyor...</center>'),
   success: function(ajaxCevap) {
      $('#voteload').html(ajaxCevap);
   }
});
}
</script>
<%set anket=Conne.Execute("select * from anket") %>
    <div id="voteload" align="right">
    <table align="right" style="margin-right:10"><tr> <td>
	<fieldset >
    <legend ><strong class="stioff">Anket</strong></legend>
<%if sy="" Then%>
 <form action="javascript:vote('anket')" name="anket" id="anket">
 <table width="200" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="center"><strong><%=anket("anketsoru")%></strong></td>
  </tr>
  <tr>
    <td align="center"><%set toplamoy=Conne.Execute("select count(vote) as toplam from anketsecmen ")
	Response.Write "Toplam Oy: "&toplamoy("toplam")%></td>
  </tr>
  <% set topoy=Conne.Execute("select count(vote) as toplam from anketsecmen")
	for ixx=1 to 5
	if not anket("anketsec"&ixx)="" Then %>    
  <tr valign="middle">
    <td ><input name="anketoy" type="radio" value="<%=ixx%>"  id="<%=ixx%>" style="position:relative;top:3">
      <label for="<%=ixx%>"><%=anket("anketsec"&ixx)%></label></td>
	<td><%set ason=Conne.Execute("select count(vote) as toplam from anketsecmen where vote="&ixx&"")
	if not topoy("toplam")=0 Then
	Response.Write round(ason("toplam")/topoy("toplam")*100)&"%"
	else
	Response.Write "0%"
	End If%></td>
  </tr>
	<%End If
	next%>
  <tr>
    <td align="center"><input type="submit" class="styleform" value="Oy Ver"></td>
  </tr>
</table>
	</form>
    

<%elseif sy="reg" Then
if Session("login")="ok" Then
oy=request.form("anketoy")
if isnumeric(oy)=false Then
Response.End
End If
set oykontrol=Conne.Execute("select * from anketsecmen where voter='"&Session("username")&"'")
if oykontrol.eof Then
set ekle=Conne.Execute("insert into anketsecmen(voter,vote) values('"&Session("username")&"','"&oy&"')")
%>
 <table width="200" border="0" cellpadding="2" cellspacing="0">
  <tr>
    <td align="center"><strong><%=anket("anketsoru")%></strong></td>
  </tr>
  <tr>
    <td align="center"><%set toplamoy=Conne.Execute("select count(vote) as toplam from anketsecmen ")
	toplamoy=toplamoy("toplam")
	Response.Write "Toplam Oy: "&toplamoy %></td>
  </tr>
  <% for i=1 to 5
	if len(anket("anketsec"&i))>0 Then %>    
  <tr>
    <td><%Response.Write anket("anketsec"&i)
	set oy=Conne.Execute("select count(*) as toplam from anketsecmen where vote="&i&"")
	oy=oy("toplam")
	Response.Write " ("&oy&")"&"  "&cint(oy/toplamoy*100)&"%" %></td>
  </tr>
	<%End If
	next%>
  <tr>
    <td align="center"></td>
  </tr>
</table>
<%else
Response.Write "<font size='1'><b>Bu Ankete 1 Kez Oy Kullanabilirsiniz.</font><br><a href='javascript:loads()'>Geri dön</a>"
End If
else
Response.Write "<font size='1'><b>Lütfen Kullanýcý Giriþi Yapýnýz</b></font><br><a href='javascript:loads()'>Geri dön</a>"
End If
End If%>    </fieldset>
    </td>
</tr>
</table></div>