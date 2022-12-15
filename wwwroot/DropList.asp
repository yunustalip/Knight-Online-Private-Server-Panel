<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<%response.expires=0
response.charset="iso-8859-9"
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Droplist'")
If MenuAyar("PSt")=1 Then%>
<script language="javascript">
function ara(){
$.ajax({
   type: 'POST',
   url: 'dropara.asp',
   data: $('#droplist').serialize() ,
   start:  $('#ortabolum').html('<br><br><br><br><center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Arama Yapýlýyor Lütfen Bekleyin.</center>'),
   success: function(ajaxCevap) {
      $('#ortabolum').html(ajaxCevap);
   }
});
}
</script><br><img src="imgs/droplist.gif" /><br><br><br>
<form action="javascript:ara()" method="post" name="droplist" id="droplist">
<input type="text" name="dropname" style="background:#8E6400;color:#F9EED8">
<input type="submit" value="Ara" style="background:#8E6400;color:#F9EED8; border-collapse: separate">
</form>
<%Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>