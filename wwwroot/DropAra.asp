<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<br><img src="imgs/droplist.gif" /><br><br><br><%response.expires=0
response.charset="iso-8859-9"
Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Droplist'")
If MenuAyar("PSt")=1 Then%>
<script language="javascript">
function ara(num){
$.ajax({
   url: 'dropara.asp?id='+num,
   start:  $('#ortabolum').html('<br><br><center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Arama Yapýlýyor Lütfen Bekleyin.</center>'),
   success: function(ajaxCevap) {
      $('#ortabolum').html(ajaxCevap);
   }
});
}
</script>
<%dim dropname,id
dropname=secur(request.form("dropname"))
id=trim(secur(Request.Querystring("id")))

if not id="" Then
if isnumeric(id)=false Then
Response.End
End If
Dim Mon,drop
Set Mon=Conne.Execute("select * from k_monster_item where iItem01="&id&" or iItem02="&id&" or iItem03="&id&" or iItem04="&id&" or iItem05="&id&"")
Set drop=Conne.Execute("select num,strname from item where num="&id&" order by strname desc")
%>
<table width="250">
<tr><td>&nbsp;</td></td>
<tr>
<td style="font-size:12px"><b>Monster</b></td><td style="font-size:12px"><b>Drop</b></td></tr>
<%If Not mon.Eof Then
Do While Not mon.Eof
Dim mons
Set mons=Conne.Execute("Select Strname,Ssid from k_monster where ssid="&mon("sindex")&"")
Response.Write "<tr><td style=""font-weight: bold; color: black;"">"&mons("strname")&"</td><td>"&drop("strname")&"</td></tr>"
mon.MoveNext
Loop
Else
Response.Write "<tr><td style=""font-weight: bold; color: black;"">Bu Item droplarda çýkmamaktadýr.</td></tr>"

End If%>

<%Else
If dropname="" Then
Response.Write "<br><b>Lütfen Bir Ýtem Adý Giriniz</b>"
Response.End
End If
If Len(dropname)<3 Then
Response.Write "<br><br><b>En Az 3 Karakter Giriniz.</b>"
Else
Set drop=Conne.Execute("select top 50 num,strname  from item where strname like '%"&dropname&"%' order by strname")
If not drop.eof Then
Response.Write "<br><br><b>Item Seçiniz</b><br>"
Do while not drop.eof
Response.Write "<a href=""#"" onclick=""ara('"&drop("num")&"');return false"">"&drop("strname")&"</a><br>"
Drop.MoveNext
Loop


Else
Response.Write "<br><br><b>Böyle Bir Item Bulunamadý.</b>"
End If
End If
End If

Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing
%>