<% if Session("yetki")="1" Then%>
<!--#include file="../_inc/conn.asp"-->
<!--#include file="../md5.asp"-->
<body bgcolor="#DCD1BA">
<script type="text/javascript" src="../js/jquery.js"></script>
<style>
body{
background-color:#000000;
font-weight:bold;
color:#00FF00;
}
input{
background-color:#000000;
font-weight:bold;
color:#00FF00;
border:solid 1px;

}
</style>
<%komut=Request.Querystring("komut")

if komut="savas" Then%>
<style>
body{
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
	text-decoration: none;
	font-weight: bold;
}

</style><center>Oto Savaþ</center><br>
<form action="otokomut.asp?komut=savasac" method="post">
Savaþ Saati 1: <input type="text" name="saat1" size="1" style="border-style:solid 1px">:<input type="text" name="dakika1" size="1" style="border-style:solid 1px"> Savaþ Komutu 1: <input type="text" name="savaskomut1" size="6" style="border-style:solid 1px" value="/open2"><br>
Savaþ Saati 2: <input type="text" name="saat2" size="1" style="border-style:solid 1px">:<input type="text" name="dakika2" size="1" style="border-style:solid 1px"> Savaþ Komutu 2: <input type="text" name="savaskomut2" size="6" style="border-style:solid 1px" value="/open2"><br>
Savaþ Saati 3: <input type="text" name="saat3" size="1" style="border-style:solid 1px">:<input type="text" name="dakika3" size="1" style="border-style:solid 1px"> Savaþ Komutu 3: <input type="text" name="savaskomut3" size="6" style="border-style:solid 1px" value="/open2"><br>
<br>Savaþ Süresi: <input type="text" name="savassaat" size="1" style="border-style:solid 1px">:<input type="text" name="savasdakika" size="1" style="border-style:solid 1px"><br>
<br><br>
<input type="submit" value="Kayýt Et" style="width:200px"><br><br>
Kullaným: Saat veya dakikan girerken baþýna 0 koyarak giriniz.(Örn: 19:03)
Kullanmadýðýnýz savaþ saatleri ni boþ býrakýnýz
</form>
<%elseif komut="savasac" Then
saat1=request.form("saat1")
saat2=request.form("saat2")
saat3=request.form("saat3")
dakika1=request.form("dakika1")
dakika2=request.form("dakika2")
dakika3=request.form("dakika3")
ssaat=request.form("savassaat")
sdakika=request.form("savasdakika")
notice1=request.form("savaskomut1")
notice2=request.form("savaskomut2")
notice3=request.form("savaskomut3")

if saat1="" Then
saat1=99
End If
if saat2="" Then
saat2=99
End If
if saat3="" Then
saat3=99
End If
if dakika1="" Then
dakika1=99
End If
if dakika2="" Then
dakika2=99
End If
if dakika3="" Then
dakika3=99
End If
tarih=date()
saat=time()


%>
<script type="text/javascript">

function AddZero(rakam)
{
return (rakam < 10) ? '0' + rakam : rakam;
}
	function timeDiff()	
	{
	var timeDifferense;
	var serverClock = new Date(<%=right(tarih,4)&","&mid(tarih,4,2)&","&mid(tarih,1,2)&","&replace(saat,":",",")%>);
	var clientClock = new Date();
	timeDiff = clientClock.getTime() - serverClock.getTime() - 500;
	runClock(timeDiff);
	}
	function runClock(timeDiff)
	{
	var now = new Date();
	var newTime;
	newTime = now.getTime() ;
	now.setTime(newTime);
var saat=now.getHours();
var dakika=now.getMinutes();
var saniye=now.getSeconds();
<%
dakika1=cint(dakika1)
dakika2=cint(dakika2)
dakika3=cint(dakika3)

kalansaat15=cint(saat1)+cint(ssaat)
kalansaat14=cint(saat1)+cint(ssaat)
kalansaat13=cint(saat1)+cint(ssaat)
kalansaat12=cint(saat1)+cint(ssaat)
kalansaat11=cint(saat1)+cint(ssaat)

kalansaat25=cint(saat2)+cint(ssaat)
kalansaat24=cint(saat2)+cint(ssaat)
kalansaat23=cint(saat2)+cint(ssaat)
kalansaat22=cint(saat2)+cint(ssaat)
kalansaat21=cint(saat2)+cint(ssaat)

kalansaat35=cint(saat3)+cint(ssaat)
kalansaat34=cint(saat3)+cint(ssaat)
kalansaat33=cint(saat3)+cint(ssaat)
kalansaat32=cint(saat3)+cint(ssaat)
kalansaat31=cint(saat3)+cint(ssaat)

kalandakika15=dakika1-5
kalandakika14=dakika1-4
kalandakika13=dakika1-3
kalandakika12=dakika1-2
kalandakika11=dakika1-1

if kalandakika15<0 Then
kalandakika15=dakika1+55
kalansaat15=kalansaat15-1
End If
if kalandakika14<0 Then
kalandakika14=dakika1+56
kalansaat14=kalansaat14-1
End If
if kalandakika13<0 Then
kalandakika13=dakika1+57
kalansaat13=kalansaat13-1
End If
if kalandakika12<0 Then
kalandakika12=dakika1+58
kalansaat12=kalansaat12-1
End If
if kalandakika11<0 Then
kalandakika11=dakika1+59
kalansaat11=kalansaat11-1
End If

kalandakika25=dakika2-5
kalandakika24=dakika2-4
kalandakika23=dakika2-3
kalandakika22=dakika2-2
kalandakika21=dakika2-1

if kalandakika25<0 Then
kalandakika25=dakika2+55
kalansaat25=kalansaat25-1
End If
if kalandakika24<0 Then
kalandakika24=dakika2+56
kalansaat24=kalansaat24-1
End If
if kalandakika23<0 Then
kalandakika23=dakika2+57
kalansaat23=kalansaat23-1
End If
if kalandakika22<0 Then
kalandakika22=dakika2+58
kalansaat22=kalansaat22-1
End If
if kalandakika21<0 Then
kalandakika21=dakika2+59
kalansaat21=kalansaat21-1
End If


kalandakika35=dakika3-5
kalandakika34=dakika3-4
kalandakika33=dakika3-3
kalandakika32=dakika3-2
kalandakika31=dakika3-1

if kalandakika35<0 Then
kalandakika35=dakika3+55
kalansaat35=kalansaat35-1
End If
if kalandakika34<0 Then
kalandakika34=dakika3+56
kalansaat34=kalansaat34-1
End If
if kalandakika33<0 Then
kalandakika33=dakika3+57
kalansaat33=kalansaat33-1
End If
if kalandakika32<0 Then
kalandakika32=dakika3+58
kalansaat32=kalansaat32-1
End If
if kalandakika31<0 Then
kalandakika31=dakika3+59
kalansaat31=kalansaat31-1
End If
%>


{
document.getElementById("clock_area").innerHTML = 'Server Saati: '+AddZero(saat) + ':' + AddZero(dakika) + ':' + AddZero(saniye) ;

<%if saat1<>99  Then%>
if (saat==<%=kalansaat15%>&&dakika==<%=kalandakika15%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 5 Dakika...");
}
if (saat==<%=kalansaat14%>&&dakika==<%=kalandakika14%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 4 Dakika...");
}
if (saat==<%=kalansaat13%>&&dakika==<%=kalandakika13%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 3 Dakika...");
}
if (saat==<%=kalansaat12%>&&dakika==<%=kalandakika12%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 2 Dakika...");
}
if (saat==<%=kalansaat11%>&&dakika==<%=kalandakika11%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 1 Dakika...");
}

if (AddZero(saat)==<%=saat1%>&&AddZero(dakika)==<%=dakika1%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþ Baþlamýþtýr. Herkese Ýyi Savaþlar Dileriz...");
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=<%=notice1%>");
}
if (saat==<%=cint(saat1)+cint(ssaat)%>&&dakika==<%=cint(dakika1)+cint(sdakika)%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=/close");
}
<%End If
if saat2<>99 Then%>
if (saat==<%=kalansaat25%>&&dakika==<%=kalandakika25%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 5 Dakika...");
}
if (saat==<%=kalansaat24%>&&dakika==<%=kalandakika24%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 4 Dakika...");
}
if (saat==<%=kalansaat23%>&&dakika==<%=kalandakika23%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 3 Dakika...");
}
if (saat==<%=kalansaat22%>&&dakika==<%=kalandakika22%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 2 Dakika...");
}
if (saat==<%=kalansaat21%>&&dakika==<%=kalandakika21%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 1 Dakika...");
}

if (AddZero(saat)==<%=saat2%>){
if(AddZero(dakika)==<%=dakika2%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþ Baþlamýþtýr. Herkese Ýyi Savaþlar Dileriz...");
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=<%=notice2%>");
}
if (saat==<%=cint(saat2)+cint(ssaat)%>&&dakika==<%=cint(dakika2)+cint(sdakika)%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=/close");
}
}
<%End If
if saat3<>99 Then%>
if (saat==<%=kalansaat35%>&&dakika==<%=kalandakika35%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 5 Dakika...");
}
if (saat==<%=kalansaat34%>&&dakika==<%=kalandakika34%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 4 Dakika...");
}
if (saat==<%=kalansaat33%>&&dakika==<%=kalandakika33%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 3 Dakika...");
}
if (saat==<%=kalansaat32%>&&dakika==<%=kalandakika32%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 2 Dakika...");
}
if (saat==<%=kalansaat31%>&&dakika==<%=kalandakika31%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþa Son 1 Dakika...");
}

if (AddZero(saat)==<%=saat3%>){
if(AddZero(dakika)==<%=dakika3%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=Savaþ Baþlamýþtýr. Herkese Ýyi Savaþlar Dileriz...");
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=<%=notice3%>");
}
if (saat==<%=cint(saat3)+cint(ssaat)%>&&dakika==<%=cint(dakika3)+cint(sdakika)%>&&AddZero(saniye)==00){
$('#notice').attr("src","../GmPage/gamem.asp?user=gmkomut&komut=/close");
}
}
<%End If%>


}
setTimeout('runClock(timeDiff)',1000);
}
</script>
<strong><div id="clock_area" align="left" style="margin-left:0;color:#000;" title="Server saati"></div></strong><script>timeDiff();</script>
<br><%if not saat1=99 Then
Response.Write "1. Savaþ Saati: "&saat1&":"&dakika1&" Bitiþ Saati: "& cint(saat1)+cint(ssaat)&":"&cint(dakika1)+cint(sdakika)
End If
if not saat2=99 Then
Response.Write "<br>2. Savaþ Saati: "&saat2&":"&dakika2&" Bitiþ Saati: "& cint(saat2)+cint(ssaat)&":"&cint(dakika2)+cint(sdakika)
End If
if not saat3=99 Then
Response.Write "<br>3. Savaþ Saati: "&saat3&":"&dakika3&" Bitiþ Saati: "& cint(saat3)+cint(ssaat)&":"&cint(dakika3)+cint(sdakika)
End If
Response.Write "<br><a href=""otokomut.asp?komut=savas"">Savaþ Ayarlarý</a>"
Response.Write vbcrlf&"<iframe id=""notice"" src="""" height=""0"" width=""0"" style=""visibility:hidden"">"
elseif komut="notice" Then%>
<style>
body{
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:11px;
	text-decoration: none;
	font-weight: bold;
}	
</style><center>Oto Notice</center><br>
<form action="otokomut.asp?komut=noticekayit" method="post">
Notice: <input type="text" name="notice" size="50" ><br>
Mssql Veri: <input type="text" name="sql" size="50"><br>
Notice Sýklýðý: <input type="text" name="dakika" size="3" >Dk. <input type="text" name="saniye" size="3" >Sn.<br><br>
<input type="submit" value="Kayýt Et" style="width:200px">
</form><br><br>Kullaným: Notice yazarken databaseden veri çekecekseniz deðiþkenleri komut1,komut2 ... þeklinde yazýnýz.
<br><u>Örn:</u><br>Notice: <font color="red">Online Oyuncu Sayýsý: komut1</font>
<br>Mssql Veri: <font color="red">select count(*) from currentuser</font>
<br>Çýktýsý þu þekilde olacatýr. <font color="red">Online Oyunucu Sayýsý: 38</font> gibi...
<%
elseif komut="noticekayit" Then
notice=request.form("notice")
dakika=request.form("dakika")
saniye=request.form("saniye")

If dakika="" Then
dakika=0
End If
If saniye="" Then saniye=0

if isnumeric(dakika)=false Then
Response.Write "Dakika Bölümünü Boþ Býrakmayýn"
Response.End
End If


If saniye=0 and dakika=0 Then
Response.Write "Geçerli Süre Girin."
Response.End
End If

sql=request.form("sql")
if not sql="" Then
set komutcalistir=Conne.Execute(sql)
if not komutcalistir.eof Then
for x=0 to 50
if instr(notice,"komut"&x+1)>0 Then
notice=replace(notice,"komut"&x+1,komutcalistir(x))
End If
next
End If
End If

tarih=date()
saat=time()
%>
<script type="text/javascript">
var counter = 30;

function AddZero(rakam)
	{
	return (rakam < 10) ? '0' + rakam : rakam;
	}

function AddZeroMnth(rakam)
	{
	rakam = rakam + 1
	return (rakam < 10) ? '0' + rakam : rakam;
	}

	function timeDiff()	
		{
		var timeDifferense;
		var serverClock = new Date(<%=right(tarih,4)&","&mid(tarih,4,2)&","&mid(tarih,1,2)&","&replace(saat,":",",")%>);
		var clientClock = new Date();
		var serverSeconds;
		var clientSeconds;
	
		timeDiff = clientClock.getTime() - serverClock.getTime() - 500;
		runClock(timeDiff);
		}
	function runClock(timeDiff)
		{

		var now = new Date();
		var newTime;
		newTime = now.getTime() - timeDiff;
		now.setTime(newTime);
		{
			if (counter > 0){
			document.getElementById('clock_area').title = 'Server saati';
			}
			counter--;
			document.getElementById("clock_area").innerHTML = 'Server Saati: '+AddZero(now.getHours()) + ':' + AddZero(now.getMinutes()) + ':' + AddZero(now.getSeconds()) ;
		}
		setTimeout('runClock(timeDiff)',1000);
		}


    $(document).ready(function() {
    $('#log').append('<br><%=notice%>');
  });


var dakika = <%=dakika%>;
var saniye = <%=saniye+1%>;

function Kontrol() {
saniye = saniye-1;

if(dakika>9){
$("span#2").text(dakika+":");
}
else{
$("span#2").text("0"+dakika+":");
}
if (saniye>9){
$("span#1").text(saniye);
}
else{
$("span#1").text("0"+saniye);
}

if(saniye>0) {
setTimeout("Kontrol()", 1000)
}
else if(saniye=1&&dakika>0){
dakika=dakika-1;
saniye=60;
setTimeout("Kontrol()", 1000);
}
else{
dakika=<%=dakika%>;
saniye=<%=saniye+1%>;
setTimeout("Kontrol()", 1000);
$('#notice').attr("src","../Gmpage/gamem.asp?user=gmkomut&komut=<%=notice%>");
$('#log').append('<br><%=notice%>')
}



}
  $(document).ready(function() {
    Kontrol();
  });

</script><strong><div id="clock_area" align="left" style="margin-left:0" title="Server saati"></div></strong><script>timeDiff();</script>
Bir Sonraki Noticeye Kalan Süre: <span id="2"> </span><span id="1"> </span>
<br><b>Notice</u>: <code><%=notice%></code>

<%Response.Write "<iframe id=""notice"" src=""../Gmpage/gamem.asp?user=gmkomut&komut="&notice&""" height=""0"" width=""0"" style=""visibility:hidden"">"
End If %>



<%End If%>