<%

Set adocon = Server.CreateObject("ADODB.Connection")
adocon.open= "driver={SQL Server};server=localhost;uid=radyolarimadmin;pwd=864327142358;database=radyolarimdb" 



Set musiclist = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM musiclist"
musiclist.Open strSQL, adoCon ,1,3
set msc=adocon.execute("select * from musiclist where id=859") 
Set BinaryStream = Server.CreateObject("ADODB.Stream") 
BinaryStream.Type = 1
BinaryStream.Open 
binarystream.write msc("albumcover")
BinaryStream.SaveToFile "D:\ARонV\Radyo\RadyoPlayer\SarkiBilgi\images\deneme.png", 2
Set BinaryStream = Nothing
%>