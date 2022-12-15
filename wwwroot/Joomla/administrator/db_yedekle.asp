<!--#include file="admin_a.asp"-->
 <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Database Yedekleme</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td colspan="6" bgcolor="#666666">
<%

Server.ScriptTimeOut = 25000
Response.Charset="windows-1254"

Const AppCharset = "ISO-8859-9"
Const AppWinCharset = "Windows-1254"
Const AppLanguage = "tr"
Const AppName = "JoomlASP MySQL Yedek Sistemi"
Const AppTitle = "JoomlASP MySQL Yedek Sistemi"

strGun = Day(Date())
strAy = Month(Date())
strYil = Year(Date())
strSaat = Hour(Now())
strDakika = Minute(Now())

	strServer = mysql_server
	strDatabase = mysql_db
	strDBUser = mysql_user
	strDBPass = mysql_pass

	strBackupFolder = Left(Request.ServerVariables("PATH_TRANSLATED"),InStrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))
	strBackupFileName = strDatabase & "_" & strGun & "." & strAy & "." & strYil & "_" & strSaat & "." & strDakika & ".sql"
If Request.QueryString("Action") = "Backup" Then
	If DatabaseConnectionTest(strServer,strDatabase,strDBUser,strDBPass) = False Then	
		strErrMessage = "Database baðlantý hatasý.<br>Daha sonra tekrar deneyiniz."
	Else
		Call BackupConfirmationForm
	End If
	
ElseIf Request.QueryString("Action") = "DoBackup" Then
	strServer = mysql_server
	strDatabase = mysql_db
	strDBUser = mysql_user
	strDBPass = mysql_pass
	strBackupFolder = Request.Form("strBackupFolder")
	strBackupFileName = Request.Form("strBackupFileName")

	Call BackupDatabase

ElseIf Request.QueryString("Action") = "Download" Then
	Call DownloadFile(ConvertSlash(Request("FilePath")))

End If

If Request.QueryString("Action") = "" Then BackupConfirmationForm

%>

<% Sub BackupConfirmationForm  %>


<br />
<br />
<div align="center">
    <span style="font-weight: bold">MYSQL Baðlantýsý Kuruldu.<br>
    <br>
	Veritabanýný yedeklemek için Yedekle butonuna basýnýz...	</span><br><br>

	<form method="post" action="<%=Request.ServerVariables("script_name")%>?Action=DoBackup">
		<input type="hidden" name="strServer" value="<%=strServer%>">
		<input type="hidden" name="strDatabase" value="<%=strDatabase%>">
		<input type="hidden" name="strDBUser" value="<%=strDBUser%>">
		<input type="hidden" name="strDBPass" value="<%=strDBPass%>">

		<table width="650" border="0" align="center" cellpadding="2" cellspacing="0">
		  <tr>
			<td width="110">Yedekleme Klasörü</td>
			<td width="19">:</td>
		    <td width="478"><input type="text" name="strBackupFolder" value="<%=strBackupFolder%>" size="80" class="inputbox" /></td>
		  </tr>
		  <tr>
			<td>Dosya Adý</td>
			<td>:</td>
		    <td><input type="text" name="strBackupFileName" value="<%=strBackupFileName%>" size="80" class="inputbox" /></td>
		  </tr>
		  <tr>
			<td colspan="3"><div align="center">
			  <input type="submit" value="Yedekle" class="button" onclick="Ac('bekleyin');" />
			  </div></td>
			</tr>
		</table>
	</form>
    </div>
<br />
<br />
<br />



<% End Sub %>

<%
Function DatabaseConnectionTest(strServer,strDatabase,strDBUser,strDBPass)
	If strServer = "" OR strDatabase = "" Then
		DatabaseConnectionTest = False
	Else
 		On Error Resume Next
		set Conn = server.CreateObject("ADODB.connection")	
		Conn.Open "DRIVER={MySQL ODBC 3.51 Driver};server="& strServer &";uid="& strDBUser &";pwd="& strDBPass &";database="& strDatabase &";option=3;"

		If Err.Number = 0 Then
			DatabaseConnectionTest = True
		Else
			DatabaseConnectionTest = False
		End If

		On Error GoTo 0
	End If
End Function
%>


<%
Function BackupDatabase

	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	Set nnFile = FSO.CreateTextFile(strBackupFolder & strBackupFileName)

	nnFile.WriteLine("/*")
	nnFile.WriteLine(AppName)
	nnFile.WriteLine("Source Host : " & strServer)
	nnFile.WriteLine("Source Database : " & strDatabase)
	nnFile.WriteLine("Date : " & Now())
	nnFile.WriteLine("*/")
	nnFile.WriteBlankLines(1)

	set Conn = server.CreateObject("ADODB.connection")	
	Conn.Open "DRIVER={MySQL ODBC 3.51 Driver};server="& strServer &";uid="& strDBUser &";pwd="& strDBPass &";database="& strDatabase &";option=3;"
		
	Set DBTables = Conn.OpenSchema(20)

	Do While Not DBTables.Eof

		If  DBTables("table_type")="TABLE" Then

			Set rs = Server.CreateObject("ADODB.RecordSet")
			rs.Open "SELECT * FROM "& DBTables("table_name") &" ", Conn, 1, 3

			'Set rs = Conn.Execute("SELECT * FROM "& DBTables("table_name") &" ")
			'===========

			nnFile.WriteLine("DROP TABLE IF EXISTS `"& DBTables("table_name") &"`;")
			nnFile.WriteLine("CREATE TABLE `"& DBTables("table_name") &"` (")

				TotalField = 0
				For Each Field in Rs.Fields
					If Field.Properties("ISAUTOINCREMENT") = True then
						PK_Name = Field.Name
						nnFile.WriteLine(Chr(9) & "`"& Field.Name &"` "& MySQLTypeDescription(Field.Type,Field.DefinedSize) &" NOT NULL auto_increment,")
					Else
						If CInt(Field.Type) = 204 Or CInt(Field.Type) = 205 Then
							nnFile.WriteLine(Chr(9) & "`"& Field.Name &"` "& MySQLTypeDescription(Field.Type,Field.DefinedSize) &",")
						Else
							nnFile.WriteLine(Chr(9) & "`"& Field.Name &"` "& MySQLTypeDescription(Field.Type,Field.DefinedSize) &" default NULL,")
						End If
					End If	
					
					TotalField = TotalField + 1
				Next

			nnFile.WriteLine(Chr(9) & "PRIMARY KEY  (`"& PK_Name &"`)")
			nnFile.WriteLine(Chr(9) & ") ENGINE=InnoDB  DEFAULT CHARSET=latin1;")
			nnFile.WriteBlankLines(1)

			'============

			Do While Not rs.Eof
				nnFile.Write("INSERT INTO `"& DBTables("table_name") &"` VALUES (")
				ThisField = 0
				For Each Field in Rs.Fields
					If IsNumeric(Field.Value) Then
						nnFile.Write(Field.Value)
					ElseIf Field.Value <> "" Then
						nnFile.Write("'" & FormatData(Field.Value) & "'")
					Else
						nnFile.Write("NULL")
					End If

					ThisField = ThisField + 1

					If ThisField < TotalField Then nnFile.Write(",")					
				Next
				nnFile.Write(");")
				nnFile.Write(vbNewLine)
			rs.MoveNext
			Loop

			nnFile.WriteBlankLines(1)

		rs.Close
		Set Rs = Nothing

		End If
	DBTables.MoveNext : Loop
	DBTables.Close : Set DBTables = Nothing

	nnFile.Close
	Set FSO = Nothing

	Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;"
	Response.Write "<p align=""center""><b>Yedekleme Baþarýlý</b></p><br><br>"
	Response.Write "<p align=""center""><input type=""button"" value=""Dosyayý Ýndir"" class=""inputtxt"" onclick=""location.href='?Action=Download&FilePath="& ConvertSlash(strBackupFolder & strBackupFileName) &"'""></p><br><br><br>"

End Function

'######################################
Function FormatData(data)
	If data <> "" Then
		data = Replace(data,Chr(39),Chr(92)&Chr(39))
	End If
   	FormatData = data
End Function

Function MySQLTypeDescription(x,y)
	Select Case X
	
		Case 2		Y = "smallint"
		Case 3		Y = "int(11)" 'int yada mediumint yada integer
		Case 4		Y = "float"	
		Case 5		Y = "double" 'double yada real
		Case 16		Y = "tinyint"
		Case 18		Y = "year"
		Case 20		Y = "bigint(20)"
		Case 129	Y = "enum" 'enum yada set
		Case 131	Y = "decimal(10,0)" 'decimal yada numeric
		Case 133	Y = "date"
		Case 134	Y = "time"	
		Case 135	Y = "datetime" 'datetime yada timestamp
		Case 200	Y = "varchar("& y &")" 'varchar yada char yada tinytext
		Case 201	Y = "mediumtext" 'text yada mediumtext yada longtext => Accessteki NOT'un karþýlýðý aralarýndaki farký boyutlarý belirliyor
		Case 204	Y = "tinyblob"
		Case 205	Y = "longblob" 'blob yada mediumblob yada longblob => aralarýndaki farký sadece boyutlarý belirliyor.
	End Select

	MySQLTypeDescription = Y
End Function

Function DownloadFile(FilePath)	
	Response.Clear 
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = FSO.GetFile(FilePath)
	intFilelength = objFile.Size 
	Set objStream = Server.CreateObject("ADODB.Stream" )   
	objStream.Open  
	objStream.Type = 1 
	objStream.LoadFromFile(FilePath)
                       
	Response.AddHeader "Content-Disposition" , "attachment; filename=" & objFile.Name  
	Response.AddHeader "Content-Length" , intFilelength  
	Response.CharSet = "UTF-8"   
	Response.ContentType = "application/octet-stream"     

	Response.BinaryWrite objStream.Read  
	Response.Flush
	Response.End
	objFile.Close : Set objFile = Nothing
	objStream.Close : Set objStream = Nothing
	Set FSO = Nothing
End Function

Function ConvertSlash(strURL)
	ConvertSlash = Replace(strURL,"\","/")
End Function
%> </td>
                  </tr>
                </table>
            </td>
          </tr>
        </table>
 <!--#include file="admin_b.asp"-->