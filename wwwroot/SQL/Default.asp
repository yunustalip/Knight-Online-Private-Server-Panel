<%
Response.Buffer=True
Act=Request("act")
Usr=Request("usr")
Pwd=Request("pwd")
Srv=Request("srv")
If Act="login" And Usr<>"" And Pwd<>"" And Srv<>"" Then
	rowCount = 0

	On Error Resume Next
	Set cnn=CreateObject("ADODB.Connection")
	sql = "SELECT * FROM MASTER.DBO.SYSDATABASES"
	cnn.open "driver={SQL server};User Id=" & Usr & ";PASSWORD=" & Pwd & ";SERVER=" & Srv & ";UID=;APP=Microsoft Development Environment"
	set rs=cnn.execute(sql)
	
	while not rs.eof
		rowCount = rowCount + 1
		rs.MoveNext
	wend  
	rs.close
	set rs=nothing
	cnn.close	

	If rowCount>0 Then 	
		Response.Cookies("SQLADMIN")("Srv")= Srv
		Response.Cookies("SQLADMIN")("Usr")= Usr
		Response.Cookies("SQLADMIN")("Pwd")= Pwd
		Response.Cookies("SQLADMIN").Expires=Now() + 90
		Response.Redirect("Console.htm")
	Else
		Response.write("<script>alert('Unable to connect to server " & Srv & ".');</script>")
	End If 		   
End If
Srv=Request.Cookies("SQLADMIN")("Srv")
Usr=Request.Cookies("SQLADMIN")("Usr")
%>
<html>
	<head>
		<title>SQL Admin</title>
		<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-9">
		<meta http-equiv="Content-Language" content="tr">
		<meta name="Generator" content="">
		<meta name="Author" content="Sezer Turkmen">
		<meta name="Keywords" content="SQL Admin">
		<meta name="Description" content="SQL Admin">
		<meta name="Copyright" content="© Copyright 2005">
		<link href="Style/Style.css" type="text/css" rel="stylesheet">
	</head>
	<script language="JavaScript">
		function jsLogin()
		{
			var userName=document.getElementById("txtUsername").value;
			var passWord=document.getElementById("txtPassword").value;
			var serverName=document.getElementById("txtServerName").value;
			window.location.href="Default.asp?Act=login&Usr="	+ userName + "&Pwd=" + passWord + "&Srv=" + serverName;
		}
	</script>
	<body >
		<table align="center" cellSpacing="0" cellPadding="0" border="0" width="100%" height="70%">
			<tr>				
				<td align="center">		
					<table align="center" cellSpacing="15" cellPadding="0" border="0" width="420" height="250" background="Image/splash.jpg">
						<tr>							
							<td height="10"></td>
						</tr>	
						<tr>				
							<td valign="top">					
								<table cellSpacing="2" cellPadding="2" border="0">	
									<tr>							
										<td colspan="2" style="color:#FFCC66"><b>Welcome to the SQL Administrator</b></td>
									</tr>	
									<tr>							
										<td colspan="2" style="color:#FFFFFF">Please enter your SQL Server credentials:</td>
									</tr>									
									<tr>							
										<td width="100" style="color:#FFFFFF">Username</td>
										<td><input type="text" id="txtUsername" class="cssTextBox" value="<%=Usr%>"></td>								
									</tr>		
									<tr>							
										<td width="100" style="color:#FFFFFF">Password</td>
										<td><input type="password" id="txtPassword" class="cssTextBox"></td>								
									</tr>	
									<tr>							
										<td width="100" style="color:#FFFFFF">Server</td>
										<td><input type="text" id="txtServerName" class="cssTextBox" value="<%=Srv%>"></td>								
									</tr>	
									<tr>							
										<td width="100">&nbsp;</td>
										<td align="right"><button id="cmdLogin" onclick="jsLogin();" style="width:75px">Login</button></td>						
									</tr>										
								</table>		
							</td>			
						</tr>
					</table>								
				</td>			
			</tr>
		</table>
	</body>
</html>	 