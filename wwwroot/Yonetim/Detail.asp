<%@ Language=VBScript %>
<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"-->
<%
	Response.Buffer=True	
	Action=Request("act")
	ObjectId=Request("objid")
	ObjectType=Request("objtype")
	ObjectName=Request("objname")
	DatabaseName=Request("dbname")
	SqlQuery=Request("SqlQuery")		
	

	ObjectText=""
	If Action="exec" Then
	   ObjectText=SqlQuery
	Else 
		Select Case ObjectType
		Case "Tables"	
			sql = "SELECT INFORMATION_SCHEMA.COLUMNS.TABLE_NAME AS TableName, INFORMATION_SCHEMA.COLUMNS.COLUMN_NAME AS ColumnName, INFORMATION_SCHEMA.COLUMNS.ORDINAL_POSITION AS OrdinalPosition,INFORMATION_SCHEMA.COLUMNS.DATA_TYPE AS DataType FROM INFORMATION_SCHEMA.COLUMNS WHERE INFORMATION_SCHEMA.COLUMNS.TABLE_NAME='" & ObjectName & "' ORDER BY INFORMATION_SCHEMA.COLUMNS.ORDINAL_POSITION"
			set rs= conne.Execute(sql)
			while not rs.eof
				ObjectText=ObjectText & Rpad(rs(1)," ",30) & UCase(rs(3))  & vbNewLine 
				rs.movenext
			wend   
			rs.close
			set rs=Nothing	
		Case "Views" 		
			sql = "SELECT SYSOBJECTS.NAME, SYSCOMMENTS.TEXT FROM SYSOBJECTS INNER JOIN SYSCOMMENTS ON SYSOBJECTS.id = SYSCOMMENTS.id WHERE SYSOBJECTS.NAME='" & ObjectName & "'"
			set rs= conne.Execute(sql)
			while not rs.eof
				ObjectText = ObjectText & rs("text")
				rs.movenext
			wend
			rs.close
			set rs=Nothing
		Case "Procedures" 
			ObjectText=""
			sql = "SELECT SYSOBJECTS.NAME, SYSCOMMENTS.TEXT FROM SYSOBJECTS INNER JOIN SYSCOMMENTS ON SYSOBJECTS.id = SYSCOMMENTS.id WHERE SYSOBJECTS.NAME='" & ObjectName & "'"
			set rs= conne.Execute(sql)
			while not rs.eof
				ObjectText = ObjectText & rs("text")
				rs.movenext
			wend
			rs.close
			set rs=Nothing
		Case "Functions"	 
			sql = "SELECT SYSOBJECTS.NAME, SYSCOMMENTS.TEXT FROM SYSOBJECTS INNER JOIN SYSCOMMENTS ON SYSOBJECTS.id = SYSCOMMENTS.id WHERE SYSOBJECTS.NAME='" & ObjectName & "'"
			set rs= conne.Execute(sql)
			while not rs.eof
				ObjectText = ObjectText & rs("text")
				rs.movenext
			wend
			rs.close
			set rs=Nothing
		End Select				   
	End If
	   
	Function Rpad (sValue, sPadchar, iLength)
	  Rpad = sValue & string(iLength - Len(sValue), sPadchar)
	End Function
					  
	Function Lpad (sValue, sPadchar, iLength)
	  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
	End Function
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
		<base target="_self">
	</head>
<body bgcolor="threedface" leftmargin="0" topmargin="0">
<table cellspacing="2" cellpadding="0" width="100%" height="100%" border="0" align="center">			
	<tr>				
		<td>					
			<table cellSpacing="3" cellPadding="0" width="100%" border="0">						
				<tr>							
					<td width="100%" style="font-size:12px;"><b><%=ObjectName%></b></td>								
				</tr>					
			</table>				
		</td>			
	</tr>
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="100%">
			<table width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">					
						<table width="100%" height="100%" cellSpacing="3" cellPadding="3" border="0">	
							<form name="frmSource" method="post" action="Detail.asp?dbname=<%=DatabaseName%>&objtype=<%=ObjectType%>&objid=<%=ObjectId%>&objname=<%=ObjectName%>&act=exec">
							<tr>
								<td width="100%" height="90%" align="center">
									<textarea wrap="off" name="SqlQuery" id="SqlQuery" cols="50" rows="50" style="width:100%;height:100%"><%=ObjectText%></textarea>
								</td>
							</tr>							
							</form>
						</table>   					
					</td>						
				</tr>	
			</table>
		</td>						
	</tr>	
	<tr><td height="1px"></td></tr>						
	<tr>				
		<td height="40" valign="TOP">
			<table cheight="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td height="100%" vAlign="top" align="center">					
						<table cellSpacing="1" cellPadding="0" border="0">
							<tr>
								<% If ObjectType<>"Tables" Then %><td vAlign="top"><button id="cmdOK" onclick="document.frmSource.submit();" style="width:75px;height:25px;">Save</button></td><% End If  %>
								<td vAlign="top"><button id="cmdCancel" onclick="window.close();" style="width:75px;height:25px;">Close</button></td>
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
<%		
If Action="exec" Then
	On Error Resume Next	
	Select Case ObjectType
	Case "Views"
		If INSTR(1,SqlQuery,ObjectName)>0 And INSTR(1,UCASE(SqlQuery),"VIEW")>0 And (INSTR(1,UCASE(SqlQuery),"CREATE")>0 Or INSTR(1,UCASE(SqlQuery),"ALTER")>0)  Then 
			conne.Execute(SqlQuery)	
			If ERR<>0 then 
				Response.write("<script>alert(""" & err.Description & """);</script>")
			Else
				Response.write("<script>alert('Your SQL Query Successfuly Executed...');</script>")
			End If	
		Else 
			Response.write("<script>alert('Error! Your SQL must include CREATE or ALTER VIEW keys and view name must be " & ObjectName & "...');</script>")
		End If 
	Case "Procedures"
		If INSTR(1,SqlQuery,ObjectName)>0 And INSTR(1,UCASE(SqlQuery),"PROCEDURE")>0 And (INSTR(1,UCASE(SqlQuery),"CREATE")>0 Or INSTR(1,UCASE(SqlQuery),"ALTER")>0) Then 
			conne.Execute(SqlQuery)	
			If ERR<>0 then 
				Response.write("<script>alert(""" & err.Description & """);</script>")
			Else
				Response.write("<script>alert('Your SQL Query Successfuly Executed...');</script>")
			End If	
		Else 
			Response.write("<script>alert('Error! Your SQL must include CREATE or ALTER PROCEDURE keys and procedure name must be " & ObjectName & "...');</script>")
		End If 
	Case "Functions"
		If INSTR(1,SqlQuery,ObjectName)>0 And INSTR(1,UCASE(SqlQuery),"FUNCTION")>0 And (INSTR(1,UCASE(SqlQuery),"CREATE")>0 Or INSTR(1,UCASE(SqlQuery),"ALTER")>0) Then 
			conne.Execute(SqlQuery)	
			If ERR<>0 then 
				Response.write("<script>alert(""" & err.Description & """);</script>")
			Else
				Response.write("<script>alert('Your SQL Query Successfuly Executed...');</script>")
			End If	
		Else 
			Response.write("<script>alert('Error! Your SQL must include CREATE or ALTER FUNCTION keys and function name must be " & ObjectName & "...');</script>")
		End If 
	End Select	
End If 
conne.Close	 
%>				
<script>document.body.oncontextmenu=function (event){return true};</script>		 