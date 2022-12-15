<%


'******************************************
'*** 	  Database System Type         ****
'******************************************

'Database Type
'strDatabaseType = "SQLServer"	'Microsoft SQL Server 2000, 2005, 2008, 2008 R2 (Supports Enterprise, Standard, Workgroup, Web, and Express Editions)
'strDatabaseType = "mySQL"	'MySQL 4.1 or MySQL 5.x
strDatabaseType = "Access"	'Microsoft Access Database (Very slow, not very good, best off avoided)



'******************************************
'*** 	      Microsoft Access         ****
'******************************************

'Microsoft Access is a flat file database system, it suffers from slow performance, limited 
'connections, and as a flat file it can be easly downloaded by a hacker if you do not secure 
'the database file!

'Virtual path to database
strDbPathAndName = Server.MapPath("database\cyber-warrior.mdb")  'This is the path of the database from the applications location

'Physical path to database
'strDbPathAndName = "" 'Use this if you use the physical server path, eg:- "C:\Inetpub\private\wwForum.mdb"


'PLEASE NOTE: - For extra security it is highly recommended you change the name of the database, wwForum.mdb, 
'to another name and then replace the wwForum.mdb found above with the name you changed the forum database to.



'**********************************************************
'*** 	   Microsoft SQL Server and MySQL Server        ****
'**********************************************************

'Enter the details of your Microsoft SQL Server or MySQL Server and database below
'*********************************************************************************

strSQLServerName = "127.0.0.1" 'Holds the name of the SQL Server (This is the name/location or IP address of the SQL Server)
strSQLDBUserName = "talip" 'Holds the user name (for SQL Server Authentication)
strSQLDBPassword = "864327" 'Holds the password (for SQL Server Authentication)
strSQLDBName = "kn_online"

'*** Advanced Paging - Performance Boost ***
'Set this to true for advanced paging in SQL Server 2005/2008 and mySQL 
'If you use SQL Server 2005/2008 or mySQL this will give a massive performance boost to your forum
blnSqlSvrAdvPaging = True


'*** SQL Server DBO Owner ***
''Sets the schema owner for SQL Server (Usually DBO (DataBase Owner))
strDBO = "DBO"


'*** mySQL Database Driver ***
'Web Wiz Forums supports both myODBC 3.51 and myODBC 5.1 database drivers when used with the mySQL database. 
'Most web host support myODBC 3.51, but if your web host supports myODBC 5.1 I would recommend that you use that instead
strMyODBCDriver = "3.51"


'Set up the database table name prefix
'(This is useful if you are running multiple forums from one database)
strDbTable = "tbl"


%>
