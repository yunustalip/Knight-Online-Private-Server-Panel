<%

Set objFileScripting = CreateObject("Scripting.FileSystemObject")

Set objFolder = objFileScripting.GetFolder(server.mappath("/images"))


Set filecollection = objFolder.Files
For Each filename in filecollection
Filename=right(Filename,len(Filename)-InStrRev(Filename, "\"))
Response.Write"<A HREF=""http://resimindir.somee.com/images/" & filename & """>" & filename & "</A><BR>"
Next
%>
