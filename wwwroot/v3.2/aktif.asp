<%
Sub LogActiveUser
Dim strActiveUserList
Dim intUserStart, intUserEnd
Dim strUser
Dim strDate

strActiveUserList = Application("ActiveUserList")

If Instr(1, strActiveUserList, Session.SessionID) > 0 Then
Application.Lock
intUserStart = Instr(1, strActiveUserList, Session.SessionID)
intUserEnd = Instr(intUserStart, strActiveUserList, "|")
strUser = Mid(strActiveUserList, intUserStart, intUserEnd - intUserStart)
strActiveUserList = Replace(strActiveUserList, strUser, Session.SessionID & ":" & Now())
Application("ActiveUserList") = strActiveUserList
Application.UnLock
Else
Application.Lock
Application("ActiveUsers") = CInt(Application("ActiveUsers")) + 1
Application("ActiveUserList") = Application("ActiveUserList") & Session.SessionID & ":" & Now() & "|"
Application.UnLock
End If
End Sub

Sub ActiveUserCleanup
Dim ix
Dim intUsers
Dim strActiveUserList
Dim aActiveUsers
Dim intActiveUserCleanupTime
Dim intActiveUserTimeout

intActiveUserCleanupTime = 0.5 
intActiveUserTimeout = 0.5  

If Application("ActiveUserList") = "" Then Exit Sub

If DateDiff("n", Application("ActiveUsersLastCleanup"), Now()) > intActiveUserCleanupTime Then

Application.Lock
Application("ActiveUsersLastCleanup") = Now()
Application.Unlock

intUsers = 0
strActiveUserList = Application("ActiveUserList")
strActiveUserList = Left(strActiveUserList, Len(strActiveUserList) - 1)

aActiveUsers = Split(strActiveUserList, "|")

For ix = 0 To UBound(aActiveUsers)
If DateDiff("n", Mid(aActiveUsers(ix), Instr(1, aActiveUsers(ix), ":") + 1, Len(aActiveUsers(ix))), Now()) > intActiveUserTimeout Then
aActiveUsers(ix) = "XXXX"
Else
intUsers = intUsers + 1
End If 
Next

strActiveUserList = Join(aActiveUsers, "|") & "|"
strActiveUserList = Replace(strActiveUserList, "XXXX|", "")

Application.Lock
Application("ActiveUserList") = strActiveUserList
Application("ActiveUsers") = intUsers
Application.UnLock

End If

End Sub



Call LogActiveUser()
Call ActiveUserCleanup()

Response.Write ""  &  Application("ActiveUsers") &  ""

%>