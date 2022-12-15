<%Session.codepage="1254"
response.expires=-1

REFERER_URL = Request.ServerVariables("HTTP_REFERER")

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If

If REFERER_DOMAIN="http://fmradyodinle.net" or  REFERER_DOMAIN="http://www.fmradyodinle.net" Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

If Instr(Request.ServerVariables("ALL_HTTP"),"HTTP_X_REQUESTED_WITH:")>0 Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

url= "http://www.showradyo.com.tr/rds/vivasongdata.xml"
Set createxml = Server.CreateObject("msxml2.DOMDocument")
createxml.async = false
createxml.SetProperty "ServerHTTPRequest", True
createxml.Load(url)
Set xmllist = createxml.getElementsByTagName("rds")

for each xmlveri in xmllist
artist=xmlveri.childNodes(0).attributes.getNamedItem("artist").nodeValue
album=xmlveri.childNodes(0).attributes.getNamedItem("album").nodeValue
cover=xmlveri.childNodes(0).attributes.getNamedItem("cover").nodeValue
next

Function Base64Decode(ByVal base64String)
Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim dataLength, sOut, groupBegin
 
base64String = Replace(base64String, vbCrLf, "")
base64String = Replace(base64String, vbTab, "")
base64String = Replace(base64String, " ", "")

dataLength = Len(base64String)
If dataLength Mod 4 <> 0 Then
Err.Raise 1, "Base64Decode", "Bad Base64 string."
Exit Function
End If

For groupBegin = 1 To dataLength Step 4
Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
numDataBytes = 3
nGroup = 0

For CharCounter = 0 To 3

thisChar = Mid(base64String, groupBegin + CharCounter, 1)

If thisChar = "=" Then
numDataBytes = numDataBytes - 1
thisData = 0
Else
thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
End If
If thisData = -1 Then
Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
Exit Function
End If

nGroup = 64 * nGroup + thisData
Next

nGroup = Hex(nGroup)

nGroup = String(6 - Len(nGroup), "0") & nGroup

pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 5, 2)))

sOut = sOut & Left(pOut, numDataBytes)
Next

Base64Decode = sOut
End Function

if len(cover)>0 Then
kod=Base64Decode(cover)

Function StringToMultiByte(S)
Dim i, MultiByte
For i=1 To Len(S)
MultiByte = MultiByte & ChrB(Asc(Mid(S,i,1)))
Next
StringToMultiByte = MultiByte
End Function

artist=replace(artist," ","_")
album=replace(album," ","_")

Response.AddHeader "Content-Disposition", "inline; filename="&artist&"-"&album&".png"
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.ContentType = "image/png"
Response.binarywrite StringToMultiByte(kod)
End If
%>