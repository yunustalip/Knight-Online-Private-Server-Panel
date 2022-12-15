<%Function FormatText(sIn)
	Dim charPos, sOut, curChar, urlChar, urlPos, urlString, imgChar, imgPos, imgString, smilebit
	
	For charPos = 1 To Len(sIn)
		curChar = Mid(sIn, charPos, 1)
		if mid(sIn, charPos, 5) = "[IMG]" Then
			sOut = sOut + "<img src="""
			charPos = charPos + 5
			imgString = ""
			imgChar = ""
			imgPos = charPos
			do until mid(sIn, imgPos, 6) = "[/IMG]"
				imgChar = mid(sIn, imgPos, 1)
				imgString = imgString + imgChar
				imgPos = imgPos + 1
				if imgPos > len(sIn) Then
					sOut = sOut + imgString + """>"
					exit for
				End If


		loop
			imgString = left(imgString, len(imgString))
			sOut = sOut + imgString + """>"
			charPos = charpos + (imgPos - charPos + 5)
			
		else
			sOut = sOut + curChar
		End If

     

	Next

	
	sOut = Replace(sOut, vbcrlf, "<br>")
	sOut = Replace(sOut, "[b]", "<b>")
	sOut = Replace(sOut, "[s]", "<s>")
	sOut = Replace(sOut, "[/s]", "</s>")
	sOut = Replace(sOut, "[/b]", "</b>")
	sOut = Replace(sOut, "[i]", "<i>")
	sOut = Replace(sOut, "[mail]", "<a href=""mailto:")
	sOut = Replace(sOut, "[/mail]", """>")
	sOut = Replace(sOut, "[/mailend]", "</a>")
	sOut = Replace(sOut, "[/i]", "</i>")
	sOut = Replace(sOut, "[u]", "<u>")
	sOut = Replace(sOut, "[hr]", "<hr>")
	sOut = Replace(sOut, "[list]", "<ul>")
	sOut = Replace(sOut, "[li]", "<li>")
	sOut = Replace(sOut, "[/li]", "</li>")
	sOut = Replace(sOut, "[/list]", "</ul>")
	sOut = Replace(sOut, "[left]", "<div align=""left"">")
	sOut = Replace(sOut, "[right]", "<div align=""right"">")
	sOut = Replace(sOut, "[center]", "<div align=""center"">")
	sOut = Replace(sOut, "[/left]", "</div>")
	sOut = Replace(sOut, "[/right]", "</div>")
	sOut = Replace(sOut, "[/center]", "</div>")
	sOut = Replace(sOut, "[size]", "<font style=""font-size:9pt"">")
	sOut = Replace(sOut, "[color]", "<font color=""black"">")
	sOut = Replace(sOut, "[color:black]", "<font color=""black"">")
	sOut = Replace(sOut, "[color:blue]", "<font color=""blue"">")
	sOut = Replace(sOut, "[color:red]", "<font color=""red"">")
	sOut = Replace(sOut, "[color:darkred]", "<font color=""darkred"">")
	sOut = Replace(sOut, "[color:yellow]", "<font color=""yellow"">")
	sOut = Replace(sOut, "[color:orange]", "<font color=""#FF9900"">")
	sOut = Replace(sOut, "[color:darkorange]", "<font color=""#FF6600"">")
	sOut = Replace(sOut, "[color:darkblue]", "<font color=""#330099"">")
	sOut = Replace(sOut, "[color:green]", "<font color=""#669933"">")
	sOut = Replace(sOut, "[color:mor]", "<font color=""#800080"">")
	sOut = Replace(sOut, "[color:lightgreen]", "<font color=""#66FF00"">")
	sOut = Replace(sOut, "[color:grey]", "<font color=""#CCCCCC"">")
	sOut = Replace(sOut, "[size:8pt]", "<font style=""font-size:8pt"">")
	sOut = Replace(sOut, "[size:9pt]", "<font style=""font-size:9pt"">")
	sOut = Replace(sOut, "[size:10pt]", "<font style=""font-size:10pt"">")
	sOut = Replace(sOut, "[size:11pt]", "<font style=""font-size:11pt"">")
	sOut = Replace(sOut, "[size:12pt]", "<font style=""font-size:12pt"">")
	sOut = Replace(sOut, "[size:13pt]", "<font style=""font-size:13pt"">")
	sOut = Replace(sOut, "[size:14pt]", "<font style=""font-size:14pt"">")
	sOut = Replace(sOut, "[size:18pt]", "<font style=""font-size:18pt"">")
	sOut = Replace(sOut, "[face]", "<font face=""Arial"">")
	sOut = Replace(sOut, "[face:Verdana]", "<font face=""Verdana"">")
	sOut = Replace(sOut, "[face:Arial]", "<font face=""Arial"">")
	sOut = Replace(sOut, "[face:Comic Sans MS]", "<font face=""Comic Sans MS"">")
	sOut = Replace(sOut, "[face:Times New Roman]", "<font face=""Times New Roman"">")
	sOut = Replace(sOut, "[/face]", "</font>")
	sOut = Replace(sOut, "[/color]", "</font>")
	sOut = Replace(sOut, "[/size]", "</font>")
	sOut = Replace(sOut, "[/u]", "</u>")
	sOut = Replace(sOut, "[quote]", "<blockquote style=""border-style:solid;border-width:thin;border-color:#990000""><font face=""verdana,arial,helvetica"" size=""1""><i><b>Quote</b></i></font><hr><div style=""background-color:#FFFFFF;border-color:#000000;border-style:solid;border-width:thin""><font color=""#990000"" style=""font-size:8pt;"" face=""Verdana, Arial, Helvetica, sans-serif"">")
	sOut = Replace(sOut, "[/quote]", "</font></div><hr></blockquote>")
	sOut = Replace(sOut, "[code]", "<blockquote style=""background-color:#FFFFFF;border-color:#990000""><font face=""verdana,arial,helvetica"" size=""1""><i><b>Code</b></i></font><hr><div style=""background-color:#CCCCCC""><font color=""#990000"" style=""font-size:8pt;"" face=""Verdana, Arial, Helvetica, sans-serif"">")
	sOut = Replace(sOut, "[/code]", "</font></div><hr></blockquote>")
	sOut = Replace(sOut, "</blockquote><br><br><br>", "</blockquote>")
	sOut = Replace(sOut, "</blockquote><br><br>", "</blockquote>")
	sOut = Replace(sOut, "</blockquote><br>", "</blockquote>")
	sOut=Replace(sOut,"[yt]","<object width=""425"" height=""355""><embed src='http://www.youtube.com/v/",1,-1,1)
	sOut=Replace(sOut,"[/yt]","' type='application/x-shockwave-flash' wmode='transparent' width='425' height='348'></embed></object>",1,-1,1)
        sOut=Replace(sOut, "[img]","<img src=""",1,-1,1) 
	sOut=Replace(sOut, "[/img]",""" border=""0"">",1,-1,1)

Do while InStr(1,sOut,"[url=", 1)>0 and InStr(1,sOut,"[/url]",1)>0
linkbaslangic=InStr(1,sOut,"[url=",1)
linkbitis=InStr(linkbaslangic,sOut,"[/url]",1)+6
if linkbitis<linkbaslangic Then linkbitis=linkbaslangic+7
linkbasligi=Trim(Mid(sOut,linkbaslangic,(linkbitis-linkbaslangic)))
linkadresi=linkbasligi
linkadresi=Replace(linkadresi,"[url=","<a target='_blank' href=""",1,-1,1)
if InStr(1,linkadresi,"[/url]",1) Then
linkadresi=Replace(linkadresi,"[/url]","</a>",1,-1,1)
linkadresi=Replace(linkadresi,"]",""">",1,-1,1)
Else
linkadresi=linkadresi&">"
End If
sOut=Replace(sOut,linkbasligi,linkadresi,1,-1,1)
Loop
Do while InStr(1,sOut,"[url]",1)>0  and InStr(1,sOut,"[/URL]",1)>0
linkbaslangic=InStr(1,sOut,"[url]",1)
linkbitis=InStr(linkbaslangic,sOut,"[/url]",1)+6
if linkbitis<linkbaslangic Then linkbitis=linkbaslangic+6
linkbasligi=Trim(Mid(sOut,linkbaslangic,(linkbitis-linkbaslangic)))
linkadresi=linkbasligi
linkadresi=Replace(linkadresi,"[url]","",1,-1,1)
linkadresi=Replace(linkadresi,"[/url]","",1,-1,1)
linkadresi="<a href="""&linkadresi&""">"&linkadresi&"</a>"
sOut=Replace(sOut,linkbasligi,linkadresi,1,-1,1)
Loop
	
	FormatText = sOut
End Function

%>