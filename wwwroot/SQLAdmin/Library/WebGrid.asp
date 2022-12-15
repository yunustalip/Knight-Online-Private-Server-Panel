<%
Function GetGridColumns(DataSet)
	colIndex=0
	colTitles=""
	colHeader=""
	colCount=DataSet.Fields.Count	
	For colIndex=0 To (colCount-1)
		If (colIndex>0) Then colTitles=colTitles & ","
		colHeader=DataSet(colIndex).Name
		colHeader=ClearString(colHeader)		
		colTitles=colTitles & "'" & colHeader & "'"
	Next
	GetGridColumns=colTitles
End Function 	   

Function GetGridRows(DataSet)
	rowIndex=0
	colIndex=0
	rowValues=""
	cellValue=""
	colCount=DataSet.Fields.Count	
	While Not DataSet.EOF
		If (rowIndex>0) Then rowValues=rowValues & ","
		rowValues=rowValues & "["
		For colIndex=0 To (colCount-1)
			If (colIndex>0) Then rowValues=rowValues & ","
			cellValue=DataSet(colIndex)
			cellValue=ClearString(cellValue)
			rowValues=rowValues & "'" & cellValue & "'"
		Next
		rowValues=rowValues & "]"
		rowIndex=rowIndex+1
		DataSet.MoveNext
	Wend
	GetGridRows=rowValues
End Function 

Function GetGridData(dataSet,gridName)
	rowIndex=0
	colIndex=0
	colText=""
	colName=""
	xmlString=""
	colCount=dataSet.Fields.Count 
	xmlString=xmlString & "<xml id=""" & gridName & "_XmlData"">" & vbNewLine
	xmlString=xmlString & " <XmlData>" & vbNewLine
	While Not dataSet.EOF
		xmlString=xmlString & "  <XmlRow>" & vbNewLine
		For colIndex=0 To (colCount-1)					
			colName=dataSet(colIndex).Name
			colType=dataSet(colIndex).Type
			colText=dataSet(colIndex)				
			If colType=205 Or colType=204 Or colType=128 Then  colText="<Binary>"
			colName=ClearString(colName)
			colName=Replace(colName," ","")				
			colText=ClearString(colText)			
			xmlString=xmlString & "    <" & colName & ">" & colText & "</" & colName & ">" & vbNewLine
		Next
		xmlString=xmlString & "  </XmlRow>" & vbNewLine
		dataSet.MoveNext
	Wend
	xmlString=xmlString & " </XmlData>" & vbNewLine
	xmlString=xmlString & "</xml>" & vbNewLine
	Response.Write(xmlString)
	GetGridData=xmlString
End Function  

Function WebGrid(gridName,dataSet,pageSize,rowClickEvent,doubleClickEvent,rowSelectEvent) 
	rowCount=0
	While Not dataSet.EOF
		rowCount=rowCount+1
		dataSet.MoveNext
	Wend
	If rowCount>0 Then dataSet.MoveFirst 
	colCount=dataSet.Fields.Count 
	Str = vbNewLine
	Str = Str & GetGridData(dataSet,gridName) & vbNewLine 
	Str = Str & "<table width=""100%"" height=""100%"" cellSpacing=""0"" cellPadding=""0"" border=""0""><tr><td width=""100%"" height=""100%"" align=""center"">" & vbNewLine	
	Str = Str & "<script type=""text/javascript"" language=""JavaScript"">" & vbNewLine
	Str = Str & "try" & vbNewLine
	
	Str = Str & "{"& vbNewLine
	Str = Str & "  var " & gridName & "=new Active.Controls.Grid;" & vbNewLine
	Str = Str & "  var " & gridName & "_Columns=[" & GetGridColumns(dataSet) & "];" & vbNewLine
	Str = Str & "  var " & gridName & "_Table = new Active.XML.Table;" & vbNewLine
	Str = Str & "  var " & gridName & "_XML=document.getElementById('" & gridName & "_XmlData');" & vbNewLine
	Str = Str & "  " & gridName & "_Table.setXML(" & gridName & "_XML);" & vbNewLine
	Str = Str & "  " & gridName & ".setDataModel(" & gridName & "_Table);" & vbNewLine 
	Str = Str & "  " & gridName & ".setId('" & gridName & "');" & vbNewLine	
	Str = Str & "  " & gridName & ".setStyle('width','100%');" & vbNewLine	
	Str = Str & "  " & gridName & ".setStyle('height','100%');" & vbNewLine	
	Str = Str & "  " & gridName & ".setStyle('font','menu');" & vbNewLine	
	Str = Str & "  " & gridName & ".getRowTemplate().setStyle('border-bottom', '1px solid threedlightshadow');" & vbNewLine	
	Str = Str & "  " & gridName & ".getColumnTemplate().setStyle('border-right', '1px solid threedlightshadow');" & vbNewLine	
	Str = Str & "  " & gridName & ".setRowHeaderWidth('36px');" & vbNewLine	
	Str = Str & "  " & gridName & ".setColumnHeaderHeight('21px');" & vbNewLine	
	Str = Str & "  " & gridName & ".setRowCount(" & rowCount & ");" & vbNewLine	
	Str = Str & "  " & gridName & ".setColumnCount(" & colCount & ");" & vbNewLine	  
	Str = Str & "  " & gridName + ".setModel('row', new Active.Rows.Page);" & vbNewLine
	Str = Str & "  " & gridName + ".setRowProperty('pageSize', " & pageSize & ");" & vbNewLine
	Str = Str & "  " & gridName + ".setColumnText(function(i){return " & gridName & "_Columns[i]});" & vbNewLine
	Str = Str & "  " & gridName + ".setColumnTooltip(function(i){return " & gridName & "_Columns[i]});" & vbNewLine
	If (rowClickEvent<>"")    Then Str = Str & "  " & gridName & ".setAction('click', function(src){return " & this.rowClickEvent & "(src)});" & vbNewLine
	If (doubleClickEvent<>"") Then Str = Str & "  " & gridName & ".setAction('dblclick', function(src){return " & doubleClickEvent& "(src)});" & vbNewLine
	If (rowSelectEvent<>"")   Then Str = Str & "  " & gridName & ".setAction('selectionChanged', function(src){return " & rowSelectEvent & "(src)});" & vbNewLine
	Str = Str & "  document.write(" & gridName & ");" & vbNewLine
	Str = Str & "}" & vbNewLine
	Str = Str & "  catch (error)" & vbNewLine
	Str = Str & "{"& vbNewLine
	Str = Str & "  document.write(error.description);" & vbNewLine
	Str = Str & "}" & vbNewLine		
	Str = Str & "</script>" & vbNewLine
	Str = Str & "</td></tr><tr><td>" & vbNewLine
	If (rowCount/pageSize)>1 Then 
	Str = Str & "<table style=""border-top:3px ridge;background-color: threedface"" cellSpacing=""1"" cellPadding=""0"" width=""100%"" border=""0""><tr><td height=""22"" align=""left""><table cellSpacing=""2"" cellPadding=""0"" border=""0""><tr>" & vbNewLine
	Str = Str & "<td align=""center""><img src=""Image/first.gif"" width=""22"" height=""22"" border=""0"" alt=""First page"" style=""cursor=hand;"" onclick=""js" & gridName & "GoToFirst();""></td>" & vbNewLine
	Str = Str & "<td align=""center""><img src=""Image/previous.gif"" width=""22"" height=""22"" border=""0"" alt=""Previous page"" style=""cursor=hand;"" onclick=""js" & gridName & "GoToBack();""></td>" & vbNewLine
	Str = Str & "<td align=""center""><select id=""cmb" & gridName & "Index"" style=""width:50px;height:22px"" onchange=""js" & gridName & "IndexChanged();""></select></td>" & vbNewLine
	Str = Str & "<td align=""center""><img src=""Image/next.gif"" width=""22"" height=""22"" border=""0"" alt=""Next page"" style=""cursor=hand;"" onclick=""js" & gridName & "GoToNext();""></td>" & vbNewLine
	Str = Str & "<td align=""center""><img src=""Image/last.gif"" width=""22"" height=""22"" border=""0"" alt=""Last page"" style=""cursor=hand;"" onclick=""js" & gridName & "GoToLast();""></td></tr></table>" & vbNewLine
	Str = Str & "<script>function js" & gridName & "GoToFirst(){js" & gridName & "GoToPage(0)}" & vbNewLine
	Str = Str & "function js" & gridName & "GoToBack(){var pageIndex=" & gridName & ".getRowProperty('pageNumber');if(pageIndex>0)pageIndex=pageIndex-1;js" & gridName & "GoToPage(pageIndex);}" & vbNewLine
	Str = Str & "function js" & gridName & "GoToNext(){var pageCount=" & gridName & ".getRowProperty('pageCount');var pageIndex=" & gridName & ".getRowProperty('pageNumber');if(pageIndex<pageCount-1)pageIndex=pageIndex+1;js" & gridName & "GoToPage(pageIndex);}" & vbNewLine
	Str = Str & "function js" & gridName & "GoToLast(){var pageCount=" & gridName & ".getRowProperty('pageCount');js" & gridName & "GoToPage(pageCount-1);}" & vbNewLine
	Str = Str & "function js" & gridName & "GoToPage(pageIndex){document.getElementById(""cmb" & gridName & "Index"").options[pageIndex].selected=true;" & gridName & ".setRowProperty('pageNumber',pageIndex);" & gridName & ".refresh();}" & vbNewLine
	Str = Str & "function js" & gridName & "IndexChanged(){var obj=document.getElementById('cmb" & gridName & "Index');var opt=obj.options[obj.selectedIndex];var val=opt.value-1;js" & gridName & "GoToPage(val);}" & vbNewLine
	Str = Str & "function js" & gridName & "AppendOptions(){js" & gridName & "ClearOptions();var pageCount=" & gridName & ".getRowProperty('pageCount');var obj=document.getElementById(""cmb" & gridName & "Index"");for(i=1;i<pageCount+1;i++){var opt=document.createElement(""Option"");opt.text=i;opt.value=i;obj.options.add(opt);}}" & vbNewLine
	Str = Str & "function js" & gridName & "ClearOptions(){var obj=document.getElementById('cmb" & gridName & "Index');for(i=obj.length-1;i>=0;i){if(obj.options[i].selected){obj.remove(i);}}}js" & gridName & "AppendOptions();js" & gridName & "GoToPage(0)" & vbNewLine
	Str = Str & "</script></td></tr></table></td>" & vbNewLine
	End If 
	Str = Str & "</tr></table>" & vbNewLine	
	WebGrid = Str	
End Function 	   

Function ClearString(val)	
	If IsNull(val) Then val=""
	val = Replace(val, "<","&lt;")
	val = Replace(val, ">","&gt;")
	val = Replace(val, "&", "&amp;")
	val = Replace(val, "\", "\\")
	val = Replace(val, """", "\""")
	val = Replace(val, vbCrLf , " ")
	val = Replace(val, vbNewLine, " ") 
	ClearString=val
End Function
%>