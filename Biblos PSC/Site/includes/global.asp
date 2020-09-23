<%
Dim strTitle
Dim strCurrentUser
Dim strConnString
Dim strRSSFeed

strTitle="T&iacute;tulo de Secci&oacute;n"
strCurrentUser=session("username")

'modificar esto desde el administrador, y guardarlo en la base de datos.
strRSSFeed = "http://localhost/RSS/RSSFeed.xml" '"http://weblog.educ.ar/educacion-tics/index.xml"
Const strLogoPath="/images/logo_new_2.gif"
Const RECORDS_PER_PAGE 	= 10
'Fin de cosas que se configuran desde el admin

Function RecordsetFromXMLDocument(XMLDOMDocument)
Dim oRecordset
	
	Set oRecordset = Server.CreateObject("ADODB.Recordset")
	oRecordset.Open XMLDOMDocument 'pass the DOM Document instance as the Source argument
	Set RecordsetFromXMLDocument = oRecordset  'return the recordset
	Set oRecordset = Nothing

End Function


Function PadDigits(n, totalDigits) 
	if totalDigits > len(n) then 
		PadDigits = String(totalDigits-len(n),"0") & n 
	else 
		PadDigits = n 
	end if 
End Function 


Function CheckString(sValue)
Dim i
Dim invalidList
    i = 0 
 
    ' set up a list of unacceptable characters 
    ' this includes spaces, dashes and underscores 
    ' you can leave these out of the list 
    ' you may need to add other characters, e.g. copied from MSWord 
 
    invalidList = "*,<.>?;:'@#~]}[{=+)(*&^%$£!`¬| -_%" 
 
    ' check for " which can't be inside the string 
 
    if Instr(sValue,chr(34))>0 then 
       sValue =  Replace(sValue, chr(34),"&quot")
    else 
        ' loop through, making sure no characters 
        ' are in the 'reserved characters' list 
 
        for i = 1 to len(invalidList) 
            if Instr(sValue,Mid(invalidList,i,1))>0 then 
                 sValue = Replace(sValue, Asc(Mid(invalidList,i,1)),"%" & Hex(Asc(Mid(invalidList,i,1)))) 
            end if 
        next 
    end if 

	CheckString = sValue

	sValue = "%" & Hex(asc("%"))
End Function

'Format SQL Query function
Function FormatSQLInput(sValue)

	'Remove malicious characters from links and images
	sValue = Replace(sValue, "<", "&lt;")
	sValue = Replace(sValue, ">", "&gt;")
	sValue = Replace(sValue, "[", "[")
	sValue = Replace(sValue, "]", "]")
	sValue = Replace(sValue, """", "", 1, -1, 1)
	sValue = Replace(sValue, "=", "=", 1, -1, 1)
	sValue = Replace(sValue, "'", "''", 1, -1, 1)
	sValue = Replace(sValue, "select", "select", 1, -1, 1)
	sValue = Replace(sValue, "join", "join", 1, -1, 1)
	sValue = Replace(sValue, "union", "union", 1, -1, 1)
	sValue = Replace(sValue, "where", "where", 1, -1, 1)
	sValue = Replace(sValue, "insert", "insert", 1, -1, 1)
	sValue = Replace(sValue, "delete", "delete", 1, -1, 1)
	sValue = Replace(sValue, "update", "update", 1, -1, 1)
	sValue = Replace(sValue, "like", "like", 1, -1, 1)
	sValue = Replace(sValue, "drop", "drop", 1, -1, 1)
	sValue = Replace(sValue, "create", "create", 1, -1, 1)
	sValue = Replace(sValue, "modify", "modify", 1, -1, 1)
	sValue = Replace(sValue, "rename", "rename", 1, -1, 1)
	sValue = Replace(sValue, "alter", "alter", 1, -1, 1)
	sValue = Replace(sValue, "cast", "cast", 1, -1, 1)

	'Return
	FormatSQLInput = sValue
End Function

Function FormatDate(sDate)
	FormatDate = DatePart("YYYY",sDate)

	If len(DatePart("m", sDate)) = 1 then 
		FormatDate = FormatDate & "0" & DatePart("m", sDate)
	else
		FormatDate = FormatDate & DatePart("m", sDate)
	end if
		
	If len(DatePart("d", sDate)) = 1 then 
		FormatDate = FormatDate & "0" & DatePart("d", sDate)
	else
		FormatDate = FormatDate & DatePart("d", sDate)
	end if
End Function

Function IIf(psdStr, trueStr, falseStr)
	if psdStr then
		IIf = trueStr
	else 
		IIf = falseStr
	end if
end function

Function PaginationNavigation(ByVal pLngpage, ByRef pLngpageCount, ByVal pStrURL, ByRef pLngMax)
	
	Dim lStrNav
	Dim lLngFirst
	Dim lLngLast
	Dim lLngIndex
	
	If pLngpageCount < 2 Then Exit Function
	If Not IsNumeric(pLngpage) Or pLngpage = "" Then pLngpage = 1
	pLngpage = CLng(pLngpage)
	If pLngpage < 1 Then pLngpage = 1
	
	lLngFirst = pLngpage - (pLngMax \ 2)
	lLngLast = lLngFirst + (pLngMax - 1)
	If lLngFirst < 1 Then
		lLngFirst = 1
		lLngLast = pLngMax
	End If
	
	If lLngLast > pLngpageCount Then
		lLngLast = pLngpageCount
		lLngFirst = lLngLast - (pLngMax - 1)
	End If
	
	If lLngFirst < 1 Then lLngFirst = 1
	
	lStrNav = "<div class=""l_text_page"">"
	
	If InStr(1, pStrURL, "?") = 0 Then
		pStrURL = pStrURL & "?page="
	Else
		If Right(pStrURL, 1) = "?" Then
			pStrURL = pStrURL & "page="
		Else
			pStrURL = pStrURL & "&page="
		End If
	End If
	
	If Not lLngFirst = 1 Then
		lStrNav = lStrNav & "<A class=""l_text_page"" href=""" & pStrURL & "1"">1</A> ... "
	End If
	
	For lLngIndex = lLngFirst To lLngLast
		If lLngIndex = pLngpage Then lStrNav = lStrNav & "<font class=""m_error"">["
		lStrNav = lStrNav & "<A class=""l_text_page"" href=""" & pStrURL & lLngIndex & """>" & lLngIndex & "</A>"
		If lLngIndex = pLngpage Then lStrNav = lStrNav & "]</font>"
		lStrNav = lStrNav & " "
	Next
	
	If Not lLngLast = pLngpageCount Then
		lStrNav = lStrNav & " ... <A class=""l_text_page"" href=""" & pStrURL & pLngpageCount & """>" & pLngpageCount & "</A> "
	End If
	
	PaginationNavigation = lStrNav & "</div>"
	
End Function

Sub GetRss(sRSS, iMaxNews)
Dim extURL
Dim xmlDoc
Dim channelNodes
Dim entry
Dim item
Dim itemNodes
Dim strChannelTitle, strChannelDescription, strChannelLink
Dim strItemTitle, strItemDescription, strItemLink
Dim arrItemTitle, arrItemDescription, arrItemLink
Dim i, a

	'Direccion del feed RSS
	extURL = sRSS

	set xmlDoc = createObject("Msxml.DOMDocument")
	xmlDoc.async = false
	xmlDoc.setProperty "ServerHTTPRequest", true
	xmlDoc.load(extURL)

	If (xmlDoc.parseError.errorCode <> 0) then
		Response.Write "XML error: " & xmlDoc.parseError.reason
	Else

		set channelNodes = xmlDoc.selectNodes("//channel/*")

		for each entry in channelNodes
			if entry.tagName = "title" then
				strChannelTitle = entry.text
			elseif entry.tagName = "description" then
				strChannelDescription = entry.text
			elseif entry.tagName = "link" then
				strChannelLink = entry.text
			end if
		next

		response.write "<div class=""l1_text_bolder"">Noticias desde: " & strChannelTitle & "</div>"
		
		set itemNodes = xmlDoc.selectNodes("//item/*")
		i = 0
		For each item in itemNodes
			if item.tagName = "title" then
				strItemTitle = strItemTitle & item.text & "#%#"
			elseif item.tagName = "link" then
				strItemLink = strItemLink & item.text & "#%#"
			elseif item.tagName = "description" then
				strItemDescription = strItemDescription & item.text & "#%#"
				i = i + 1
				if i = iMaxNews then exit for
			end if
		next

		arrItemTitle = split(strItemTitle,"#%#")
		arrItemLink = split(strItemLink,"#%#")
		arrItemDescription = split(strItemDescription,"#%#")

		response.write "<ul style=""margin-top:10px; margin-right:px"">"
			for a = 0 to UBound(arrItemTitle) - 1
				response.write "<li>"
				response.write "<a class = ""h_text"" href='" & arrItemLink(a) & "'>" & arrItemTitle(a) & "</a> "
					if strItemDescription <> "" then
						response.write "<div align=""left"" class=""m_text"">" & arrItemDescription(a) & "</div>"
					end if
				response.write "</li>"
			next
		response.write "</ul>"

		set channelNodes = nothing
		set itemNodes = nothing

	End If

End Sub
%>