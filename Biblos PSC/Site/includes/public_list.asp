<%	
Public Sub CreateTable(sXML, sObject, iRecordsPerpage, iCurrpage, isBibliotecario)
	Dim ipageCurrent	' Current page number
	Dim iRecCount	' Number of records found
	Dim ipageCount	' Number of pages of records we have
	Dim strURL
	Dim strURLnoPage
	Dim oRs
    Dim oDOM
    Dim childNode
    Dim i, j, q
	Dim iShowTooltip
	Dim iFieldsLimit
	Dim iFields

	Dim strLastBook
	Dim bNoStock

		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
		
		If oDOM.loadXML(sXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			If Len(Request.QueryString("sort")) > 0 And Len(Request.QueryString("order")) > 0 then oRs.Sort = Request.QueryString("sort") & " " & Request.QueryString("order")
			
			iFields = oRs.Fields.Count

			if sObject = "roles" then
				iFieldsLimit = 5
				iShowTooltip = 0
			Else
				if sObject = "items" then
					iFieldsLimit = 7
					iShowTooltip = 0
				Else
					iFieldsLimit = IIf(oRs.Fields.Count > 6, 6, oRs.Fields.Count)
					iShowTooltip = IIf(iFieldsLimit = oRs.Fields.Count, 0, 1)
				End if
			End if
			
%>
<table width="85%"  border="0" align="center" cellpadding="2" cellspacing="1" style="margin: 0px 0px 0px 0px;">
  <tr>
  <td colspan="<%=iFieldsLimit + (1)%>"><div align="left" style="margin: 2px 0px 0px 0px;"><div class="m_error" align="center">&nbsp;</div></div></td>
  </tr>
  <tr>
<%			

			If not oRs.EOF then

				' Let's see what page are we looking at right now
				ipageCurrent = iCurrpage

				' Get records count
				iRecCount = oRs.RecordCount

				' Tell recordset to split records in the pages of our size
				oRs.pageSize = iRecordsPerpage

				' How many pages we've got
				ipageCount = oRs.pageCount

				' Make sure that the page parameter passed to us is within the range
				If ipageCurrent < 1 Or ipageCurrent > ipageCount Then
					' Ops - bad page number
					' let's fix it
					ipageCurrent = 1			
				End If

				' Position recordset to the page we want to see
				oRs.Absolutepage = ipageCurrent

				For i = 1 To iFieldsLimit - 1

				strURL = sObject & "_search.asp?page=" & ipageCurrent 
				strURLnoPage = sObject & "_search.asp"
%>
	<td class="h_text_table" ><div align="center" style="v-align=middle"><%=oRs(i).Name%>&nbsp;&nbsp;&nbsp;&nbsp;<a class="h_text_table" href="<%=strURL%>"></div></td>
<%					
				Next
if isBibliotecario then

%>
	<td bgcolor="#EEEEEE"></td>
<%End If%>
	<td bgcolor="#EEEEEE"></td>
  </tr>
  <tr>
<%
			Else
%>
	<td nowrap class="h_text_table" ><div align="center">No se encontraron registros.</div></td>
<%		
			End If
			j = 0
			if not oRs.EOF then strLastBook = oRs(2)
			While Not (oRs.Eof OR oRs.Absolutepage <> ipageCurrent)

			if sObject = "items" then 
				
			if strLastBook = oRs(2) AND Cint(oRs(6)) < 0 AND j <> 0 then 'este ya esta prestado
				oRs.Movenext	
				j = j + 1
			else
				if (strLastBook <> oRs(1) AND  Cint(oRs(6)) < 0) OR (j = 0 AND Cint(oRs(6)) < 0 ) then  'no hay stock
					bNoStock = true	
				Else
					bNoStock = false
				End if

				For i = 1 To iFieldsLimit - 1
%>
	<td class="m1_text" <%If bNoStock then%>bgcolor="#F2AAB1"<%End If%>><div align="center"><%
	
	If iShowTooltip = 1 Then
	
	%><a class="m2_text" onClick="myRef = window.open('<%=sObject%>_info.asp?ID=<%=oRs(0)%>','mywin',
'left=200,top=200,width=130,height=300,scrollbars=1,toolbar=0,resizable=0');
myRef.focus()" onmouseover="return escape('Haga click aqui para ver toda la información del item seleccionado.')" href="javascript:void(0);"><%=oRs(i)%></a><%
	Else
		if i = 6 and bNoStock then 
			response.write "0"
		else
			response.write oRs(i)
		end if
	End if

%></div></td>
<%	
				Next
				if not ors.eof then strLastBook = oRs(2)
	if isBibliotecario then
%>
    <td class="l_text" bgcolor="#FFFFFF"><div align="center" class="m_error_11">
	<%if bNoStock then %>No Disponible<%else%><a class="l_text" href="<%=sObject%>_borrow_insert.asp?ID=<%=oRs(0)%>">prestar</a><%end if%></div></td>
<%End if%>
	<td class="l_text" bgcolor="#FFFFFF" ><div align="center"><a class="l_text" href="<%=sObject%>_reserve_insert.asp?ID=<%=oRs(0)%>">reservar</a></div></td>
  </tr>
  <tr <%
	oRs.MoveNext
	
	If j mod 2 = 0 then
		response.write "bgcolor=""#EEEEEE"""
	End If
  
  %>>
<%

			End if
		end if
		
		j = j + 1
		Wend
q = j
For i = 1 to iRecordsPerpage - j 
%>
 	  <td class="m1_text" colspan="<%=iFieldsLimit - 1%>"><div align="left">&nbsp;</div></td>
  </tr>
  <tr <%=IIf(q mod 2 = 0, "bgcolor=""#EEEEEE""","" )%>>
<%
	q = q + 1
Next
%>
  <tr>
	  <td colspan="<%=iFieldsLimit + 1%>"><div align="left"><%=PaginationNavigation(ipageCurrent, ipageCount, strURLnoPage, 4)%></div></td>
  </tr>
</table>
<%
			Set oRs = Nothing
		Else
			'Wrong
		End If

		Set oDOM = Nothing

End Sub
%>