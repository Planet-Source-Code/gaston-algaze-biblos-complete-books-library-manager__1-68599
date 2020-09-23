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
	Dim iShowDeleted
	Dim iShowTooltip
	Dim iFieldsLimit
	Dim iFields

		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
		
		If oDOM.loadXML(sXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			If Len(Request.QueryString("sort")) > 0 And Len(Request.QueryString("order")) > 0 then oRs.Sort = Request.QueryString("sort") & " " & Request.QueryString("order")
			
			iFields = oRs.Fields.Count

			iFieldsLimit = 7
			iShowTooltip = 0
			
%>
<table width="85%"  border="0" align="center" cellpadding="2" cellspacing="1" style="margin: 0px 0px 0px 0px;">
  <tr>
  <td colspan="<%=iFieldsLimit + (1 - iShowDeleted)%>"><div align="left" style="margin: 2px 0px 0px 0px;"><div class="m_error" align="center">&nbsp;</div></div></td>
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

				strURL = sObject & "_list.asp?sort=" & oRs(i).Name & "&order=" & IIf(Request.QueryString("order") = "ASC" OR Request.QueryString("order") = "", "DESC", "ASC") & "&page=" & ipageCurrent 
				strURLnoPage = sObject & "_list.asp?sort=" & oRs(i).Name & "&order=" & Request.QueryString("order") 
%>
	<td class="h_text_table" ><div align="center" style="v-align=middle"><a class="h_text_table" href="<%=strURL%>"><%=oRs(i).Name%></a>&nbsp;&nbsp;&nbsp;&nbsp;<a class="h_text_table" href="<%=strURL%>"><img align="absmiddle" src=<%
	If Request.QueryString("sort") = oRs(i).Name then
		If Request.QueryString("order") = "ASC" then
			response.write "/images/sort_ascending.gif"
		Else 
			response.write "/images/sort_descending.gif"
		End If
	Else
		response.write "/images/sort_none.gif"
	End If
	
	%> WIDTH="15" HEIGHT="16" BORDER="0" ALT=""></a></div></td>
<%					
				Next
%>
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td bgcolor="#EEEEEE">&nbsp;</td>
    <td bgcolor="#EEEEEE">&nbsp;</td>
  </tr>
  <tr>
<%
			Else
%>
	<td nowrap class="h_text_table" ><div align="center">No se encontraron registros.</div></td>
<%		
			End If
			j = 0
			While Not (oRs.Eof OR oRs.Absolutepage <> ipageCurrent)
				For i = 1 To iFieldsLimit - 1
%>
	<td class="m1_text"><div align="center"><%
	
	If iShowTooltip = 1 Then
	
	%><a class="m2_text" onClick="myRef = window.open('<%=sObject%>_info.asp?ID=<%=oRs(0)%>','mywin',
'left=200,top=200,width=130,height=300,scrollbars=1,toolbar=0,resizable=0');
myRef.focus()" onmouseover="return escape('Haga click aqui para ver toda la información del item seleccionado.')" href="javascript:void(0);"><%=oRs(i)%></a><%
	Else
		Response.write oRs(i)
	End if

%></div></td>
<%	
				Next
	if isBibliotecario and formatdate(oRs("fecha_reserva")) = formatdate(date) then
%>
	<td class="l_text" bgcolor="#FFFFFF"><div align="center"><a class="l_text" href="items_borrow_insert.asp?ID=<%=oRs("itemID")%>&usuarioID=<%=oRs("usuarioID")%>&reservaID=<%=oRs("ID")%>">prestar</a></div></td>
<%Else%>
	<td class="l_text" bgcolor="#FFFFFF">&nbsp;</td>
<%End If%>
	<td class="l_text" bgcolor="#FFFFFF"><div align="center"><a class="l_text" href="<%=sObject%>_update.asp?ID=<%=oRs(0)%>">editar</a></div></td>
    <td class="l_text" bgcolor="#FFFFFF"><div align="center"><a class="l_text" href="<%=sObject%>_cancel.asp?ID=<%=oRs(0)%>" onclick="return confirm('¿Está seguro que desea cancelar la reserva?')">cancelar</a></div></td>  
  </tr>
  <tr <%
	oRs.MoveNext
				
	If j mod 2 = 0 then
		response.write "bgcolor=""#EEEEEE"""
	End If
  
  %>>
<%
				j = j + 1
			Wend
q = j
For i = 1 to iRecordsPerpage - j 
%>
 	  <td class="m1_text" colspan="<%=iFieldsLimit - 1%>"><div align="left">&nbsp;</div></td>
	  <td class="l_text" bgcolor="#FFFFFF"><div align="center">&nbsp;</div></td>
      <td class="l_text" bgcolor="#FFFFFF"><div align="center">&nbsp;</div></td>  
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