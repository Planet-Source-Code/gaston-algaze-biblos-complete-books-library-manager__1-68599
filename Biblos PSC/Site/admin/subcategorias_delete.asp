<%
Dim lErrNum, sErrDesc, sErrSource

Dim oSubCategoria
Dim iID, iCatID

Set oSubCategoria = Server.CreateObject("Biblos_BR.cSubCategoria")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

iCatID = Server.HTMLEncode(Request.Querystring("CatID"))
If len(iCatID) = 0 Or isNumeric(iCatID) = False Then
	iCatID = "sql_injection_attempt"
End If

oSubCategoria.ID = iID
oSubCategoria.categoriaID = iCatID

If oSubCategoria.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oSubCategoria = Nothing
	response.redirect "items_SubCategorias.asp?catID=" & iCatID
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oSubCategoria = Nothing
%>
