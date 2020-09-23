<%
Dim lErrNum, sErrDesc, sErrSource

Dim oCategoria
Dim iID, iItemTipoID

Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

iItemTipoID = Server.HTMLEncode(Request.Querystring("itemtipoid"))
If len(iItemTipoID) = 0 Or isNumeric(iItemTipoID) = False Then
	iItemTipoID = "sql_injection_attempt"
End If

oCategoria.ID = iID
oCategoria.ItemTipoID = iItemTipoID

If oCategoria.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oCategoria = Nothing
	response.redirect "items_Categorias.asp?ItemTipoID=" & iItemTipoID
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oCategoria = Nothing
%>
