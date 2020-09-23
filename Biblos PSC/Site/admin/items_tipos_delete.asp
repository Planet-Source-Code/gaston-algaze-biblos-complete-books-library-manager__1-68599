<%
Dim lErrNum, sErrDesc, sErrSource

Dim oItemTipo
Dim iID, iCatID

Set oItemTipo = Server.CreateObject("Biblos_BR.cItemTipo")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oItemTipo.ID = iID

If oItemTipo.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oItemTipo = Nothing
	response.redirect "items_Tipos.asp"
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oItemTipo = Nothing
%>
