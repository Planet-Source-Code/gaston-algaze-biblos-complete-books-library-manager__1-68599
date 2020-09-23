<%
Dim lErrNum, sErrDesc, sErrSource

Dim oEditorial
Dim iID, iCatID

Set oEditorial = Server.CreateObject("Biblos_BR.cEditorial")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oEditorial.ID = iID

If oEditorial.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oEditorial = Nothing
	response.redirect "items_editoriales.asp"
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oEditorial = Nothing
%>
