<%
Dim lErrNum, sErrDesc, sErrSource

Dim oPrestamo
Dim iID, iCatID

Set oPrestamo = Server.CreateObject("Biblos_BR.cPrestamo")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oPrestamo.ID = iID

If oPrestamo.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oPrestamo = Nothing
	response.redirect "items_Prestamos.asp"
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oPrestamo = Nothing
%>
