<%
Dim lErrNum, sErrDesc, sErrSource

Dim oUbicacion
Dim iID, iCatID

Set oUbicacion = Server.CreateObject("Biblos_BR.cUbicacion")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oUbicacion.ID = iID

If oUbicacion.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oUbicacion = Nothing
	response.redirect "items_Ubicaciones.asp"
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oUbicacion = Nothing
%>
