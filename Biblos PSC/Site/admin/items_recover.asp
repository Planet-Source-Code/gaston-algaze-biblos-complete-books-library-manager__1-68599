<%
Dim lErrNum, sErrDesc, sErrSource

Dim oUbicacion
Dim iID

Set oUbicacion = Server.CreateObject("Biblos_BR.cUbicacion")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oUbicacion.ID = iID

If oUbicacion.Recover(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oUbicacion = Nothing
	response.redirect "items_list.asp?msg=Objeto recuperado con éxito."
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oUbicacion = Nothing
%>
