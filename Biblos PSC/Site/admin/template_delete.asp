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

If oUbicacion.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oUbicacion = Nothing
	response.redirect "template_list.asp?msgID=0&msg=Objeto eliminado con éxito."
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oUbicacion = Nothing
%>
