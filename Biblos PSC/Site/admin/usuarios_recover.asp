<%
Dim lErrNum, sErrDesc, sErrSource

Dim oUsuario
Dim iID

Set oUsuario = Server.CreateObject("Biblos_BR.cUsuario")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oUsuario.ID = iID

If oUsuario.Recover(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oUsuario = Nothing
	response.redirect "usuarios_list.asp?msgID=0&msg=Objeto recuperado con éxito."
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oUsuario = Nothing
%>
