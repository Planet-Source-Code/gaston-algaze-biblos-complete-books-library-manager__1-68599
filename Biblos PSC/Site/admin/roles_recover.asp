<%
Dim lErrNum, sErrDesc, sErrSource

Dim oRol
Dim iID

Set oRol = Server.CreateObject("Biblos_BR.cRol")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oRol.ID = iID

If oRol.Recover(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oRol = Nothing
	response.redirect "roles_list.asp?msgID=0&msg=Objeto recuperado con éxito."
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oRol = Nothing
%>
