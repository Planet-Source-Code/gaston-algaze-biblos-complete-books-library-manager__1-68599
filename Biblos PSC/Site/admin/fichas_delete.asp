<%
Dim lErrNum, sErrDesc, sErrSource

Dim oFicha
Dim iID

Set oFicha = Server.CreateObject("Biblos_BR.cFicha")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oFicha.ID = iID

If oFicha.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oFicha = Nothing
	response.redirect "Fichas_list.asp?msgID=0&msg=Objeto eliminado con éxito."
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oFicha = Nothing
%>
