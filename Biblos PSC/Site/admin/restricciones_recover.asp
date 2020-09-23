<%
Dim lErrNum, sErrDesc, sErrSource

Dim oRestriccion
Dim iID, iIDback

Set oRestriccion = Server.CreateObject("Biblos_BR.cRestriccion")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

iIDback = Server.HTMLEncode(Request.Querystring("IDback"))
If len(iIDback) = 0 Or isNumeric(iIDback) = False Then
	iIDback = "sql_injection_attempt"
End If

oRestriccion.ID = iID

If oRestriccion.Recover(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oRestriccion = Nothing
	response.redirect "restricciones_iframe.asp?ID=" & iIDback
Else
	response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
End If

Set oRestriccion = Nothing
%>
