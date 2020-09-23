<%
Dim lErrNum, sErrDesc, sErrSource

Dim oReserva
Dim iID, iCatID

Set oReserva = Server.CreateObject("Biblos_BR.cReserva")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oReserva.ID = iID

If oReserva.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oReserva = Nothing
	response.redirect "items_reserve_list.asp?msgID=0&msg=Reserva%20cancelada%20con%20éxito."
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oReserva = Nothing
%>
