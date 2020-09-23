<%
Dim lErrNum, sErrDesc, sErrSource

Dim oItem
Dim iID, iCatID

Set oItem = Server.CreateObject("Biblos_BR.cItem")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oItem.ID = iID

If oItem.AddCopy(session("userID"), lErrNum, sErrDesc, sErrSource) Then
	Set oItem = Nothing
	response.redirect "items_list.asp?msgID=0&msg=Objeto copiado con éxito."
Else
	%>
	  <script language = JavaScript>
	   parent.location.href = "/admin/error.asp?delete=1&title=Error&msg=<%=sErrDesc%>";
	  </script>
	<%
End If

Set oItem = Nothing
%>
