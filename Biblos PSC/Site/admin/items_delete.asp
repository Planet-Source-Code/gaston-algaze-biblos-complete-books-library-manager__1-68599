<%
Option Explicit

Dim lErrNum, sErrDesc, sErrSource

Dim oItem
Dim iID, iCopy
Dim strTitle

strTitle = "Eliminación de Item"

Set oItem = Server.CreateObject("Biblos_BR.cItem")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

iCopy = Server.HTMLEncode(Request.Querystring("Copy"))

oItem.ID = iID

if len(iCopy) > 0 then 
	If oItem.DeleteCopy(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oItem = Nothing
		response.redirect "items_list.asp?msgID=0&msg=Objeto eliminado con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
else
	If oItem.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oItem = Nothing
		response.redirect "items_list.asp?msgID=0&msg=Objetos eliminados con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
End if
%>
