<%
Option Explicit

Dim lErrNum, sErrDesc, sErrSource

Dim oLink
Dim iID, iCopy
Dim strTitle

strTitle = "Eliminación de Link"

Set oLink = Server.CreateObject("Biblos_BR.cLink")

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

iCopy = Server.HTMLEncode(Request.Querystring("Copy"))

oLink.ID = iID

if len(iCopy) > 0 then 
	If oLink.DeleteCopy(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oLink = Nothing
		response.redirect "Links_list.asp?msg=Objeto eliminado con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
else
	If oLink.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oLink = Nothing
		response.redirect "Links_list.asp?msgID=0&msg=Objetos eliminados con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
End if
%>
