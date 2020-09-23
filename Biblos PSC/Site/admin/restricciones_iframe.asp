<%Option Explicit%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<link href="/styles/styles.css" rel="stylesheet" type="text/css">
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/admin_roles.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
</HEAD>
<%
Dim lErrNum, sErrDesc, sErrSource

Dim oRol
Dim iID
Dim strXML
Dim strXMLRestricciones

iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If
Set oRol = Server.CreateObject("Biblos_BR.cRol")
If oRol.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then
	If oRol.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
%>
<table width="505"  border="1" cellpadding="2" cellspacing="2" bordercolor="#E2E2E2">
  <tr>
	<td  class="h_text_table"><div align="center">Tabla</div></td>
	<td  class="h_text_table"><div align="center">Campo</div></td>
	<td  class="h_text_table"><div align="center">Operacion</div></td>
	<td  class="h_text_table"><div align="center">Valor</div></td>
	<td  class="h_text_table" bgcolor="#EEEEEE"><div align="center">&nbsp;</div></td>
	<td  class="h_text_table" bgcolor="#EEEEEE"><div align="center">&nbsp;</div></td>
  </tr>
  <% 
	If oRol.GetRestricciones(session("userID"), strXMLRestricciones, lErrNum, sErrDesc, sErrSource) then
		DrawRestrictions strXMLRestricciones, iID
	Else
		response.write "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End if
  %>
</table>
<BR>
							  <a class="m2_text" onClick="myRef = window.open('restricciones_insert.asp?ID=<%=iID%>','mywin',
'left=200,top=200,width=600,height=1,scrollbars=0,toolbar=0,resizable=0,menubar=0');
myRef.focus()" onmouseover="return escape('Haga click aqui para adicionar una nueva restricción.')" href="javascript:void(0);"><IMG SRC="/images/nueva.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Nueva Restricci&oacute;n"></a>
<%
	End if
End if
Set oRol = nothing
%>
</BODY>
</HTML>

