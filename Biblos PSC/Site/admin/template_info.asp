<%Option Explicit%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="content-type" content="text/html;charset=ISO-8859-1">
<meta http-equiv="Content-Script-Type" content="text/javascript; charset=iso-8859-1">
<title>:: M&oacute;dulo <%=session("rol")%> - Sistema Biblos ::</title>
<link rel="icon" href="/favicon.ico" type="image/x-icon">
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon">
<link href="/styles/styles.css" rel="stylesheet" type="text/css">
<script src="/includes/js/basic_functions.js"></script>
<script src="/includes/js/validate.js"></script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
Dim lErrNum, sErrDesc, sErrSource

Dim oUbicacion
Dim strXML
Dim iID

Set oUbicacion = Server.CreateObject("Biblos_BR.cUbicacion")


iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oUbicacion.Search session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource
oUbicacion.Read session("userID"), strXML, lErrNum, sErrDesc, sErrSource
With oUbicacion
%>
</head>

<body bgcolor="#FFFFFF">
	<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td width="100%" valign="middle">
<!-- ACA COMIENZA EL BODY -->
					<table width="90%"  border="0" align="center" cellpadding="1" cellspacing="0" >
					  <tr>
						<td width="200" class="h_text"><div align="right">Título</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Titulo%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Descripcion</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Descripcion%></div></td>
					  </tr>
					  <tr>
						   <td align="center" colspan="2" class="h_text"><BR><A href="javascript:void(0);" onclick="window.close();return false"><IMG SRC="/images/cerrar.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></td>
					  </tr>
					</table>
<!-- ACA TERMINA EL BODY -->
		</td>
	  </tr>
	</table>
</body>
</html>
<%
	End With
%>