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

Dim oFicha
Dim strXML
Dim iID

Set oFicha = Server.CreateObject("Biblos_BR.cFicha")


iID = Server.HTMLEncode(Request.Querystring("ID"))
If len(iID) = 0 Or isNumeric(iID) = False Then
	iID = "sql_injection_attempt"
End If

oFicha.Search session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource
oFicha.Read session("userID"), strXML, lErrNum, sErrDesc, sErrSource
With oFicha
%>
</head>

<body bgcolor="#FFFFFF">
	<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td width="100%" valign="middle">
<!-- ACA COMIENZA EL BODY -->
					<table width="90%"  border="0" align="center" cellpadding="1" cellspacing="0" >
					  <tr>
						<td width="200" class="h_text"><div align="right">Username</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Username%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Nombre</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Nombre%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Apellido</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Apellido%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">E-Mail</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Mail%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">DNI</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.DNI%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Matricula</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Matricula%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Fecha de Nacimiento</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Fecha_Nacimiento%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Calle</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Domicilio_Calle%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Nro.</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Domicilio_Nro%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Piso</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Domicilio_Piso%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Unidad</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Domicilio_Unidad%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Cod. Postal</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Domicilio_Cod_Postal%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Tel. 1</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Tel1%></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Tel. 2</div></td>
                          <td width="200"><div align="left" class="m1_text">&nbsp;&nbsp;<%=.Tel2%></div></td>
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