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
<script language='javascript' src="/includes/js/basic_functions.js"></script>
<script language='javascript' src="/includes/js/validate.js"></script>
<script language='javascript' src="/includes/js/calendar/popcalendar.js"></script>
<script language='javascript' src="/includes/js/calendar/lw_layers.js"></script>
<script language='javascript' src="/includes/js/calendar/lw_menu.js"></script>
<script type="text/javascript">

<!--

function validate_form ()
{
		valid = checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');
		if (valid != false){
			if ( document.myform.cboRol.value == "-1" )
			{
					alert ( "Por favor seleccione un rol." );
					document.myform.cboRol.focus();
					valid=false;
			}
		};
		return valid;

}

//-->

</script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
strTitle = "usuarios_insert"
if Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	Dim oUsuario

	Set oUsuario = Server.CreateObject("Biblos_BR.cUsuario")

	With oUsuario
		.username = Request.Form("username")
		.password = Request.Form("password")
		.nombre = Request.Form("nombre")
		.apellido = Request.Form("apellido")
		.mail = Request.Form("mail")
		.dni = Request.Form("dni")
		.matricula = Request.Form("matricula")			
		.fecha_nacimiento = FormatDate(Request.Form("fecha_nacimiento"))
		.domicilio_calle = Request.Form("calle")
		.domicilio_nro = Request.Form("nro")
		.domicilio_piso = Request.Form("piso")
		.domicilio_unidad = Request.Form("unidad")
		.domicilio_cod_postal = Request.Form("cod_postal")
		.tel1 = Request.Form("tel1")
		.tel2 = Request.Form("tel2")
		.rolID = Request.Form("cboRol")
	End With

	If oUsuario.Add(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oUsuario = Nothing
		response.redirect "usuarios_list.asp?msgID=0&msg=Objeto insertado con �xito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If

	Set oUsuario = Nothing
Else
	Dim oRol
	Dim strXML
	Dim oDOM
	Dim oRs

	Set oRol = Server.CreateObject("Biblos_BR.cRol")
%>
</head>

<body bgcolor="#FFFFFF" style="background-image: url(/images/f-l.gif); background-repeat: repeat-y; background-position: left; margin: 0px 0px 0px 0px;">
	<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td width="100%" style="height:120px" valign="top">
			<table width="100%" style="background-image: url(/images/t-dr.gif); background-repeat: repeat-x; height:120px" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="673" style="background-image: url(/images/t-l.gif); background-repeat: repeat-x; height:120px" valign="top">
<!-- ACA COMIENZA EL HEADER -->
<%
		Header_Admin()
%>					
<!-- ACA TERMINA EL HEADER -->				
				</td>
				<td width="100%" style="background-image: url(/images/t-r.gif); background-repeat: no-repeat; background-position: right;" valign="top"><div><img  src="/images/spacer.gif" alt="" width="93" height="1"  border="0"></div></td>
			  </tr>
			</table>
		</td>
	  </tr>
	  <tr>
		<td width="100%" valign="top">
			<table width="100%" height="100%" border="0" align="left" cellpadding="0" cellspacing="0" >
			  <tr>
				<td align="center" valign="top" style="height: 1px">
				<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" align="center">
                  <tr>
                    <td width="100" valign="top">
<%
	MenuBar_Admin session("rolID"), session("userID")
%>
					</td>
                    <td width="800" align="center" valign="top">
<!-- ACA COMIENZA EL BODY -->

<!-- ACA COMIENZA EL FORM-->
					<form action="usuarios_insert.asp" method="POST" name="myform" onsubmit="return validate_form();">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5"><div align="left" style="margin: 2px 0px 0px 0px;">&nbsp;</div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Username</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="username" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_username">
                          </div></td>
						  <td width="300px"><div id="usernameError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Password</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="password" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_password">
                          </div></td>
						  <td width="300px"><div id="passwordError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Nombre</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="nombre" value=""><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_nombre">
                          </div></td>
						  <td><div id="nombreError"></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Apellido</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="apellido" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_apellido">
                          </div></td>
						  <td width="300px"><div id="apellidoError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">E-Mail</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="mail" value=""><font class="m_text">*</font>
						<input type="hidden" name="@email_NoBlank_mail">
                          </div></td>
						  <td><div id="mailError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">DNI</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="dni" value=""><font class="m_text">*</font>
						<input type="hidden" name="@dni_NoBlank_dni">
                          </div></td>
						  <td><div id="dniError"></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Matricula / Legajo</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="matricula" value=""  maxlength="10"><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_matricula">
                          </div></td>
						  <td width="300px"><div id="matriculaError"></div></td>
					  </tr>
					  <tr>
						<td width="220" class="h_text"><div align="right">Fecha de Nacimiento</div></td>
                          <td width="240"><div align="left">
                            <input type="text" name="fecha_nacimiento" value="" OnFocus="this.blur();">
						<input type="hidden" name="@ _NoBlank_fecha_nacimiento">&nbsp;<a href="javascript:void(0);" onClick="popUpCalendar(this, myform.fecha_nacimiento, 'dd/mm/yyyy');"><img src="/images/show-calendar.gif" width="24" height="22" border="0" align="absmiddle"></a>&nbsp;<font class="m_text">*</font></div></td>
						  <td><div id="fecha_nacimientoError"></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Calle</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="calle" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_calle">
                          </div></td>
						  <td width="300px"><div id="calleError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Nro.</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="nro" value=""><font class="m_text">*</font>
						<input type="hidden" name="@number_NoBlank_nro">
                          </div></td>
						  <td><div id="nroError"></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Piso</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="piso" value="">
                          </div></td>
						  <td width="300px"></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Unidad</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="unidad" value="">
                          </div></td>
						  <td></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Cod. Postal</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="cod_postal" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_cod_postal">
                          </div></td>
						  <td width="300px"><div id="cod_postalError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Tel. 1</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="tel1" value=""><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_tel1">
                          </div></td>
						  <td><div id="tel1Error"></div></td>
					  </tr>
					   <tr>
						<td width="200" class="h_text"><div align="right">Tel. 2</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="tel2" value="">
						    <input type="hidden" name="@ _Blank_tel2">
                          </div></td>
						  <td width="300px"><div id="tel2Error"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Rol</div></td>
                          <td width="200"><div align="left">
						  <select name="cboRol" id="cboRol" style="width:100px;">
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  if oRol.Search(session("userID"), strXMl, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXMl) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=" & oRs("ID") & ">" & oRs("Descripcion") &  "</option>"
									oRs.MoveNext
								Wend
							End if
						  End if
						  Set oRol = nothing
						  Set oDOM = nothing
						  Set oRs = nothing
						  %>
							  </select><font class="m_text">&nbsp;*</font>
                          </div></td>
						  <td width="300px"></td>
					  </tr>
					  <tr>
						   <td width="200" class="h_text"><div align="right"></div></td>
                          <td width="200"><div align="left"><INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">&nbsp;&nbsp;<A href="javascript:void(0);" onclick="history.back();return false"><IMG SRC="/images/volver.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
						  <td><IMG SRC="/images/spacer.gif" WIDTH="240" HEIGHT="1" BORDER="0" ALT=""></td>
					  </tr>
					  <tr>
					  <td colspan="5"><div align="center" style="margin: 2px 0px 0px 0px;"><BR><font class="m_text">Los campos marcados con (*) son obligatorios.</font></div></td>
					  </tr>
					</table>
					</td>
                  </tr>
                </table>
				<INPUT TYPE="hidden" name="submit" value="true">
				</form>

<!-- ACA TERMINA EL FORM -->

<!-- ACA TERMINA EL BODY -->
				</td>
			  <td width="92" align="center" valign="top" style="background-image: url(/images/t-r-line.gif); background-repeat: repeat-y;width:92px"></td>
			  </tr>
			</table>
		</td>
	  </tr>
	  <tr>
		<td width="100%" style="vertical-align: top;">
<!-- ACA COMIENZA EL FOOTER -->
<%
		Footer_Admin()
%>			
<!-- ACA TERMINA EL FOOTER -->		
		</td>
	  </tr>
	</table>
</body>
</html>
<%
End if
%>