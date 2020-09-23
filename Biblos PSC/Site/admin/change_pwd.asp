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
		valid = true;

        if ( document.myform.Contraseña°Nueva.value != document.myform.Contraseña°Nueva°Confirmación.value )
        {
                alert ( "La confirmación de la contraseña nueva no concuerda." );
				document.myform.Contraseña°Nueva°Confirmación.focus();
                valid = false;
        }
		
		if ( valid == true ) {
			return checkform(myform, '#ffcccc', '#ffffff', true, true, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');
		}else{
			return valid;
		}
}

//-->

</script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
strTitle = "Cambio de Contraseña"
Dim lErrNum, sErrDesc, sErrSource
Dim iID
Dim oRs 
Dim oUsuario
Dim oDOM

if len(request.form("submit")) > 0 then
	
	Set oUsuario = Server.CreateObject("Biblos_BR.cUsuario")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	Set oDOM = Server.CreateObject("MSXML2.DOMDocument")

	oRs.Fields.Append "userID", adBSTR
	oRs.Fields.Append "pwdold", adBSTR
	oRs.Fields.Append "pwdnew1", adBSTR
	oRs.Fields.Append "pwdnew2", adBSTR

	
	oRs.Open
	
	'categorias
	oRs.AddNew
	oRs(0) = Request.Form("ID")
	oRs(1) = Request.Form("Contraseña°Anterior")
	oRs(2) = Request.Form("Contraseña°Nueva")
	oRs(3) = Request.Form("Contraseña°Nueva°Confirmación")
	oRs.Update

	oRs.save oDOM, adPersistXML

	If oUsuario.ChangePassword(request.form("ID"), oDOM.xml, lErrNum, sErrDesc, sErrSource) Then
		Set oUsuario = Nothing
		Set oRs = Nothing
		Set oDOM = Nothing
		Select Case session("rol")
			Case "Administrador"
				response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/roles_list.asp?msgid=0&msg=Contraseña modificada correctamente")
			Case "Bibliotecario"
				response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp?msgid=0&msg=Contraseña modificada correctamente")
			Case "Docente"
				response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp?msgid=0&msg=Contraseña modificada correctamente")
			Case "Alumno"
				response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp?msgid=0&msg=Contraseña modificada correctamente")
			Case Else
				response.redirect "index.asp?msgid=1&msg=El Rol utilizado todavia no fue implementado.<BR>Por favor contacte a su administrador del sistema."
		End Select
	Else
		Set oUsuario = Nothing
		Set oRs = Nothing
		Set oDOM = Nothing
		response.redirect "change_pwd.asp?msgid=1&ID=" & Request.Form("ID") & "&msg=" & sErrDesc 
	End If
	
	Set oUsuario = Nothing
	Set oRs = Nothing
	Set oDOM = Nothing
Else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		response.redirect "/admin/error.asp?msgid=1&title=" & strTitle & "&msg=No se puede cambiar la contraseña."
	End If

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
					<form action="change_pwd.asp" method="POST" name="myform" onsubmit="return validate_form();">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5" class="c_text"><div align="center" style="margin: 2px 0px 0px 0px; font-weight:bolder; color=#FFA8A8;"><% if len(request.querystring("msg")) > 0 then %>
			<iframe id="iFrameMsg" name='iFrameMsg' FRAMEBORDER="0" SCROLLING='no' WIDTH="85%"  HEIGHT="30" src="mensajes.asp?msgid=<%=request.querystring("msgID")%>&error=<%=request.querystring("msg")%>&msg=">
			</iframe>
		<%end if%></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Contraseña anterior:</div></td>
                          <td width="200"><div align="left">
                            <input type="password" name="Contraseña°Anterior" value=""><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_Contraseña°Anterior">
                          </div></td>
						  <Td><div id="Contraseña°AnteriorError"></div>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Contraseña nueva:</div></td>
                          <td width="200"><div align="left">
                            <input type="password" name="Contraseña°Nueva" value=""><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_Contraseña°Nueva">
                          </div></td>
						  <Td><div id="Contraseña°NuevaError"></div>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Contraseña nueva (Confirmación):</div></td>
                          <td width="200"><div align="left">
                            <input type="password" name="Contraseña°Nueva°Confirmación" value=""><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_Contraseña°Nueva°Confirmación">
                          </div></td>
						  <Td><div id="Contraseña°Nueva°ConfirmaciónError"></div>
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
				<INPUT TYPE="hidden" name="ID" value="<%=iID%>">
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
End If
%>