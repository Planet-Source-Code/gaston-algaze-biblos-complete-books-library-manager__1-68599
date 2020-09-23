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
		
		return valid;

}

//-->

</script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
strTitle = "Actualización de Fichas"
Dim iID
Dim lErrNum, sErrDesc, sErrSource

Dim oFicha
Dim strXML
	
	Set oFicha = Server.CreateObject("Biblos_BR.cFicha")
	
if len(request.form("submit")) > 0 then

	With oFicha
		.ID = request.form("id")
		.titulo = request.form("titulo")
		.archivo_nombre = request.form("archivo_nombre")
		.usuarioID = request.form("usuarioID")
	End With

	If oFicha.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oFicha = Nothing
		response.redirect "fichas_list.asp?msgID=0&msg=Objeto modificado con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
	
	Set oFicha = Nothing
else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oFicha.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then
		If oFicha.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			With oFicha
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
					<form method="POST" name="myform" action="fichas_update.asp" onclick="return validate_form();">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5"><div align="left" style="margin: 2px 0px 0px 0px;">
					  <DIV id="div_msg" align="center">&nbsp;</DIV></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Título</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="titulo" value="<%=.Titulo%>"><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_titulo">
                          </div></td>
						  <td width="300px"><div id="tituloError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Archivo</div></td>
                          <td width="200"><div align="left" class="m_text"><%=.archivo_nombre%></div></td>
						  <td width="300px"></td>
					  </tr>
					  <tr>
						   <td width="200" class="h_text"><div align="right"></div></td>
						   <input type="hidden" name="ID" value="<%=iID%>">
						   <input type="hidden" name="usuarioID" value="<%=.usuarioID%>">
						   <input type="hidden" name="archivo_nombre" value="<%=.archivo_nombre%>">
                          <td width="200"><div align="left"><INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">&nbsp;&nbsp;<A href="/admin/fichas_list.asp"><IMG SRC="/images/volver.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
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
			End With
		Else
			Set oFicha = Nothing
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If
	Else
		Set oFicha = Nothing
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If

	Set oFicha = Nothing
End if

%>