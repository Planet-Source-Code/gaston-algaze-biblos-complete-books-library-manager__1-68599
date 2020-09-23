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
function validate_form()
{
		valid = true;

        if ( document.myform.cboPrestamoTipo.value == "-1" )
        {
                alert ( "Por favor seleccione un prestamo." );
                valid = false;
        }else{
			 if ( document.myform.cboUsuario.value == "-1" )
			{
					alert ( "Por favor seleccione un usuario." );
					valid = false;
			}
		}
		
		if ( valid == true ) {
			return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');
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
strTitle = "Alta de Prestamo"

Dim lErrNum, sErrDesc, sErrSource

Dim oPrestamo, oUsuario, oItem, oPrestamoTipo
Dim strXML
Dim iID, iUsuarioID
Dim oDom
Dim oRs
Dim arrAux 
Dim dateAux

Set oPrestamo = Server.CreateObject("Biblos_BR.cPrestamo")
Set oUsuario = Server.CreateObject("Biblos_BR.cUsuario")
Set oPrestamoTipo = Server.CreateObject("Biblos_BR.cPrestamoTipo")

if len(request.form("submit")) > 0 then

	arrAux = Split(request("cboPrestamoTipo"),"_")

	With oPrestamo
		.Fecha_Desde = Now()
		dateAux = DateAdd("d", arrAux(1), Now())
		if WeekDay(dateAux) = 1 then
			dateAux	= DateAdd("d", 1, dateAux)
		elseif WeekDay(dateAux) = 7 then
			dateAux	= DateAdd("d", 2, dateAux)
		End if
		.Fecha_Hasta =dateAux
		.UsuarioID = cint(request.form("cboUsuario"))
		.BibliotecariaID = session("userID")
		.ItemID = request.form("ItemID")
		.ItemTipoID = request.form("ItemTipoID")
		.Tipo_PrestamoID = arrAux(0)
		'response.write .Fecha_Desde & "<BR>" 
		'response.write .Fecha_Hasta & "<BR>" 
		'response.write .UsuarioID & "<BR>" 
		'response.write .BibliotecariaID & "<BR>" 
		'response.write .ItemID & "<BR>" 
		'response.write .Tipo_PrestamoID & "<BR>" 
	End With

	If oPrestamo.Add(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oPrestamo = Nothing
'''''''''''
	if len(Request.form("reservaID")) > 0 then

		Dim oReserva

		Set oReserva = Server.CreateObject("Biblos_BR.cReserva")

		iID = Request.form("reservaID")
		If len(iID) = 0 Or isNumeric(iID) = False Then
			iID = "sql_injection_attempt"
		End If

		oReserva.ID = iID

		If oReserva.Delete(session("userID"), lErrNum, sErrDesc, sErrSource) Then
			Set oReserva = Nothing

			response.redirect "items_search.asp?msgID=0&msg=Item prestado con éxito."
		Else
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If

		Set oReserva = Nothing

	Else
		response.redirect "items_search.asp?msgID=0&msg=Item prestado con éxito."
	End if


'''''''''''
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
	
	Set oPrestamo = Nothing
else
	Set oItem = Server.CreateObject("Biblos_BR.cItem")

	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	iUsuarioID = Server.HTMLEncode(Request.Querystring("usuarioID"))
	If len(iUsuarioID) = 0 Or isNumeric(iUsuarioID) = False Then
		iUsuarioID = "sql_injection_attempt"
	End If

	If oItem.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then
		If oItem.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			With oItem

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
					<form action="Items_borrow_insert.asp" method="POST" name="myform" onsubmit="return validate_form();">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5"><div align="left" style="margin: 2px 0px 0px 0px;">&nbsp;</div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Título</div></td>
                          <td width="200"><div align="left" class="h_text_black"><%=.Titulo%>
						    <input type="hidden" name="@ _NoBlank_titulo">
                          </div></td>
						  <Td width="300px"><div id="tituloError"></div>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Autor</div></td>
                          <td width="200"><div align="left" class="h_text_black"><%=.autor%>
						<input type="hidden" name="@ _NoBlank_autor">
                          </div></td>
						  <Td><div id="autorError"></div>
					  </tr>
					  <tr>
						<td width="220" class="h_text"><div align="right">Desde</div></td>
                          <td width="240"><div align="left" class="h_text_black"><%=date()%>
                            <input type="hidden" name="desde" value="<%=FormatDate(Now())%>">
						</div></td>
						  <td><div id="fecha_desde"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Tipo de Prestamo</div></td>
                          <td width="200"><div align="left">
						  <select name="cboPrestamoTipo" id="cboPrestamoTipo" style="width:200px;">
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  if oPrestamoTipo.Search(session("userID"), strXMl, , "descripcion", "ASC", lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXMl) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=" & oRs("ID") & "_" & oRs("duracion") & " selected>" & oRs("Descripcion") & ", " & oRs("duracion") & " días</option>"
									oRs.MoveNext
								Wend
							End if
						else
							response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
						  End if
						  Set oPrestamoTipo = nothing
						  Set oDOM = nothing
						  Set oRs = nothing
						  %>
							  </select><font class="m_text">&nbsp;*</font>
                          </div></td>
						  <td width="300px"></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Usuario</div></td>
                          <td width="200"><div align="left">
						  <select name="cboUsuario" id="cboUsuario" style="width:200px;">
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  if oUsuario.Search(session("userID"), strXML, "rol <> 'administrador' AND fecha_baja IS NULL ", "apellido", "ASC", lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=""" & oRs("ID") & """ " & IIf(cstr(oRs("ID")) = cstr(iUsuarioID),"selected","") & " >" & oRs("Apellido") & ", " & oRs("Nombre") & "</option>"
									oRs.MoveNext
								Wend
							End if
						else
							response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
						  End if
						  Set oUsuario = nothing
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
					</table>
					
					</td>
                  </tr>
                </table>
				<INPUT TYPE="hidden" name="ItemID" value="<%=.ID%>">
				<INPUT TYPE="hidden" name="ItemTipoID" value="<%=.ItemTipoID%>">
				<INPUT TYPE="hidden" name="reservaID" value="<%=request.querystring("reservaID")%>">
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
			Set oItem = Nothing
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If
	Else
		Set oItem = Nothing
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If

	Set oItem = Nothing
End if
%>