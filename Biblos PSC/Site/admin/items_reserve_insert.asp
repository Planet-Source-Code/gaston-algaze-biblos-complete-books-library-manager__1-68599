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
		var oDate       = new Date();
		var intMonth	= oDate.getMonth() + 1;
		var intDay		= oDate.getDate();
		var intYear		= oDate.getYear();

		if(intYear < 2000) { intYear = intYear + 1900; }
		var strDate = intDay + '/' + intMonth + '/' + intYear;

		valid = true;
		
		if(invFecha(1,strDate) > invFecha(1,document.myform.reserva.value) && document.myform.reserva.value.length != 0 ) {
			alert("Fecha de reserva inválida.\nLa misma no puede ser menor a la fecha de hoy.")
		return false;
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
strTitle = "Alta de Reserva"

Dim lErrNum, sErrDesc, sErrSource

Dim oReserva, oReservaAux, oItem
Dim strXML
Dim iID
Dim oDom
Dim oRs
Dim arrAux 
Dim dateAux
Dim strMsg

Set oReserva = Server.CreateObject("Biblos_BR.cReserva")
Set oReservaAux = Server.CreateObject("Biblos_BR.cReserva")
Set oItem = Server.CreateObject("Biblos_BR.cItem")

if len(request.form("submit")) > 0 then

	arrAux = Split(request("cboReservaTipo"),"_")

	if formatdate(request.form("reserva")) = formatdate(date) then
		strMsg = "No esta permitido reservar un item para el mismo día."
			response.redirect "items_reserve_insert.asp?msgID=2&msg=" & strMsg & "&id=" & request.form("ItemID")
	End if

	With oReserva
		.Fecha_reserva = request.form("reserva")
		.UsuarioID = request.form("usuarioID")
		.itemtipoID = request.form("itemtipoID")
	End With

	If oItem.Search(session("userID"), strXML, "id = " & request.form("ItemID"), , , lErrNum, sErrDesc, sErrSource) Then
		If oItem.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then

			Set oRs = Server.CreateObject("ADODB.Recordset")
			Set oDOM = Server.CreateObject("MSXML2.DOMDocument")

			oRs.fields.Append "titulo", adBSTR
			oRs.fields.Append "autor", adBSTR
			oRs.fields.Append "fecha_reserva", adBSTR
			oRs.fields.Append "usuarioID", adBSTR
			oRs.fields.Append "itemTipoID", adBSTR

			oRs.Open
			
			oRs.AddNew

			oRs(0) = oItem.Titulo
			oRs(1) = oItem.Autor
			oRs(2) = oReserva.Fecha_reserva
			oRs(3) = vbNull
			ors(4) = oItem.itemtipoID

			if WeekDay(oReserva.Fecha_reserva) = 1 or WeekDay(oReserva.Fecha_reserva) = 7 then
				strMsg = "Por favor seleccione un día laborable."
					response.redirect "items_reserve_insert.asp?msgID=2&msg=" & strMsg & "&id=" & request.form("ItemID")
			End if

			oRs.Update
			oRs.save oDOM, adPersistXML

			If oItem.SearchForReserve(session("userID"), strXML, oDOM.XML, , , , lErrNum, sErrDesc, sErrSource) Then
				Set oDom = Nothing
				Set oRs = Nothing
				If oItem.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
					oReserva.ItemID = oItem.ID
					if oReservaAux.Search(session("userID"), strXML, "Fecha_reserva = " & formatdate(request.form("reserva")) & " AND titulo = '" & oItem.titulo & "' AND itemtipoID = " & oItem.itemtipoID & " AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
						If oReservaAux.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								if not oRs.EOF then strMsg = "El item ha sido reservado, pero existen reservas (" & oRs.RecordCount & ") anteriores a la suya."
							End If
							Set oDOM = Nothing
							Set oRs = Nothing
						else
							strMsg = "Item reservado con éxito."
						End if

						If oReserva.Add(session("userID"), lErrNum, sErrDesc, sErrSource) Then
							Set oReserva = Nothing
							response.redirect "items_search.asp?msgID=0&msg=" & strMsg
						Else
							response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
						End If
					Else
						response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
					End if
				Else
					strMsg = "No existen copias disponibles para el " & request.form("reserva") & "."
					response.redirect "items_reserve_insert.asp?msgID=2&msg=" & strMsg & "&id=" & request.form("ItemID")
				end if
			Else
				response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
			end if
		Else
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End if
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End if
	Set oReserva = Nothing
else

	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
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
					<form action="items_reserve_insert.asp" method="POST" name="myform" onsubmit="return validate_form();">
					<% if len(request.querystring("msg")) > 0 then %>
			<iframe id="iFrameMsg" name='iFrameMsg' FRAMEBORDER="0" SCROLLING='no' WIDTH="85%"  HEIGHT="30" src="mensajes.asp?error=<%=request.querystring("msg")%>&msgID=<%=request.querystring("msgID")%>">
			</iframe>
		<%end if%>
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
						<td width="220" class="h_text"><div align="right">Fecha de Reserva</div></td>
                          <td width="220" nowrap><div align="left">
                            <input type="text" name="reserva" value="" OnFocus="this.blur();">
						<input type="hidden" name="@ _NoBlank_reserva">&nbsp;<a href="javascript:void(0);" onClick="popUpCalendar(this, myform.reserva, 'dd/mm/yyyy');"><img src="/images/show-calendar.gif" width="24" height="22" border="0" align="absmiddle"></a>&nbsp;<font class="m_text">*</font></div></td>
						  <td nowrap><font class="m_text">(Le recordamos que la fecha de reserva no puede ser la de hoy)</font><div id="fecha_reserva"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Usuario</div></td>
                          <td width="200" class="h_text_black"><div align="left"><%=session("username")%></div></td>
						  <td width="300px"></td>
					  </tr>
					  <tr>
						   <td width="200" class="h_text"><div align="right"></div></td>
                          <td width="200"><div align="left"><INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">&nbsp;&nbsp;<A href="javascript:void(0);" onclick="history.back();return false"><IMG SRC="/images/volver.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
						  <td><IMG SRC="/images/spacer.gif" WIDTH="240" HEIGHT="1" BORDER="0" ALT=""></td>
					  </tr>
					  <tr>
					  <td colspan="3"><div align="center" style="margin: 2px 0px 0px 0px;"><BR><font class="m_text">Los campos marcados con (*) son obligatorios.</font></div></td>
					  </tr>
					</table>
					
					</td>
                  </tr>
                </table>
				<INPUT TYPE="hidden" name="ItemID" value="<%=.ID%>">
				<INPUT TYPE="hidden" name="titulo" value="<%=.titulo%>">
				<INPUT TYPE="hidden" name="autor" value="<%=.autor%>">
				<INPUT TYPE="hidden" name="itemtipoID" value="<%=.itemtipoID%>">
				<INPUT TYPE="hidden" name="usuarioID" value="<%=session("userID")%>">
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