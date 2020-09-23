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
<script src="/includes/js/tablas_campos.js"></script>

<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/admin_Reservas.asp"-->
<%
strTitle = "Reservas"

Dim oReserva
Dim oCategoria
Dim strXML
Dim strSearch
Dim oDOM
Dim oRs

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
		<% if len(request.querystring("msg")) > 0 then %>
			<iframe id="iFrameMsg" name='iFrameMsg' FRAMEBORDER="0" SCROLLING='no' WIDTH="85%"  HEIGHT="30" src="mensajes.asp?msgid=<%=request.querystring("msgID")%>&error=<%=request.querystring("msg")%>&msg=">
			</iframe>
		<%end if%>
		<form action="items_reserve_list.asp" method="POST" name="myform">
				<TABLE width="85%">
				<tr>
					  <td colspan="5" class="c_text"><div align="center" style="margin: 2px 0px 0px 0px; font-weight:bolder; color=#FFA8A8;">&nbsp;</div></td>
					  </tr>
                    <TR class="l_text">
                    	<TD width="16%"><div align="right">Fecha de Reserva</div></TD>
                    	<TD width="15%"><input type="text" name="reserva" value="<%=request.form("reserva")%>" OnFocus="this.blur();"><input type="hidden" name="@ _Blank_reserva">&nbsp;<a href="javascript:void(0);" onClick="popUpCalendar(this, myform.reserva, 'dd/mm/yyyy');"><img src="/images/show-calendar.gif" width="24" height="22" border="0" align="absmiddle"></a></TD><td><div id="codigoError"></div></td>
                        <TD><div align="right">ISBN:</div></TD>
                        <TD><input type="text" name="isbn" value="<%=request.form("isbn")%>">
                            <input type="hidden" name="@ _NoBlank_isbn"></TD>
                        <td><div id="ISBNError"></div></td>
                    </TR>
                    <TR class="l_text">
                    	<TD><div align="right">T&iacute;tulo:</div></TD>
                    	<TD><input type="text" name="titulo" value="<%=request.form("titulo")%>"><input type="hidden" name="@ _NoBlank_titulo"></TD><td><div id="tituloError"></div></td>
                        <TD><div align="right">Autor:</div></TD>
                        <TD><input type="text" name="autor" value="<%=request.form("autor")%>">
                            <input type="hidden" name="@ _NoBlank_autor"></TD>
                        <td><div id="autorError"></div></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="right">Matricula:</div></TD>
                      <TD><input type="text" name="matricula" value="<%=request.form("matricula")%>"  maxlength="10"><input type="hidden" name="@ _NoBlank_matricula"></TD><td><div id="matriculaError"></div></td>
                      <TD><div align="right">Apellido:</div></TD>
                      <TD><input type="text" name="apellido" value="<%=request.form("apellido")%>"><input type="hidden" name="@ _NoBlank_apellido"></TD><td><div id="apellidoError"></div></td>
                      <td><div id="apellidoError"></div></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="right">
                        <input type="image" SRC="/images/buscar.gif" name="Submit" value="Submit">
						<input type="hidden" name="Submit" value="true">
                      </div></TD>
					  <TD>&nbsp;&nbsp;<A href="javascript:void(0);" onclick="document.myform.reset(); return false"><IMG SRC="/images/cancelar.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></TD>
                      <TD colspan="2">&nbsp;</TD>
                    </TR>
                    </TABLE>
		</form>
<%
	
	Set oReserva = Server.CreateObject("Biblos_BR.cReserva")

	if Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	if len(request.form("reserva")) > 0 then
			strSearch = strSearch & "fecha_reserva = " & FormatDate(request.form("reserva")) & " "
		end if		
		if len(request.form("titulo")) > 0 then
			if len(strSearch) > 0 then strSearch = strSearch & "AND "
			strSearch = strSearch & "titulo LIKE '%" & request.form("titulo") & "%' "
		end if		
		if len(request.form("autor")) > 0 then
			if len(strSearch) > 0 then strSearch = strSearch & "AND "
			strSearch = strSearch & "autor LIKE '%" & request.form("autor") & "%' "
		end if	
		if len(request.form("isbn")) > 0 then
			if len(strSearch) > 0 then strSearch = strSearch & "AND "
			strSearch = strSearch & "isbn LIKE '%" & request.form("isbn") & "%' "
		end if	
		if len(request.form("matricula")) > 0 then
			if len(strSearch) > 0 then strSearch = strSearch & "AND "
			strSearch = strSearch & "matricula LIKE '%" & request.form("matricula") & "%' "
		end if	
		if len(request.form("apellido")) > 0 then
			if len(strSearch) > 0 then strSearch = strSearch & "AND "
			strSearch = strSearch & "apellido LIKE '%" & request.form("apellido") & "%' "
		end if	
	end if

	If session("rol") <> "Bibliotecario" then
		if len(strSearch) > 0 then
			strSearch = strSearch & " AND usuarioID = " & session("userID")
		else
			strSearch = strSearch & "usuarioID = " & session("userID")
		end if
	End If
	
	if len(strSearch) > 0 then
		strSearch = strSearch & " AND fecha_baja IS NULL"
	else
		strSearch = "fecha_baja IS NULL"
	end if

	If oReserva.Search(session("userID"), strXML, CStr(strSearch), "fecha_reserva", "ASC", lErrNum, sErrDesc, sErrSource) Then	
		CreateTable strXML, "items_reserve", 10, CInt(Request.QueryString("page")), IIf(session("rol")="Bibliotecario","true","false")
	Else
		'Response.Write "<div class=""m_error"" align=""right"">Error Nro: " & cStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource  & "</div>"
		response.redirect "/admin/error.asp?msgID=1&title=" & strTitle & "&msg=" & sErrDesc
	End If

Set oReserva = Nothing

%>					
					</td>
                  </tr>
                </table>
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
	<script language="JavaScript" type="text/javascript" src="/includes/js/wz_tooltip.js"></script>
</body>
</html>