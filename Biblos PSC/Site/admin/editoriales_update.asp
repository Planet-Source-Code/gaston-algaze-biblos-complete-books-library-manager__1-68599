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
<script src="/includes/js/tablas_campos.js"></script>
<!--#include virtual="/includes/global.asp"-->
</head>
<%
strTitle = "Editoriales_update"

Dim oEditorial
Dim strXML
Dim iID

Set oEditorial = Server.CreateObject("Biblos_BR.cEditorial")

If Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	With oEditorial
		.ID = request.form("ID")
		.nombre = request.form("nombre")
		.tel1 = request.form("tel1")
		.tel2 = request.form("tel2")
		.mail = request.form("mail")
		.web = request.form("web")
		.domicilio_calle = request.form("calle")
		.domicilio_nro = request.form("nro")
		.domicilio_piso = request.form("piso")
		.domicilio_unidad = request.form("unidad")
		.domicilio_cod_postal = request.form("cod_postal")
	End With

	If oEditorial.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oEditorial = Nothing
		%>
		  <script language = JavaScript>
		   //Si el boton nueva estuviera en el parent
		   //window.opener.Reload();
		   //como esta en el iframe...
		   window.opener.parent.iframe3.location.href = "items_Editoriales.asp?ID=<%=request.form("ID")%>";
		   self.close();
		  </script>
		<%
	Else
		%>
		  <script language = JavaScript>
			window.opener.parent.location.href = "/admin/error.asp?title=Error&msg=<%=sErrDesc%>";
		   self.close();
		  </script>
		<%
	End If

	Set oEditorial = Nothing
Else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oEditorial.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then

		If oEditorial.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			
			With oEditorial
%>
<form action="Editoriales_update.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" bordercolor="#E2E2E2" style="margin: 0px 0px 0px 0px;">
  <tr>
    <td colspan="5" class="h_text_bold">Actualización Editorial</td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Nombre</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="nombre" value="<%=.nombre%>"><font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_nombre">
	  </div></td>
	  <td width="300px"><div id="nombreError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">E-Mail</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="mail" value="<%=.mail%>"><font class="m_text">*</font>
	<input type="hidden" name="@email_NoBlank_mail">
	  </div></td>
	  <td><div id="mailError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Web</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="web" value="<%=.web%>"><font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_web">
	  </div></td>
	  <td width="300px"><div id="webError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Calle</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="calle" value="<%=.domicilio_calle%>"><font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_calle">
	  </div></td>
	  <td width="300px"><div id="calleError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Nro.</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="nro" value="<%=.domicilio_nro%>"><font class="m_text">*</font>
	<input type="hidden" name="@number_NoBlank_nro">
	  </div></td>
	  <td><div id="nroError"></div></td>
  </tr>
   <tr>
	<td width="200" class="h_text"><div align="right">Piso</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="piso" value="<%=.domicilio_piso%>">
	  </div></td>
	  <td width="300px"></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Unidad</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="unidad" value="<%=.domicilio_unidad%>">
	  </div></td>
	  <td></td>
  </tr>
   <tr>
	<td width="200" class="h_text"><div align="right">Cod. Postal</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="cod_postal" value="<%=.domicilio_cod_postal%>"><font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_cod_postal">
	  </div></td>
	  <td width="300px"><div id="cod_postalError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Tel. 1</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="tel1" value="<%=.tel1%>"><font class="m_text">*</font>
	<input type="hidden" name="@ _NoBlank_tel1">
	  </div></td>
	  <td><div id="tel1Error"></div></td>
  </tr>
   <tr>
	<td width="200" class="h_text"><div align="right">Tel. 2</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="tel2" value="<%=.tel2%>">
		<input type="hidden" name="@ _Blank_tel2">
	  </div></td>
	  <td width="300px"><div id="tel2Error"></div></td>
  </tr>
  <tr>
	  <td colspan="4" class="h_text">    
	    <div align="right">
		  <input type="hidden" name="ID" value="<%=Request.Querystring("ID")%>">
		  <INPUT TYPE="hidden" name="submit" value="true">
          <INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">
      &nbsp;&nbsp;<A href="javascript:void(0);" onclick="window.close();return false"><IMG SRC="/images/cerrar.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
	  <td>&nbsp;</td>
  </tr>
</table>
</form>
<%
			End With
		Else
			%>
			  <script language = JavaScript>
				window.opener.parent.location.href = "/admin/error.asp?title=Error&msg=<%=sErrDesc%>";
			   self.close();
			  </script>
			<%
		End if
	Else
		%>
		  <script language = JavaScript>
			window.opener.parent.location.href = "/admin/error.asp?title=Error&msg=<%=sErrDesc%>";
		   self.close();
		  </script>
		<%
	End if
End If

%>