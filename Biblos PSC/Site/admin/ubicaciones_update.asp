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
strTitle = "Ubicaciones_update"

Dim oUbicacion
Dim strXML
Dim iID

Set oUbicacion = Server.CreateObject("Biblos_BR.cUbicacion")

If Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	With oUbicacion
		.ID = request.form("ID")
		.descripcion = request.form("descripcion")
		.titulo = request.form("titulo")
	End With

	If oUbicacion.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oUbicacion = Nothing
		%>
		  <script language = JavaScript>
		   //Si el boton nueva estuviera en el parent
		   //window.opener.Reload();
		   //como esta en el iframe...
		   window.opener.parent.iframe4.location.href = "items_Ubicaciones.asp?ID=<%=request.form("ID")%>";
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

	Set oUbicacion = Nothing
Else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oUbicacion.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then

		If oUbicacion.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			
			With oUbicacion
%>
<form action="Ubicaciones_update.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" bordercolor="#E2E2E2" style="margin: 0px 0px 0px 0px;">
  <tr>
    <td colspan="5" class="h_text_bold">Actualización Ubicacion</td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Titulo</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="titulo" value="<%=.titulo%>"><font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_titulo">
	  </div></td>
	  <td width="300px"><div id="tituloError"></div></td>
  </tr>
  <tr>
	<td width="200" class="h_text"><div align="right">Descripcion</div></td>
	  <td width="200"><div align="left">
		<input type="text" name="descripcion" value="<%=.descripcion%>"><font class="m_text">*</font>
	<input type="hidden" name="@ _NoBlank_descripcion">
	  </div></td>
	  <td><div id="descripcionError"></div></td>
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