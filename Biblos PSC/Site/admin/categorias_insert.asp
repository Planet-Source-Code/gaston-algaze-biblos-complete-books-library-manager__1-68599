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
strTitle = "Categorias_insert"

If Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	Dim oCategoria

	Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")

	With oCategoria
		.Descripcion = request.form("Descripcion")
		.itemtipoID = request.form("itemtipoID")
	End With

	If oCategoria.Add(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oCategoria = Nothing
		%>
		  <script language = JavaScript>
		   //Si el boton nueva estuviera en el parent
		   //window.opener.Reload();
		   //como esta en el iframe...
		   window.opener.parent.iframe1.location.href =  "items_categorias.asp?ID=<%=request.form("ID")%>&ItemTipoID=<%=request.form("itemtipoID")%>";
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

	Set oCategoria = Nothing
Else
%>
<form action="Categorias_insert.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" bordercolor="#E2E2E2" style="margin: 0px 0px 0px 0px;">
  <tr>
    <td colspan="5" class="h_text_bold">Nueva Categor�a</td>
  </tr>
  <tr>
    <td class="h_text_table">Descripci�n:<BR>
		<input name="descripcion" type="text" size="25" maxlength="255">
		<font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_descripcion">
		<input type="hidden" name="ID" value="<%=Request.Querystring("ID")%>">
		<input type="hidden" name="itemtipoid" value="<%=Request.Querystring("itemtipoid")%>">
		<INPUT TYPE="hidden" name="submit" value="true">
	</td> 
  </tr>
  <tr>
	  <td colspan="4" class="h_text">    
	    <div align="right">
          <INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">
      &nbsp;&nbsp;<A href="javascript:void(0);" onclick="window.close();return false"><IMG SRC="/images/cerrar.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
	  <td>&nbsp;</td>
  </tr>
</table>
</form>
<%End If%>