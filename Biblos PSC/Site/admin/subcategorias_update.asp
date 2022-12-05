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
Dim strXML
Dim iID
Dim lErrNum, sErrDesc, sErrSource
Dim oSubCategoria

Set oSubCategoria = Server.CreateObject("Biblos_BR.cSubCategoria")

strTitle = "SubCategorias_update"

If Len(Request.Form("submit")) > 0 then

	With oSubCategoria
		.ID = request.form("id")
		.CategoriaID = request.form("categoriaID")
		.Descripcion = request.form("Descripcion")
	End With

	If oSubCategoria.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oSubCategoria = Nothing
		%>
		  <script language = JavaScript>
		   //Si el boton nueva estuviera en el parent
		   //window.opener.Reload();
		   //como esta en el iframe...
		   window.opener.parent.iframe2.location.href = "items_SubCategorias.asp?ID=<%=request.form("ID")%>&CatID=<%=request.form("categoriaID")%>";
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

	Set oSubCategoria = Nothing
Else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oSubCategoria.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then

		If oSubCategoria.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			
			With oSubCategoria
%>
<form action="SubCategorias_update.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" bordercolor="#E2E2E2" style="margin: 0px 0px 0px 0px;">
  <tr>
    <td colspan="5" class="h_text_bold">Actualizaci�n Subcategor�a</td>
  </tr>
  <tr>
    <td class="h_text_table">Descripci�n:<BR>
		<input name="descripcion" type="text" size="25" maxlength="255" value="<%=.Descripcion%>">
		<font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_descripcion">
		<input type="hidden" name="id" value=<%=Request.Querystring("ID")%>>
		<input type="hidden" name="categoriaID" value="<%=Request.Querystring("catID")%>">
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