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
<script type="text/javascript">

<!--

function validate_form ()
{
		valid = true;

		document.myform.campo.value=parent.iframe1.document.form2.elements[0].value;

        if ( document.myform.cboTabla.value == "-1" )
        {
                alert ( "Por favor seleccione una tabla." );
				document.myform.cboTabla.focus();
                valid = false;
        }else{
			if ( document.myform.campo.value == "-1" )
			{
					alert ( "Por favor seleccione un campo." );
					valid = false;
			}else{
				if ( document.myform.cboOperaciones.value == "-1" )
				{
						alert ( "Por favor seleccione una operacion." );
						document.myform.cboOperaciones.focus();
						valid = false;
				}
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
</head>
<%
strTitle = "restricciones_insert"

Dim lErrNum, sErrDesc, sErrSource

Dim oTabla
Dim strXML
Dim iID, iUsuarioID, oOperacion
Dim oDom
Dim oRs
Dim arrAux 
Dim dateAux

Set oOperacion = Server.CreateObject("Biblos_BR.cOperacion")

If Len(Request.Form("submit")) > 0 then
	Dim oRestriccion

	Set oRestriccion = Server.CreateObject("Biblos_BR.cRestriccion")

	With oRestriccion
		.RolID = request.form("RolID")
		.Campo = request.form("campo")
		.Tabla = request.form("cboTabla")
		.OperacionID = request.form("cboOperaciones")
		.Valor = request.form("valor")
	End With

	If oRestriccion.Add(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oRestriccion = Nothing
		%>
		  <script language = JavaScript>
		   //Si el boton nueva estuviera en el parent
		   //window.opener.Reload();
		   //como esta en el iframe...
		   window.opener.location.href = "restricciones_iframe.asp?ID=<%=request.form("RolID")%>";
		   self.close();
		  </script>
		<%
	Else
		%>
		  <script language = JavaScript>
		   window.opener.location.href = "/admin/error.asp?title=Error&msg=<%=sErrDesc%>";
		   self.close();
		  </script>
		<%
	End If

	Set oRestriccion = Nothing
Else
%>
<form action="restricciones_insert.asp" method="POST" name="myform" onsubmit="return validate_form();">
<table width="100%"  border="0" align="center" cellpadding="3" cellspacing="2" bordercolor="#E2E2E2" style="margin: 0px 0px 0px 0px;">
  <tr>
    <td colspan="5" class="h_text_bold">Nueva Restricci&oacute;n</td>
  </tr>
  <tr>
    <td class="h_text_table">Tabla:<BR>
      <select name="cboTabla" id="cboTabla" style="width:120px;" onchange='document.getElementById("iframe1").src="restricciones_campos.asp?tabla="+ this.value;'>
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  Set oTabla = Server.CreateObject("Biblos_BR.cTabla")
						  if oTabla.Search(session("userID"), strXMl, lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXMl) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									if Instr(1,  oRs(0), "_", 1) or oRs(0) = "tablas" or oRs(0) = "campos" or oRs(0) = "restricciones" or oRs(0) = "operaciones" or oRs(0) = "roles" then 
										oRs.MoveNext
									else
										response.write "<option value=" & oRs(0) & ">" & oRs(0) & "</option>"
										if not oRs.EOF then oRs.MoveNext
									END IF
								Wend
							End if
						else
							response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
						  End if
						  Set oTabla = nothing
						  Set oDOM = nothing
						  Set oRs = nothing
						  %>
							  </select>&nbsp;<font class="m_text">*</font>
	</td>
    <td class="h_text_table">Campo:<BR><iframe id="iframe1" name='iframe1' FRAMEBORDER=0 SCROLLING='no' WIDTH="110"  HEIGHT="22" src="restricciones_campos.asp?tabla=-1"></iframe>&nbsp;<font class="m_text">*</font>
	</td>
    <td class="h_text_table">Operaci&oacute;n:<BR>
      <select name="cboOperaciones" id="cboOperaciones" style="width:120px;">
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  Set oOperacion = Server.CreateObject("Biblos_BR.cOperacion")
						  if oOperacion.Search(session("userID"), strXMl, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXMl) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									if Instr(1,  oRs(0), "_", 1) then 
										oRs.MoveNext
									else
										response.write "<option value=""" & oRs(0) & """>" & oRs(1) & "</option>"
										if not oRs.EOF then oRs.MoveNext
									END IF
								Wend
							End if
						else
							response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
						  End if
						  Set oOperacion = nothing
						  Set oDOM = nothing
						  Set oRs = nothing
						  %>
							  </select>&nbsp;<font class="m_text">*</font>
	</td>
    <td class="h_text_table">Valor:<BR>
		<input name="valor" type="text" size="20" maxlength="255">
		<font class="m_text">*</font>
		<input type="hidden" name="@ _NoBlank_valor">
		<input type="hidden" name="RolID" value=<%=Request.Querystring("ID")%>>
		<input type="hidden" name="campo">
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