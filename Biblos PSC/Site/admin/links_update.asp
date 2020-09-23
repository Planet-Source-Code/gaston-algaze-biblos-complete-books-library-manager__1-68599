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
		
		document.myform.categoria.value=parent.iframe1.document.form2.elements[0].value;
		document.myform.subcategoria.value=parent.iframe2.document.form2.elements[0].value;
		
        if ( document.myform.cboItemTipo.value == "-1" )
        {
                alert ( "Por favor seleccione un tipo de item." );
                valid = false;
        }else{
			if ( document.myform.categoria.value == "-1" )
			{
					alert ( "Por favor seleccione una categoria." );
					valid = false;
			}else{
				if ( document.myform.subcategoria.value == "-1" )
				{
						alert ( "Por favor seleccione una subcategoria." );
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
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
strTitle = "Actualización de Links"

Dim lErrNum, sErrDesc, sErrSource

Dim oLink, oCategoria
Dim strXML
Dim iID
Dim oDom

Set oLink = Server.CreateObject("Biblos_BR.cLink")
Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")

if len(request.form("submit")) > 0 then

	With oLink
		.ID = request.form("id")
		.link = request.form("link")
		.subcategoriaID = request.form("subcategoria")
		.descripcion = request.form("descripcion")
		.usuarioID = session("userID")
	End With

	If oLink.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then
		Set oLink = Nothing
		response.redirect "Links_list.asp?msgID=0&msg=Objeto actualizado con éxito."
	Else
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
	
	Set oLink = Nothing
else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oLink.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then
		If oLink.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) then
			With oLink

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
					<form action="Links_update.asp" method="POST" name="myform" onsubmit="return validate_form();">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5"><div align="left" style="margin: 2px 0px 0px 0px;">&nbsp;</div></td>
					  </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Tipo de Item</div></td>
					    <td colspan="2"><select name="cboItemTipo" id="cboItemTipo" onchange='document.getElementById("iframe1").src="items_categorias_public.asp?itemtipoID="+ this.value'>
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  Dim oItemTipo
						  Dim oRs
						  Set oRs = Server.CreateObject("ADODB.Recordset")
						  Set oItemTipo = Server.CreateObject("Biblos_BR.cItemTipo")
						  if oItemTipo.Search(session("userID"), strXML, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=""" & oRs("ID") & """" & IIf(oRs("ID") = CInt(.ItemTipoID), "selected", "") &  ">" & oRs("Descripcion") &  "</option>"
									oRs.MoveNext
								Wend
							End if
						  End if
						  %>
							  </select><font class="m_text">*</font>
						</td><td></td>
				      </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Categoria</div></td>
					    <td colspan="2"><iframe id="iframe1" name='iframe1' FRAMEBORDER=0 SCROLLING='no' WIDTH=160  HEIGHT=25 src="items_categorias_public.asp?ID=<%=.categoriaID%>&ItemTipoID=<%=.ItemTipoID%>"></iframe><input type="hidden" name="categoria"><font class="m_text">*</font>
						</td><td></td>
				      </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Subcategoria</div></td>
					    <td colspan="2">
						<iframe id="iframe2" name='iframe2' FRAMEBORDER=0 SCROLLING='no' WIDTH=160  HEIGHT=25 src="items_subcategorias_public.asp?ID=<%=.subcategoriaID%>&catID=<%=.categoriaID%>"></iframe><input type="hidden" name="subcategoria"><font class="m_text">*</font>
						</td><td></td>
				      </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Link</div></td>
                          <td nowrap><div align="left">
                            <input type="text" name="link" value="<%=.link%>"><font class="m_text">* ( Sin "http://" )</font>
						    <input type="hidden" name="@ _NoBlank_link">
                          </div></td>
						  <Td width="300px"><div id="linkError"></div>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Descripción</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="descripcion" value="<%=.descripcion%>"><font class="m_text">*</font>
						<input type="hidden" name="@ _NoBlank_descripcion">
                          </div></td>
						  <Td><div id="descripcionError"></div>
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
				<INPUT TYPE="hidden" name="ID" value="<%=.ID%>">
				<INPUT TYPE="hidden" name="submit" value="true">
				<iframe id="iframe3" name='iframe3' FRAMEBORDER=0 SCROLLING='no' WIDTH=0  HEIGHT=0 src=""></iframe>
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
			Set oLink = Nothing
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If
	Else
		Set oLink = Nothing
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If

	Set oLink = Nothing
End if
%>
