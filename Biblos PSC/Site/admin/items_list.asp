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
<script type="text/javascript">

<!--

function validate_form ()
{
		document.myform.subcategoria.value=parent.iframe1.document.form2.elements[0].value;
		valid = checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');
		if (valid != false){
			if ( document.myform.cboCategoria.value == "-1" )
			{
					alert ( "Por favor seleccione una Categoria." );
					document.myform.cboCategoria.focus();
					valid=false;
			}else{
				getIframe();
			}
		};
		return valid;

}

//-->

</script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/admin_items.asp"-->
<%
strTitle = "Items"

Dim oItem
Dim oCategoria
Dim strXML
Dim oDOM
Dim oRs

if Len(Request.Form("submit")) > 0 then
	Dim lErrNum, sErrDesc, sErrSource

	'Set oItem = Server.CreateObject("Biblos_BR.cItem")
	response.write request("subcategoria")
	Set oItem = Nothing
Else

	Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")

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
		<form action="items_list.asp" method="POST" name="myform" onsubmit="return validate_form();">
				<!-- <TABLE width="85%">
                    <TR class="l_text">
                    	<TD width="16%"><div align="right">C&oacute;digo:</div></TD>
                    	<TD width="15%"><input type="text" name="codigo"><input type="hidden" name="@ _NoBlank_codigo"></TD><td><div id="codigoError"></div></td>
                        <TD width="14%"><div align="right">Autor:</div></TD>
                        <TD width="55%"><input type="text" name="autor"><input type="hidden" name="@ _NoBlank_autor"></TD><td><div id="autorError"></div></td>
                    </TR>
                    <TR class="l_text">
                    	<TD><div align="right">T&iacute;tulo:</div></TD>
                    	<TD><input type="text" name="titulo"><input type="hidden" name="@ _NoBlank_titulo"></TD><td><div id=tituloError"></div></td>
                        <TD><div align="right">ISBN:</div></TD>
                        <TD><input type="text" name="isbn"><input type="hidden" name="@ _NoBlank_isbn"></TD><td><div id="isbnError"></div></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="right">Editorial:</div></TD>
                      <TD><input type="text" name="editorial"><input type="hidden" name="@ _NoBlank_editorial"></TD><td><div id="editorialError"></div></td>
                      <TD><div align="right">Item:</div></TD>
                      <TD><input type="text" name="Item"><input type="hidden" name="@ _NoBlank_Item"></TD><td><div id="ItemError"></div></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="right">Categoria:</div></TD>
                      <TD><select name="cboCategoria" id="cboCategoria" onchange='document.getElementById("iframe1").src="items_subcategorias_public.asp?CatID="+ this.value' style="width:100px;">
							<option value="-1" selected>-Seleccione-</option>
						  <%
						  if oCategoria.Search(session("userID"), strXML, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=""" & oRs("ID") & """>" & oRs("Descripcion") &  "</option>"
									oRs.MoveNext
								Wend
							End if
						  End if
						  %>
							  </select></TD><td><div id="categoriaError"></div></td>
                      <TD><div align="right">Subcategoria:</div></TD>
                      <TD valign="middle"><iframe id="iframe1" name='iframe1' FRAMEBORDER=0 SCROLLING='no' WIDTH=120  HEIGHT=25 src="items_subcategorias_public.asp"></iframe><input type="hidden" name="subcategoria"></TD><td><div id="subcategoriaError"></div></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="right">
                        <input type="submit" name="Submit" value="Submit">
                      </div></TD>
                      <TD colspan="3">&nbsp;</TD>
                    </TR>
                    </TABLE> -->
		</form>
<%
	
	Set oItem = Server.CreateObject("Biblos_BR.cItem")

	'If oItem.Search(session("userID"), strXML, "fecha_baja > '" & FormatDate(Now()) & "' OR fecha_baja IS NULL" , , , lErrNum, sErrDesc, sErrSource) Then
	If oItem.Search(session("userID"), strXML, "fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then	
		CreateTable strXML, "items", 10, CInt(Request.QueryString("page"))
	Else
		'Response.Write "<div class=""m_error"" align=""right"">Error Nro: " & cStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource  & "</div>"
		response.redirect "/admin/error.asp?msgID=1&title=" & strTitle & "&msg=" & sErrDesc
	End If

Set oItem = Nothing

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
<%End If%>