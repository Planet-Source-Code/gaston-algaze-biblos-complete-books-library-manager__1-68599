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
<!--#include virtual="/includes/public_list.asp"-->
<%
strTitle = "Items"

Dim oItem
Dim oCategoria
Dim oEditorial
Dim oItemTipo
Dim strXML, strSearch
Dim oDOM
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim isAlumno

Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")
Set oEditorial = Server.CreateObject("Biblos_BR.cEditorial")
Set oItemTipo = Server.CreateObject("Biblos_BR.cItemTipo")

if session("rol") = "Alumno" then 
	isAlumno = true
Else
	isAlumno = false
End if

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
<!-- ACA COMIENZA EL FORM -->
		<form action="items_search_andor.asp" method="POST" name="myform" onsubmit="document.myform.subcategoria.value=parent.iframe1.document.form2.elements[0].value; document.myform.categoria.value=parent.iframe1.document.form2.elements[0].value; document.myform.subcategoria.value=parent.iframe2.document.form2.elements[0].value;">
		<% if len(request.querystring("msg")) > 0 then %>
			<iframe id="iFrameMsg" name='iFrameMsg' FRAMEBORDER="0" SCROLLING='no' WIDTH="85%"  HEIGHT="30" src="mensajes.asp?msgid=<%=request.querystring("msgID")%>&error=<%=request.querystring("msg")%>&msg=">
			</iframe>
		<%end if%>
				<TABLE width="85%" align="center">
				<TR class="l_text">
                    	<TD><div align="right">Autor:</div></TD>
                        <TD><input type="text" name="autor" value="<%=iif(len(request.querystring("msg")) > 0, session("search_autor"), request.form("autor"))%>"><input type="hidden" name="@ _NoBlank_autor"></TD>
<TD>&nbsp;</TD>
<td><div id="autorError"></div></td>
                        <TD><div align="right">T&iacute;tulo:</div></TD>
                    	<TD><%
						Dim strAux

						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectTitulo"), request.form("cboSelectTitulo"))

						%><select name="cboSelectTitulo">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><input type="text" name="titulo" value="<%=iif(len(request.querystring("msg")) > 0, session("search_titulo"), request.form("titulo"))%>"><input type="hidden" name="@ _NoBlank_titulo"></TD>
<TD></TD>
<td><div id="tituloError"></div></td>
                    </TR>
                    <TR class="l_text">
                    	<TD><div align="right">ISBN:</div></TD>
                        <TD><%
						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectISBN"), request.form("cboSelectISBN"))

						%><select name="cboSelectISBN">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><input type="text" name="isbn" value="<%=iif(len(request.querystring("msg")) > 0, session("search_isbn"), request.form("isbn"))%>"><input type="hidden" name="@ _NoBlank_isbn"></TD>
<TD></TD>
<td><div id="isbnError"></div></td>
                        <TD><div align="right">Editorial:</div></TD>
                      <TD><%
						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectEditorial"), request.form("cboSelectEditorial"))

						%><select name="cboSelectEditorial">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><select name="cboEditorial" id="cboEditorial">
							<option value="-1" selected>-Todas-</option>
						  <%
						  strAux =iif(len(request.querystring("msg")) > 0, session("search_cboEditorial"), request.form("cboeditorial"))
						  if oEditorial.Search(session("userID"), strXML, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=""" & oRs("ID") & """" & IIf(oRs("ID") = Cint(strAux), "selected", "") &  ">" & oRs("Nombre") &  "</option>"
									oRs.MoveNext
								Wend
							End if
						  End if
						  %>
							  </select></TD>
<TD></TD>
<td><div id="categoriaError"></div></td>
                    </TR>
                      <TR class="l_text">
					<%if Not isAlumno then%>
                    	<TD nowrap align="right">Tipo de Item:</TD>
                        <TD nowrap><%
						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectItemTipo"), request.form("cboSelectItemTipo"))

						%><select name="cboSelectItemTipo">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><select name="cboItemTipo" id="cboItemTipo" onchange='document.getElementById("iframe1").src="items_categorias_public.asp?itemtipoID="+ this.value; document.getElementById("iframe2").src="items_subcategorias_public.asp"; document.getElementById("iframe3").src="links_list_public.asp";'>
							<option value="-1" selected>-Todas-</option>
						  <%
						  strAux =iif(len(request.querystring("msg")) > 0, session("search_cboItemTipo"), request.form("cboItemTipo"))
						  if oItemTipo.Search(session("userID"), strXML, , , , lErrNum, sErrDesc, sErrSource) then
							Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
							Set oRs = Server.CreateObject("ADODB.Recordset")
							
							If oDOM.loadXML(strXML) Then
								Set oRs = RecordsetFromXMLDocument(oDOM)
								While Not oRs.EOF
									response.write "<option value=""" & oRs("ID") & """" & IIf(oRs("ID") = Cint(strAux), "selected", "") &  ">" & oRs("Descripcion") &  "</option>"
									oRs.MoveNext
								Wend
							End if
						  End if
						  %></select></TD>
						<%Else%>
						<TD><div align="right">&nbsp;</div></TD>
                        <TD>&nbsp;<input type="hidden" name="cboItemTipo" value="1">
						<input type="hidden" name="cboSelectItemTipo" value="1"></TD>
						<%End If%>
						<TD nowrap>
						<TD></TD>
						 <TD><div align="right">Categoria:</div></TD>
                      <TD valign="middle"><%
						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectCategoria"), request.form("cboSelectCategoria"))

						%><select name="cboSelectCategoria">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><iframe id="iframe1" name='iframe1' FRAMEBORDER=0 SCROLLING='no' WIDTH=160  HEIGHT=21 src="items_categorias_public.asp?ID=<%=iif(len(request.querystring("msg")) > 0, session("search_categoria"), request.form("categoria"))%>&ItemTipoID=<%=iif(len(request.querystring("msg")) > 0, session("search_cboitemtipo"), IIf(IsAlumno,"1",request.form("cboitemtipo")) )%>"></iframe><input type="hidden" name="categoria" value="<%=iif(len(request.querystring("msg")) > 0, session("search_categoria"), request.form("categoria"))%>"></TD>
<TD></TD>
<td><div id="categoriaError"></div></td>
                    </TR>
					<TR class="l_text">
                      <TD colspan="4">&nbsp;</TD>
						<TD><div align="right">Subcategoria:</div></TD>
                      <TD valign="middle"><%
						strAux = iif(len(request.querystring("msg")) > 0, session("search_cboSelectSubCategoria"), request.form("cboSelectSubCategoria"))

						%><select name="cboSelectSubCategoria">
<option value="1" <%=IIf(strAux = 1,"selected","")%>>Y</option>
<option value="2" <%=IIf(strAux = 2,"selected","")%>>O</option>
                        </select><iframe valign="abs-middle" id="iframe2" name='iframe2' FRAMEBORDER=0 SCROLLING='no' WIDTH=160  HEIGHT=21 src="items_subcategorias_public.asp?ID=<%=iif(len(request.querystring("msg")) > 0, session("search_subcategoria"), request.form("subcategoria"))%>&CatID=<%=iif(len(request.querystring("msg")) > 0, session("search_categoria"), request.form("categoria"))%>"></iframe><input type="hidden" name="subcategoria" value="<%=iif(len(request.querystring("msg")) > 0, session("search_subcategoria"), request.form("subcategoria"))%>"></TD>
<TD></TD>
<td><div id="subcategoriaError"></div></td>
                      <TD><div align="right"></div></TD>
                      <TD></TD><td></td>
                    </TR>
                    <TR class="l_text">
                      <TD><div align="left">
                        <input type="image" SRC="/images/buscar.gif" name="Submit" value="Submit">
						<input type="hidden" name="Submit" value="true">
                      </div></TD>
                      <TD colspan="2">&nbsp;</TD>
                    </TR>
                    </TABLE>
		</form>
<!-- ACA TERMINA EL FORM -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
<TR>
	<TD align="center"><iframe id="iframe3" name='iframe3' FRAMEBORDER=0 SCROLLING='yes' WIDTH="85%"  HEIGHT="60" src="links_list_public.asp?subcatID=<%=iif(len(request.querystring("msg")) > 0, session("search_subcategoria"), request.form("subcategoria"))%>"></iframe></TD>
</TR>
</table>
<%
	
	Set oItem = Server.CreateObject("Biblos_BR.cItem")
	
	strSearch = "( "

	if len(request.form("submit")) > 0 then
		if len(request.form("autor")) > 0 then
			session("search_autor") = request.form("autor")
			strSearch = strSearch & "autor LIKE '%" & request.form("autor") & "%' "
		end if		
		if len(request.form("titulo")) > 0 then
			session("search_titulo") = request.form("titulo")
			session("search_cboSelectTitulo") = request.form("cboSelectTitulo")
			if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectTitulo") = 1, ") AND ( ", "OR ")
			strSearch = strSearch & "titulo LIKE '%" & request.form("titulo") & "%' "
		end if		
		if len(request.form("isbn")) > 0 then
			session("search_isbn") = request.form("isbn")
			session("search_cboSelectisbn") = request.form("cboSelectisbn")
			if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectisbn") = 1, ") AND ( ", "OR ")
			strSearch = strSearch & "isbn LIKE '%" & request.form("isbn") & "%' "
		end if	
		if not isAlumno then
			if request.form("cboItemTipo") <> "-1" then
			    session("search_cboItemTipo") = request.form("cboItemTipo")
				session("search_cboSelectItemTipo") = request.form("cboSelectItemTipo")
				if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectItemTipo") = 1, ") AND ( ", "OR ")
				strSearch = strSearch & "itemTipoID = " & request.form("cboItemTipo") & " "
			end if
		else
			session("search_cboItemTipo") = 1
			session("search_cboSelectTitulo") = 1
		end if
		if request.form("cboEditorial") <> "-1" then
			session("search_cboEditorial") = request.form("cboEditorial")
			session("search_cboSelectEditorial") = request.form("cboSelectEditorial")
			if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectEditorial") = 1, ") AND ( ", "OR ")
			strSearch = strSearch & "editorialID = " & request.form("cboEditorial") & " "
		end if		
		if len(request.form("Categoria")) > 0 AND request.form("Categoria") <> "-1" then
			session("search_Categoria") = request.form("Categoria")
			session("search_cboSelectCategoria") = request.form("cboSelectCategoria")
			if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectCategoria") = 1, ") AND ( ", "OR ")
			strSearch = strSearch & "categoriaID = " & request.form("Categoria") & " "
		end if	
		if len(request.form("SubCategoria")) > 0 AND request.form("SubCategoria") <> "-1" then
			session("search_SubCategoria") = request.form("SubCategoria")
			session("search_cboSelectsubCategoria") = request.form("cboSelectsubCategoria")
			if len(strSearch) > 2 then strSearch = strSearch & iif(request.form("cboSelectsubCategoria") = 1, ") AND ( ", "OR ")
			strSearch = strSearch & "subcategoriaID = " & request.form("SubCategoria") & " "
		end if	
	end if

	if len(strSearch) > 2 then 
		strSearch = strSearch & ") AND fecha_baja IS NULL "
	Else
		strSearch =" fecha_baja IS NULL "
	end if

	if isAlumno then strSearch = strSearch & " AND itemtipoID = 1 "

	if len(request.querystring("msg")) = 0 then session("last_search") = strsearch
	
	'response.write "strsearch: " & strsearch & "<br>"
	'response.write "session(""last_search""): " & session("last_search") & "<br>"

	If oItem.SearchForBorrow(session("userID"), strXML, iif(len(session("last_search")) > 0, session("last_search"), cstr(strSearch)) , "titulo, copias", "DESC", lErrNum, sErrDesc, sErrSource) Then
		CreateTable strXML, "items", 10, CInt(Request.QueryString("page")), IIf(session("rol")="Bibliotecario","true","false")
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