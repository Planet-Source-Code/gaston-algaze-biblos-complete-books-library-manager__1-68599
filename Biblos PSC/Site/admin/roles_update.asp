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
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/admin_roles.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<script type="text/javascript">

function Reload() {
var f = document.getElementById('iframe1');
f.src = f.src;
}

</script>
<%
strTitle = "Actualización de Rol"

Dim lErrNum, sErrDesc, sErrSource

Dim oRs
Dim oDOM
Dim oRol
Dim oPermisos
Dim strXML
Dim iID
Dim strXMLRestricciones
Dim strXMLPermisos
Dim i, j

Set oRol = Server.CreateObject("Biblos_BR.cRol")

If len(request.form("submit")) > 0 Then

	

	With oRol
		.ID = request.form("id")
		.PrivilegiosGlobales = Bin2Dec(IIf(Len(Request.Form("chkGlobalA")) = 1, "1", "0") & IIf(Len(Request.Form("chkGlobalB")) = 1, "1", "0") & IIf(Len(Request.Form("chkGlobalM")) = 1, "1", "0") & IIf(Len(Request.Form("chkGlobalL")) = 1, "1", "0"))
		.Descripcion = request.form("descripcion")
	End With

	If oRol.Update(session("userID"), lErrNum, sErrDesc, sErrSource) Then

		Set oRs = Server.CreateObject("ADODB.Recordset")
		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")

		oRs.Fields.Append "RolID", adBSTR
		oRs.Fields.Append "tabla", adBSTR
		oRs.Fields.Append "Permiso", adBSTR
		oRs.Fields.Append "Fecha_Alta", adBSTR
		
		oRs.Open
		
		i = 0
		j = 0 
		FOR I = 4 TO Request.Form.COUNT -6
			oRs.addnew
			oRs(0)= Request.Form("ID")

			for j = 1 to 4

				if Instr(1,  Request.Form(I), "chk", 1) AND Instr(1,  Request.Form(I+1), "chk", 1) then 'es una descripcion
					oRs(1)= replace(replace(replace(replace(replace(Request.Form(I),"A",""),"B",""),"M",""),"L",""),"chk","")
					oRs(2)= oRs(2) & "0" 
				else
					if Instr(1,  Request.Form(I), "chk", 1) AND not Instr(1,  Request.Form(I+1), "chk", 1) then
						oRs(1)= replace(replace(replace(replace(replace(Request.Form(I),"A",""),"B",""),"M",""),"L",""),"chk","")
						oRs(2)= oRs(2) & iif(Request.Form(I+1)="1","1","0")
						i = i + 1
					end if
				end if
				if j < 4 then i = i + 1

				oRs.Update
			next
			
			oRs(2)= Bin2Dec(oRs(2)) 
			oRs(3) = YEAR(Date()) & PadDigits(Month(date()),2) & PadDigits(DAY(date()),2) 

			if ors(1) = "Global" then 
				oRs.Delete
			Else 
				oRs.Update
			End IF
		
			'response.write i & ". descripcion: " & oRs(1) & ", valor: " & oRs(2) & "<BR>"
		NEXT
		
		oRs.save oDOM, adPersistXML

		response.write oDOM.xml
 
		If oRol.SetPermisos(oDOM.xml, lErrNum, sErrDesc, sErrSource) Then
			Set oRol = Nothing
			response.redirect "roles_list.asp?msgid=0&msg=Objeto modificado con éxito."
		Else
			Set oRol = Nothing
				
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If

	Else
		Set oRol = Nothing
	
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If
else
	iID = Server.HTMLEncode(Request.Querystring("ID"))
	If len(iID) = 0 Or isNumeric(iID) = False Then
		iID = "sql_injection_attempt"
	End If

	If oRol.Search(session("userID"), strXML, "id = " & iID, , , lErrNum, sErrDesc, sErrSource) Then
		If oRol.Read(session("userID"), strXML, lErrNum, sErrDesc, sErrSource) Then
			
			With oRol

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
					<form action="roles_update.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, true, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
					<table width="90%"  border="0" cellpadding="2" cellspacing="1" style="margin: 0px 0px 0px 0px;">
					  <tr>
					  <td colspan="6"><div align="left" style="margin: 2px 0px 0px 0px;">&nbsp;</div></td>
					  </tr>
					  <tr>
					    <td class="h_text"><div align="right">Descripci&oacute;n:</div></td>
						<td width="200"><div align="left" class="h_text_bold">
                            <%=.Descripcion%>
						    <input type="hidden" name="@ _NoBlank_descripcion">
							<input type="hidden" name="meta_descripcion" value="descripcion">
							<input type="hidden" name="descripcion" value="<%=.Descripcion%>">
                          </div></td>
					    <td width="300px"><div id="descripcionError"></div></td>                      
					  </tr>
					  <tr>
						<td height="1" colspan="4" align="center" valign="middle" class="h_text"><div align="center" class="l_text">--------------------------------------------------------------------------------
                          
                        </div></td>
					  </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Privilegios<br>
					      Globales</div></td>
					    <td colspan="2" class="h_text"><% GetGlobalPrivileges oRol.PrivilegiosGlobales
						  %></td>
				      </tr>
					  <tr>
					    <td height="1" colspan="4" align="center" valign="middle" class="h_text"><div align="center" class="l_text">-------------------------------------------------------------------------------- </div></td>
				      </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Privilegios</div></td>
					    <td colspan="2"><table width="100%"  border="1" cellpadding="2" cellspacing="2" bordercolor="#E2E2E2">
                          <tr>
                            <td width="68%" class="h_text_table"><div align="center">Tabla</div></td>
                            <td width="8%" class="h_text_table"><div align="center">A</div></td>
                            <td width="8%" class="h_text_table"><div align="center">B</div></td>
                            <td width="8%" class="h_text_table"><div align="center">M</div></td>
                            <td width="8%" class="h_text_table"><div align="center">L</div></td>
                          </tr>
                          <% 
							If oRol.GetPermisos(strXMLPermisos, lErrNum, sErrDesc, sErrSource) Then
								'response.write strXMLPermisos
								GetTablesPrivileges strXMLPermisos
							Else
								response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
							End if
                          %>
				        </table>
				        </td>
					    <td></td>
				      </tr>
					  <tr>
					    <td height="1" colspan="4" align="center" valign="middle" class="h_text"><div align="center" class="l_text">-------------------------------------------------------------------------------- </div></td>
				      </tr>
					  <tr>
					    <td valign="top" class="h_text"><div align="right">Restricciones</div></td>
					    <td colspan="2">
						<iframe id="iframe1" frameborder="0" width="100%" height="200px" src="restricciones_iframe.asp?ID=<%=iID%>"></iframe>
						</td><td></td>
				      </tr>
					  <tr>
					    <td class="h_text">&nbsp;</td>
					    <td colspan="2">&nbsp;</td>
					    <td>&nbsp;</td>
				      </tr>
					  <tr>
						   <td width="92" class="h_text"><div align="right"></div></td>
                          <td colspan="2"><div align="left"><INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">&nbsp;&nbsp;<A href="javascript:void(0);" onclick="history.back();return false"><IMG SRC="/images/volver.gif" WIDTH="76" HEIGHT="18" BORDER="0" ALT="Volver"></A></div></td>
						  <td>&nbsp;</td>
					  </tr>
					  <tr>
					  <td colspan="6"><div align="center" style="margin: 2px 0px 0px 0px;"><BR><font class="m_text">Los campos marcados con (*) son obligatorios.</font></div></td>
					  </tr>
					</table>
					
                  
                </table>
				<INPUT TYPE="hidden" name="meta_ID" value="ID">
				<INPUT TYPE="hidden" name="ID" value="<%=iID%>">
				<INPUT TYPE="hidden" name="meta_submit" value="submit">
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
			Set oRol = Nothing
			response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
		End If
	Else
		Set oRol = Nothing
		response.redirect "/admin/error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End If

	Set oRol = Nothing
End if
%>