<%Option Explicit%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="content-type" content="text/html;charset=ISO-8859-1">
<meta http-equiv="Content-Script-Type" content="text/javascript; charset=iso-8859-1">
<title>:: Sistema Biblos ::</title>
<link rel="shortcut icon" href="favicon.ico" >
<link href="styles/styles.css" rel="stylesheet" type="text/css">
<script src="includes/js/basic_functions.js"></script>
<script language='javascript' src="/includes/js/validate.js"></script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<%
strTitle = ""
Dim lErrNum, sErrDesc, sErrSource
Dim strXML
Dim oSecAgent
Dim oRs
Dim oDOM

if Len(Request.Form("submit")) > 0 then
	
	Set oSecAgent = Server.CreateObject("Biblos_BR.cSecurityAgent")
	If oSecAgent.Login(strXML, request.form("username"), request.form("password")) then
		
		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
		Set oRs = Server.CreateObject("ADODB.Recordset")
		
		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
		
			if Not oRs.EOF then
				session("userID") = oRs(0)
				session("username") = oRs(1)
				Select Case oRs(2)
				Case "Administrador"
					session("rolID") = 1
					session("rol") = oRs(2)
					response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/roles_list.asp")
				Case "Bibliotecario"
					session("rolID") = 2
					session("rol") = oRs(2)
					response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp")
				Case "Alumno"
					session("rolID") = 3
					session("rol") = oRs(2)
					response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp")
				Case "Docente"
					session("rolID") = 4
					session("rol") = oRs(2)
					response.redirect iif(len(request("ref"))>0,request("ref"), "/admin/items_search.asp")
				Case Else
					response.redirect "index.asp?msg=El Rol todavia no fue implementado.<BR>Por favor contacte a su administrador del sistema."
				End Select
			Else
				session.abandon
				response.redirect "index.asp?msg=Nombre de usuario o contraseña incorrectos."
			End if
		Else
			response.redirect "index.asp?msg=Error en la validación."
		End if
	Else
		response.redirect "\admin\error.asp?title=" & strTitle & "&msg=" & sErrDesc
	End if
	set oSecAgent = nothing

Else
%>
</head>

<body bgcolor="#FFFFFF" style="background-image: url(/images/b.gif); background-repeat: repeat-y; background-position: right;">
	<table width="100%"  style="height:820px" border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td width="100%" style="height:120px" valign="top">
			<table width="100%" style="background-image: url(/images/t-dr.gif); background-repeat: repeat-x; ;height:120px" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="673" style="background-image: url(/images/t-l.gif); background-repeat: repeat-x; ;height:120px" valign="top">
<!-- ACA COMIENZA EL HEADER -->
<%
		Header()
%>					
<!-- ACA TERMINA EL HEADER -->
				</td>
				<td width="100%" style="background-image: url(/images/t-r.gif); background-repeat: no-repeat; background-position: right;height:120px" valign="top"><div><img  src="/images/spacer.gif" alt="" width="93" height="1"  border="0"></div></td>
			  </tr>
			</table>
		</td>
	  </tr>
	  <tr>
		<td width="100%" style="height:236px" valign="top">
			<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" style="background-image: url(/images/f-dr.gif); background-repeat: repeat-x;height:236px">
			  <tr>
				<td align="center" valign="top" background="/images/f.jpg">
<!-- ACA COMIENZA EL BODY -->
				<form action="index.asp" method="POST" name="myform" onsubmit="return checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');">
				<table width="100%" height="236"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                <td height="251" valign="top"><BR><BR><BR><table border="0" align="center" cellpadding="1" class="c_text">
							<tr>
							  <td height="28" colspan="2"><div align="right">Para acceder al sistema, por favor ingrese sus datos </div></td>
						    </tr>
							<tr>
					  <td colspan="5"><div align="center" style="margin: 2px 0px 0px 0px; font-weight:bolder; color=#FFA8A8;"><% if len(request.querystring("msg")) > 0 then %>
			<iframe id="iFrameMsg" name='iFrameMsg' FRAMEBORDER="0" SCROLLING='no' WIDTH="85%"  HEIGHT="27" src="/admin/mensajes.asp?error=<%=request.querystring("msg")%>&msgID=1&msg=">
			</iframe>
		<%end if%></div></td>
					  </tr>
					  <tr>
						<td width="200" ><div align="right">Username</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="username" value="">
						    <input type="hidden" name="@ _NoBlank_username">
                          </div></td>
						  <td width="300px"><div id="usernameError"></div></td>
					  </tr>
					  <tr>
						<td width="200" ><div align="right">Password</div></td>
                          <td width="200"><div align="left">
                            <input type="password" name="password" value="">
						    <input type="hidden" name="@ _NoBlank_password">
                          </div></td>
						  <td width="300px"><div id="passwordError"></div></td>
					  </tr>
							<tr>
                              <td>&nbsp;</td>
                              <td><input type="hidden" name="submit" value="1"><INPUT TYPE="image" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos">
							  <input type="hidden" name="ref" value="<%=request.querystring("ref")%>"></td>
						  </tr>
					  </table>
					</td>
                  <td width="436" valign="top"><div align="right"><img src="/images/alfa_home_bw.jpg" width="436" height="236" align="top"></div></td>
                  </tr>
                </table>

				</FORM>
<!-- ACA TERMINA EL BODY -->
				</td>
			  <td width="92" align="center" valign="top" style="background-image: url(/images/t-r-line.gif); background-repeat: repeat-y;width:92px">&nbsp;</td>
			  </tr>
			</table>
		</td>
	  </tr>
	  <tr>
		<td width="100%" style="height:464px" valign="top">
<!-- ACA COMIENZA EL FOOTER -->
<%
		Footer()
%>			
<!-- ACA TERMINA EL FOOTER -->
		</td>
	  </tr>
	</table>
</body>
</html>
<%End If%>