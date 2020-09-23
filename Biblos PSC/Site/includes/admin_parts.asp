<%''''''''''''''''''''''''''''''''''''''''''''''''
if isEmpty(Session("userID")) then
	'Response.Redirect "/index.asp?ref="&Request.ServerVariables("URL")&IIf(Len(Request.ServerVariables("QueryString")) > 0, "?" & Server.HTMLEncode(Request.QueryString), "")
	Response.Redirect "/index.asp"
end if
%>
<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub Header_Admin()
%>									  


			        <table width="673" style="height:120px" border="0" cellspacing="0" cellpadding="0">
					  <tr>
						<td width="186" align="left" valign="top" style="height:120px">
							<div style="margin-left:44px; margin-top:19px;">
							  <div align="left"><a href="/index.asp"><img src="<%=strLogoPath%>" alt="" border="0" align="top"></a></div>
							</div>
						</td>
						<td width="487" style="height:120px ; margin-top:20px; background-image: url(/images/button-spacer.gif); background-repeat: repeat-x;" valign="top" >
							<div class="m1_text_bold" style="margin-left:0px; margin-top:30px">
							M&oacute;dulo <%=session("rol")%>
							</div>
							<div style="margin-top: 10px;"align="left" class="h_text_bold"><%=strTitle%></div>
							<TABLE style="margin-top: 7px;">
							<TR>
								<TD><div align="left" class="m1_text">Usuario:&nbsp;<b><%=strCurrentUser%></b></div></TD>
								<TD><div align="right" class="l_text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="\admin\logout.asp" class="l_text">Salir</A></div></TD>
							</TR>
							</TABLE>
						</td>
					  </tr>
					</table>
<%
End Sub
%>

<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub Footer_Admin()
%>		
								<table width="100%" height="100%" valign="bottom"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td width="293" valign="top" bgcolor="#93A81C" style="height:10%"><% Company() %></td>
									  <td valign="middle" ><div align="center" class="h_text"><a href="javascript:void(0);" onclick="myRef = window.open('/help.htm','mywin',
'left=100,top=100,width=600,height=400,scrollbars=yes,resizable=0,menubar=0');" class="h_text"><strong>Si necesitas ayuda, hace click aqui.</strong></a></div></td>
                                      <td width="93" style="background-image: url(/images/t-r-line.gif); background-repeat: repeat-y; background-position: right;" valign="top"><div><img  src="/images/spacer.gif" alt="" width="93" height="10"  border="0"></div></td>
                                    </tr>
									<tr>
                                      <td></td>
									  <td>&nbsp;</td>
                                      <td width="93" style="background-image: url(/images/t-r-line.gif); background-repeat: repeat-y; background-position: right;" valign="top"><div><img  src="/images/spacer.gif" alt="" width="93" height="10"  border="0"></div></td>
                                    </tr>
									
                                  </table>						  
									
<%
End Sub
%>


<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub MenuBar_Admin(RolID, iID)
Select Case RolID
Case 1 'Administrador
%>	
		<table width="100"  border="0" style="margin-left: 25px ;margin-top: 10px;">
          <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/roles_list.asp">Roles y Seguridad</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/change_pwd.asp?ID=<%=iID%>">Cambiar Contraseña</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/logout.asp">Salir</a></div></td>
          </tr>
		  <tr>
            <td><div><img  src="/images/spacer.gif" alt="" width="93" height="130"  border="0"></div></td>
          </tr>
		  <tr>
            <td><div><img  src="/images/spacer.gif" alt="" width="93" height="130"  border="0"></div></td>
          </tr>

        </table>
<%
Case 2 'Bibliotecario
%>	
		<table width="100"  border="0" style="margin-left: 25px ;margin-top: 10px;">
          <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_search.asp">Buscar Items</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_list.asp">Administrar Items</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_borrow_list.asp">Administrar Prestamos</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_reserve_list.asp">Administrar Reservas</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/usuarios_list.asp">Administrar Usuarios</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_batch_list.asp">Emitir Listado Diario</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/change_pwd.asp?ID=<%=iID%>">Cambiar Contraseña</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/logout.asp">Salir</a></div></td>
          </tr>
		  <tr>
            <td><div><img  src="/images/spacer.gif" alt="" width="93" height="130"  border="0"></div></td>
          </tr>

        </table>
<%
Case 3 'Alumno
%>	
		<table width="100"  border="0" style="margin-left: 25px ;margin-top: 10px;">
          <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_search.asp">Buscar Items</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_borrow_list.asp">Mis Prestamos</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_reserve_list.asp">Mis Reservas</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/change_pwd.asp?ID=<%=iID%>">Cambiar Contraseña</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/logout.asp">Salir</a></div></td>
          </tr>
		  <tr>
            <td><div><img  src="/images/spacer.gif" alt="" width="93" height="130"  border="0"></div></td>
          </tr>

        </table>
<%
Case 4 'Docente
%>	
		<table width="100"  border="0" style="margin-left: 25px ;margin-top: 10px;">
          <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_search.asp">Buscar Items</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_borrow_list.asp">Mis Prestamos</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/items_reserve_list.asp">Mis Reservas</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/fichas_list.asp">Administrar Fichas</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/links_list.asp">Administrar Links</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/change_pwd.asp?ID=<%=iID%>">Cambiar Contraseña</a></div></td>
          </tr>
		  <tr>
            <td><div style="text-align: center; background-image: url(/images/btn_admin.gif); height:32px; width:110px;"><a style="text-decoration:none; vertical-align: bottom; font-family: Tahoma; color:#ffffff; font-size:10px;" href="/admin/logout.asp">Salir</a></div></td>
          </tr>
		  <tr>
            <td><div><img  src="/images/spacer.gif" alt="" width="93" height="130"  border="0"></div></td>
          </tr>

        </table>
<%
End Select
End Sub
%>