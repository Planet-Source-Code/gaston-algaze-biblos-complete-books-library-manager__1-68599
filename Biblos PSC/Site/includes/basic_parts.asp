<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub Header()
%>									  


			        <table width="673" style="height:120px" border="0" cellspacing="0" cellpadding="0">
					  <tr>
						<td width="186" align="left" valign="top" style="height:120px">
							<div style="margin-left:44px; margin-top:19px;">
							  <div align="left"><a href="/index.asp"><img src="<%=strLogoPath%>" alt="" border="0" align="top"></a></div>
							</div>
						</td>
						<td nowrap width="487" style="height:120px ; margin-top:20px; background-image: url(/images/button-spacer.gif); background-repeat: repeat-x;" valign="top" >
							<div class="h_text_bolder" align="center" style="margin-left:0px; margin-top:35px" >&nbsp;&nbsp;&nbsp;Bienvenidos&nbsp;al&nbsp;Sistema&nbsp;Biblos</div>
						</td>
					  </tr>
					</table>
<%
End Sub
%>

<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub Footer()
%>		
		<table width="100%" style="height:464px" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="100%" style="height:464px" valign="top">
					<table width="100%" style="height:464px" border="0" cellspacing="0" cellpadding="0">
					  <tr>
						<td width="100%" style="height:425px" valign="top">
							<table width="100%" style="height:425px" border="0" cellspacing="0" cellpadding="0">
							  <tr>
								<td width="41%" style="height:100%" valign="top">
									<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td class="m_text" width="100%" style="height:100%" bgcolor="#F3F3F3" valign="top">
											<div style="margin-left:21px; margin-top:31px; margin-right:20px; margin-bottom:20px">
												<div style="margin-left:px; margin-top:px;"><img alt="" border="0" src="/images/t1.gif"></div>
												<div style="margin-left:px; margin-top:18px; margin-right:px; margin-bottom:px">
													<div><img alt="" src="/images/fotito1.png" hspace="0" vspace="0" border="0" align="left" style="margin-right:12px; margin-top:0px"></div>
													<div><img alt=""  border="0"  src="/images/spacer.gif"></div>
													<div align="left" class="_text" style="margin-top:5px"><a href="reglamento.asp" class="h_text"><strong>Reglamento de la Biblioteca</strong></a></div>
													<div align="left" class="_text" style="margin-left:px; margin-top:3px; margin-right:px">Aquí podra encontrar el reglamento vigente de nuestra biblioteca.</div>
												</div>
												<div style="margin-left:px; margin-top:35px; margin-right:px; margin-bottom:px">
													<div><img alt="" src="/images/fotito2.png" hspace="0" vspace="0" border="0" align="left" style="margin-right:12px; margin-top:0px"></div>
													<div><img alt=""  border="0"  src="/images/spacer.gif"></div>
													<div align="left" class="_text" style="margin-top:5px"><a onclick="myRef = window.open('http://www.sistemabiblos.com.ar');" href="javascript:void(0);" class="h_text"><strong>Nuestro Colegio</strong></a></div>
													<div align="left" class="_text" style="margin-left:px; margin-top:3px; margin-right:px">Si queres saber mas sobre nuestra institución, visita nuestro sitio web.</div>
												</div>
												<div style="margin-left:px; margin-top:35px; margin-right:px; margin-bottom:px">
													<div><img alt="" src="/images/fotito3.png" hspace="0" vspace="0" border="0" align="left" style="margin-right:12px; margin-top:0px"></div>
													<div><img alt=""  border="0"  src="/images/spacer.gif"></div>
													<div align="left" class="_text" style="margin-top:5px"><a href="javascript:void(0);" class="h_text"><strong>Contactenos</strong></a></div>
													<div align="left" class="_text" style="margin-left:px; margin-top:3px; margin-right:px">¿Tenes alguna consulta? Contactá a nuestro equipo.</div>
												</div>
											</div>
										</td>
									  </tr>
									  <tr>
										<td>
									<%
										Company()
									%>		
										<td>
									  <tr>
									</table>
								</td>
								<td width="6" style="height:425px" valign="top"><div><img  src="/images/spacer.gif" alt="" width="6" height="1"  border="0"></div></td>
								<td width="59%" style="height:100%" valign="top">
									<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td width="100%" style="height:100%" valign="top">
											<table width="100%" style="height:100%" border="0" cellspacing="0" cellpadding="0">
											  <tr>
												<td width="100%" style="height:148px" valign="top">
													<div style="margin-left:16px; margin-top:31px;"><img alt="" border="0" src="/images/t2.gif"></div>
													<div align="left" class="m1_text" style="margin-left:13px; margin-top:14px; margin-right:31px"><span class="h_text"><u>Items al alcance de tu mano</u></span><span class="h_text"><br>
													</span>Gracias a nuestro nuevo sistema, podrás retirar los items que necesitás de una manera rápida y sencilla. 
Ahora, podés consultar on- line nuestro catálogo, para ver si está el material que buscás Y... quién te dice,  tal vez interesarte por algún item que ni siquiera sabías que existía.  
Porque Biblos pone los items al alcance de tu mano! 
</div>
												</td>
											  </tr>
											  <tr>
												<td width="100%" style="height:1px" bgcolor="#DADADA" valign="top"></td>
											  </tr>
											  <tr>
												<td width="100%" style="height:100%" valign="top">
													<div style="margin-left:16px; margin-top:2px;"><img alt="" border="0" src="/images/t3.gif"></div>
													<%GetRSS strRSSFeed,2%>
												</td>
											  </tr>
											</table>
										</td>
									  </tr>
									  <tr>
										<td width="100%" style="background-image: url(/images/b-dr.gif); background-repeat: repeat-x; background-position:;height:59px" valign="top">
											<div style="margin-left:17px; margin-top:20px;">
												<table align="center" style="height:px" border="0" cellspacing="0" cellpadding="0">
												  <tr >
													<td valign="top" ><div class="h_text"><a href="javascript:void(0);" onclick="myRef = window.open('/help.htm','mywin',
'left=100,top=100,width=600,height=400,scrollbars=yes,toolbar=0,resizable=0,menubar=0');" class="h_text"><strong>Si necesitas ayuda, hace click aqui.</strong></a></div></td>
												  </tr>
												</table>
											</div>
										</td>
									  </tr>
									</table>
								</td>
								<td width="7" style="height:425px" valign="top"><div><img  src="/images/spacer.gif" alt="" width="7" height="1"  border="0"></div></td>
							  </tr>
							</table>
						</td>
					  </tr>
					  <tr>
						<td width="100%" style="height:39px" valign="top"></td>
					  </tr>
					</table>
				</td>
				<td width="48" style="height:464px" valign="top"><div style="margin-left:px; margin-top:px;"><img alt="" border="0" src="/images/r.gif"></div></td>
			  </tr>
			</table>
<%
End Sub
%>

<%'''''''''''''''''''''''''''''''''''''''''''''''
Sub Body()
%>
&nbsp;
<%
End Sub
%>


<%''''''''''''''''''''''''''''''''''''''''''''''''
Sub Company()
%>
									  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                       <tr>
										<td width="100%" style="height:59px" bgcolor="#93A81C" valign="top">
											<div align="left" class="c_text" style="margin-left:29px; margin-top:24px; "><a class="c_text" href="">CHMR &copy; 2006&nbsp; |&nbsp;  Todos los derechos reservados</a></div>
										</td>
									  </tr>
									 </table>

<%
End Sub
%>

				
