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
<script type="text/javascript">

<!--

function validate_form ()
{
		valid = checkform(myform, '#ffcccc', '#ffffff', true, false, 'color: #FF0000; font-weight: bold; font-family: arial; font-size: 8pt;');
		if (valid != false){
			if ( document.myform.archivo.value == "" )
			{
					alert ( "Por favor seleccione un archivo." );
					document.myform.archivo.focus();
					valid = false;
			}
		};
		if (valid != false) please_wait();
		else return false;

}

//-->

</script>

<SCRIPT LANGUAGE=JavaScript>

function please_wait(){
	document.all.pleasewaitScreen.style.pixelTop = (document.body.scrollTop + 100);
	document.all.pleasewaitScreen.style.visibility="visible";
	window.setTimeout('upload()',1000);
}

// files upload function
function upload()
{
	 // create ADO-stream Object
   var oStream = new ActiveXObject("ADODB.Stream");
 
   // create XML document with default header and primary node
   var oDOM = new ActiveXObject("MSXML2.DOMDocument");
   oDOM.loadXML('<?xml version="1.0" ?> <root/>');
   // specify namespaces datatypes
   oDOM.documentElement.setAttribute("xmlns:dt", "urn:schemas-microsoft-com:datatypes");

   // create a new node and set binary content
   var l_node1 = oDOM.createElement("file1");
   l_node1.dataType = "bin.base64";
   // open stream object and read source file
   oStream.Type = 1;  // 1=adTypeBinary 
   oStream.Open(); 
   oStream.LoadFromFile(document.myform.archivo.value);
   if (oStream.Size > 10485760){
	   div_msg.innerHTML = "<iframe id=\"iFrameMsg\" name='iFrameMsg' FRAMEBORDER=\"0\" SCROLLING='no' WIDTH=\"85%\"  HEIGHT=\"30\" src=\"mensajes.asp?msgid=1&error=El archivo debe ser menor a 10 Mb.&msg=0\"></iframe>";
	   document.all.pleasewaitScreen.style.visibility="hidden";
	   return false;
   }
   // store file content into XML node
   l_node1.nodeTypedValue = oStream.Read(-1); // -1=adReadAll
   oStream.Close();
   oDOM.documentElement.appendChild(l_node1);

   // we can create more XML nodes for multiple file upload

   // send XML documento to Web server
   var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
   xmlhttp.open("POST","./fichas_upload_insert.asp?titulo=" + document.myform.titulo.value + "&usuarioID=" + document.myform.usuarioID.value + "&filename=" + document.myform.archivo.value +"", false);
   xmlhttp.send(oDOM);
   //alert(oDOM.xml)
   // show server msg in msg-area
   document.all.pleasewaitScreen.style.visibility="hidden";
   div_msg.innerHTML = xmlhttp.ResponseText;
}
</SCRIPT>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<!--#include virtual="/includes/admin_parts.asp"-->
<!--#include virtual="/includes/adovbs.inc"-->
<%
strTitle = "Alta de Ficha"
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
					<DIV ID="pleasewaitScreen" STYLE="position:absolute;z-index:5;top:30%;left:42%;visibility:hidden">
						<TABLE BGCOLOR="#000000" BORDER="1" BORDERCOLOR="#000000" CELLPADDING="0" CELLSPACING="0" HEIGHT="150" WIDTH="250" ID="Table1">
							<TR>
								<TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#ffFFFF" ALIGN="CENTER" VALIGN="MIDDLE">
								<FONT FACE="Tahoma" SIZE="4" COLOR="#93A81C"><B>Subiendo<br>Por Favor Espere...</B></FONT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<form method="POST" name="myform" action="">
					<table width="80%"  border="0" align="center" cellpadding="1" cellspacing="0" style="margin: 0px 100px 0px 0px;">
					  <tr>
					  <td colspan="5"><div align="left" style="margin: 2px 0px 0px 0px;">
					  <DIV id="div_msg" align="center">&nbsp;</DIV></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Título</div></td>
                          <td width="200"><div align="left">
                            <input type="text" name="titulo" value=""><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_titulo">
                          </div></td>
						  <td width="300px"><div id="tituloError"></div></td>
					  </tr>
					  <tr>
						<td width="200" class="h_text"><div align="right">Archivo</div></td>
                          <td width="200"><div align="left">
                            <input type="file" name="archivo"><font class="m_text">*</font>
						    <input type="hidden" name="@ _NoBlank_archivo">
							<INPUT TYPE="hidden" NAME="usuarioID" value="<%=session("userID")%>">
                          </div></td>
						  <td width="300px"><div id="archivoError"></div></td>

					  </tr>
					  <tr>
						   <td width="200" class="h_text"><div align="right"></div></td>
                          <td width="200"><div align="left"><INPUT TYPE="button" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos" onclick="return validate_form();" value="Aceptar">&nbsp;&nbsp;<A href="javascript:void(0);" onclick="history.back();return false"><INPUT TYPE="button" src="/images/ok.gif" width="76" height="18" border="0" ALT="Enviar Datos" onclick="document.href='/admin/fichas_list.asp';" value="Volver"></A></div></td>
						  <td><IMG SRC="/images/spacer.gif" WIDTH="240" HEIGHT="1" BORDER="0" ALT=""></td>
					  </tr>
					  <tr>
						<td colspan="3" class="m_text"><div align="center"><BR>El archivo debera ser menor a 10 mb.</div></td>
					  </tr>
					  <tr>
					  <td colspan="5"><div align="center" style="margin: 2px 0px 0px 0px;"><BR><font class="m_text">Los campos marcados con (*) son obligatorios.</font></div></td>
					  </tr>
					</table>
					</td>
                  </tr>
                </table>
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