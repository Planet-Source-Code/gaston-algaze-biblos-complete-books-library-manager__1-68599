<%Option Explicit%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="content-type" content="text/html;charset=ISO-8859-1">
<meta http-equiv="Content-Script-Type" content="text/javascript; charset=iso-8859-1">
<title>:: Sistema Biblos ::</title>
<link href="styles/styles.css" rel="stylesheet" type="text/css">
<script src="includes/js/basic_functions.js"></script>
<script language='javascript' src="/includes/js/validate.js"></script>
<!--#include virtual="/includes/global.asp"-->
<!--#include virtual="/includes/basic_parts.asp"-->
<%
strTitle = ""
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
<iframe id="iframe1" name='iframe1' FRAMEBORDER=0 SCROLLING='yes' WIDTH="100%"  HEIGHT="300" src="reglamento_biblioteca.htm"></iframe>
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