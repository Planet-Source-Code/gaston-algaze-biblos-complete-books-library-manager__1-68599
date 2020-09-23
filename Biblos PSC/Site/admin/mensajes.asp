<html>
<head>
<link href="/styles/styles.css" rel="stylesheet" type="text/css">
</head>
<body style="margin-top:0px;margin-left:0px;margin-right:0px;margin-bottom:0px;">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="10">
<tr>
	<td bgcolor="#FFFFCC" height="27" style="padding-left:10px;border-bottom:1px solid #10659E;border-top:1px solid #10659E;border-left:1px solid #10659E;" width="10"><%
	select case cstr(request.querystring("msgID"))
	case "0" 'ok
		response.write "<img src=""/images/noerror.gif"" border=""0"">"
	case "1" 'error
		response.write "<img src=""/images/error.gif"" border=""0"">"
	case "2" 'exclamacion
		response.write "<img src=""/images/exclamation.gif"" border=""0"">"
	case else 'error grave
		response.write "<img src=""/images/error.gif"" border=""0"">"
	end select
	%></td>
	<td bgcolor="#FFFFCC" style="padding-right:10px;border-bottom:1px solid #10659E;border-top:1px solid #10659E;border-right:1px solid #10659E;" align="right" <%
	select case cstr(request.querystring("msgID"))
	case "0" 'ok
		response.write "class=""m_noerror"""
	case "1" 'error
		response.write "class=""m_error"""
	case "2" 'exclamacion
		response.write "class=""m_text_bolder"""
	case else 'error grave
		response.write "class=""m_error"""
	end select
	%>><%=request.querystring("error")%></td>
</tr>
<%if len(request.querystring("msg")) > 0 then%>
<tr>
	<td colspan="2" height="100%" style="padding:18 36 10 36" <%
	select case cstr(request.querystring("msgID"))
	case "0" 'ok
		response.write "class=""m_noerror"""
	case "1" 'error
		response.write "class=""m_error"""
	case "2" 'exclamacion
		response.write "class=""m_text_bolder"""
	case else 'error grave
		response.write "class=""m_error"""
	end select
	%>><%=request.querystring("msg")%></td>
</tr>
<%End If%>
</table>
</body>
</html>
