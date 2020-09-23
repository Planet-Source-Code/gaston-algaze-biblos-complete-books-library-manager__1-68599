<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<link href="/styles/styles.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include virtual="/includes/global.asp"-->
<BODY style="margin-top: 0px; margin-left: 0px;">
<%
Dim oLink
Dim strXML
Dim strSearch
Dim oRs
Dim oDOM
Dim iSubCat
Dim i
Dim lErrNum, sErrDesc, sErrSource
	
if len(request.querystring("subcatID")) > 0 AND cstr(request.querystring("subcatID")) <> "-1" then
	iSubCat = request.querystring("subcatID")
	strSearch = "subcategoriaID = " & iSubCat & " AND fecha_baja IS NULL"
else
	iSubCat = ""
	strSearch = "1 = -1"
end if

	Set oLink = Server.CreateObject("Biblos_BR.cLink")

	'If oLink.Search(session("userID"), strXML, "fecha_baja > '" & FormatDate(Now()) & "' OR fecha_baja IS NULL" , , , lErrNum, sErrDesc, sErrSource) Then
	If oLink.Search(session("userID"), strXML, Cstr(strSearch), , , lErrNum, sErrDesc, sErrSource) Then	
		%>
		<table width="100%"  border="0" align="center" cellpadding="2" cellspacing="1" style="margin: 0px 0px 0px 0px;">
  <tr>
	<td class="h_text_table" colspan="2"><div align="center" style="v-align=middle">links relacionados con la subcategoria</div></td>
	</tr>
	<%
	Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
	Set oRs = Server.CreateObject("ADODB.Recordset")
		
		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			i = 1
			if not oRs.EOF then
				while not oRS.EOF
	%>
	<tr <%=IIf((i mod 2) = 0, "bgcolor=""#EBEBEB""", "")%>>
	<td class="m1_text"><div align="right"><a onclick="myRef = window.open('http://<%=oRs("link")%>');" href="javascript:void(0);" class="m1_text"><%=oRs("descripcion")%></a></div></td>
	<td class="m1_text"><div align="left">&nbsp;-&nbsp;&nbsp;&nbsp;<a onclick="myRef = window.open('http://<%=oRs("link")%>');" href="javascript:void(0);" class="m1_text"><%=oRs("link")%></a></div></td>
	</tr>
		<%
				i = i + 1
				ors.movenext
			Wend
		Else
			%>
	<tr>
	<td class="m1_text" colspan="2"><div align="center">No se encontraron links relacionados</div></td>
	</tr>
		<%
		End If
	Else
		'Response.Write "<div class=""m_error"" align=""right"">Error Nro: " & cStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource  & "</div>"
		response.redirect "/admin/error.asp?msgID=1&title=" & strTitle & "&msg=" & sErrDesc
	End If
Else
	'Response.Write "<div class=""m_error"" align=""right"">Error Nro: " & cStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource  & "</div>"
	response.redirect "/admin/error.asp?msgID=1&title=" & strTitle & "&msg=" & sErrDesc
End If

Set oLink = Nothing

%>				
</BODY>
</HTML>
