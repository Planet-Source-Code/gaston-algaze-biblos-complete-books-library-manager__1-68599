<%Option Explicit%>
<!--#include virtual="/includes/global.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="/styles/styles.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY style="margin-top:0px; margin-left:0px;">
<form name="form2">
<%
Dim oSubcategoria
Dim oSubcategoriaAux
Dim iID, iCategoriaID, iSubcategoriaID
Dim strXML, strSearch
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim oDom

iCategoriaID = Server.HTMLEncode(Request.Querystring("catID"))
'If len(iCategoriaID) = 0 Or isNumeric(iCategoriaID) = False Then
'	iCategoriaID = "sql_injection_attempt"
'End If

iID = Server.HTMLEncode(Request.Querystring("ID"))
'If len(iID) = 0 Or isNumeric(iID) = False Then
'	iID = "sql_injection_attempt"
'End If

Set oSubcategoria = Server.CreateObject("Biblos_BR.cSubCategoria")
Set oSubcategoriaAux = Server.CreateObject("Biblos_BR.cSubCategoria")
%>
<select name="cboSubCategoria" id="cboSubCategoria" style="width:100px;" onchange='parent.document.getElementById("iframe3").src="links_list_public.asp?subcatID="+ this.value; location.href="items_subcategorias_public.asp?catID=<%=iCategoriaID%>&ID="+ this.value;'>
<option value="-1" selected>-Seleccione-</option>
<%

Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
Set oRs = Server.CreateObject("ADODB.Recordset")

strSearch = ""

If len(iID) > 0 then 'UPDATE

	strSearch = "id = " & iID
	if oSubcategoriaAux.Search(session("userID"), strXML, cstr(strSearch), , , lErrNum, sErrDesc, sErrSource) then
		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				iSubcategoriaID = oRs("ID")
			end if
		End if
	Else
		%>
		  <script language = JavaScript>
		   parent.location.href = "/admin/error.asp?title=Error&msg=<%=sErrDesc%>";
		  </script>
		<%
	End if

	set oDOM = nothing
	set oRs = nothing

	if oSubcategoria.Search(session("userID"), strXML, "categoriaID = " & iCategoriaID, , , lErrNum, sErrDesc, sErrSource) then

		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
		Set oRs = Server.CreateObject("ADODB.Recordset")

		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				While Not oRs.EOF
					response.write "<option value=""" & oRs("ID") & """ " & IIf(oRs("ID") = iSubcategoriaID,"selected","") & " >" & oRs("Descripcion") &  "</option>"
					oRs.MoveNext
				Wend
			end if
		End if

	End if

Else

	if oSubcategoria.Search(session("userID"), strXML, "categoriaID = " & iCategoriaID, , , lErrNum, sErrDesc, sErrSource) then

		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				While Not oRs.EOF
					response.write "<option value=""" & oRs("ID") & """>" & oRs("Descripcion") &  "</option>"
					oRs.MoveNext
				Wend
			end if
		End if

	End if
End if

%>
  </select>
</form>
</BODY>
</HTML>