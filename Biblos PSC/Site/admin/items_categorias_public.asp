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
Dim oCategoria
Dim oCategoriaAux
Dim iID, iCategoriaID, iItemTipoID
Dim strXML, strSearch
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim oDom

iID = Server.HTMLEncode(Request.Querystring("ID"))

iItemTipoID = Server.HTMLEncode(Request.Querystring("ItemTipoID"))


Set oCategoria = Server.CreateObject("Biblos_BR.cCategoria")
Set oCategoriaAux = Server.CreateObject("Biblos_BR.cCategoria")
%>
<select name="cboCategoria" id="cboCategoria" style="width:150px;" onchange='parent.document.getElementById("iframe2").src="items_subcategorias_public.asp?catID="+ this.value; parent.document.getElementById("iframe3").src="links_list_public.asp"; location.href="items_categorias_public.asp?ID="+ this.value + "&ItemTipoID=<%=iItemTipoID%>";'>
<option value="-1" selected>-Seleccione-</option>
<%

Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
Set oRs = Server.CreateObject("ADODB.Recordset")

strSearch = ""
If len(iID) > 0 then

	strSearch = "id = " & iID
	if oCategoriaAux.Search(session("userID"), strXML, cstr(strSearch), , , lErrNum, sErrDesc, sErrSource) then
		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				iCategoriaID = oRs("ID")
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
End if

If len(iItemTipoID) > 0 then
	strSearch = "itemtipoID = " & iItemTipoID
else
	strSearch = "itemtipoID = -1"
end if

if oCategoria.Search(session("userID"), strXML, cstr(strSearch), , , lErrNum, sErrDesc, sErrSource) then

	Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
	Set oRs = Server.CreateObject("ADODB.Recordset")

	If oDOM.loadXML(strXML) Then
		Set oRs = RecordsetFromXMLDocument(oDOM)
		if not oRs.EOF then
			While Not oRs.EOF
				response.write "<option value=""" & oRs("ID") & """ " & IIf(oRs("ID") = cint(iCategoriaID),"selected","") & " >" & oRs("Descripcion") &  "</option>"
				oRs.MoveNext
			Wend
		end if
	End if
End if

%>
  </select>
</form>
</BODY>
</HTML>