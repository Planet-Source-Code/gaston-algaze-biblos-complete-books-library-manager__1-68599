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
<select name="cboCategoria" id="cboCategoria" style="width:150px;" onchange='parent.document.getElementById("iframe2").src="items_subcategorias.asp?catID="+ this.value + "&ItemTipoID=<%=iItemTipoID%>"; location.href="items_categorias.asp?ID="+ this.value + "&ItemTipoID=<%=iItemTipoID%>";'>
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
	strSearch = ""
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
  </select>&nbsp;&nbsp;<%if iItemTipoID = "-1" then%><IMG SRC="/images/insert_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;|&nbsp;<IMG SRC="/images/update_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;|&nbsp;<IMG SRC="/images/delete_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;<font class="m_text">*</font><%Else%><a class="l_text" onClick="myRef = window.open('categorias_insert.asp?ID=<%=iID%>&ItemTipoID=<%=iItemTipoID%>','mywin',
'left=200,top=200,width=600,height=1,scrollbars=0,toolbar=0,resizable=0,menubar=0');
myRef.focus()" href="javascript:void(0);"><IMG SRC="/images/insert.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Agregar nuevo"></a>&nbsp;|&nbsp;<a class="l_text" onClick="if (document.form2.cboCategoria.value != '-1') {  myRef = window.open('categorias_update.asp?ID=<%=iCategoriaID%>&ItemTipoID=<%=iItemTipoID%>','mywin',
'left=200,top=200,width=600,height=1,scrollbars=0,toolbar=0,resizable=0,menubar=0'); myRef.focus(); } else { alert('Por favor seleccione una opcion'); return false;}" href="javascript:void(0);"><IMG SRC="/images/update.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Actualizar"></a>&nbsp;|&nbsp;<a class="l_text" href="categorias_delete.asp?ID=<%=iCategoriaID%>&ItemTipoID=<%=iItemTipoID%>" onclick="if (document.form2.cboCategoria.value != '-1') { return confirm('¿Está seguro que desea eliminar el registro?'); } else { alert('Por favor seleccione una opcion'); return false;}"><IMG SRC="/images/delete.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Borrar"></a>&nbsp;<font class="m_text">*</font>
<%End If%>
</form>
</BODY>
</HTML>