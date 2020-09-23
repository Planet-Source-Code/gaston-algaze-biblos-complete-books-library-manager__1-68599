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
Dim iID, iCategoriaID, iSubcategoriaID, iItemTipoID
Dim strXML, strSearch
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim oDom

iCategoriaID = Server.HTMLEncode(Request.Querystring("catID"))
iID = Server.HTMLEncode(Request.Querystring("ID"))
iItemTipoID = Server.HTMLEncode(Request.Querystring("ItemTipoID"))


Set oSubcategoria = Server.CreateObject("Biblos_BR.cSubCategoria")
Set oSubcategoriaAux = Server.CreateObject("Biblos_BR.cSubCategoria")
%>
<select name="cboSubCategoria" id="cboSubCategoria" style="width:150px;" onchange='location.href="items_subcategorias.asp?catID=<%=iCategoriaID%>&ID="+ this.value;'>
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
  </select>&nbsp;&nbsp;<%if iItemTipoID = "-1" then%><IMG SRC="/images/insert_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;|&nbsp;<IMG SRC="/images/update_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;|&nbsp;<IMG SRC="/images/delete_disabled.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Por favor seleccione un tipo de item">&nbsp;<font class="m_text">*</font><%Else%><a class="l_text" onClick="myRef = window.open('subcategorias_insert.asp?ID=<%=iSubcategoriaID%>&CatID=<%=iCategoriaID%>','mywin',
'left=200,top=200,width=600,height=1,scrollbars=0,toolbar=0,resizable=0,menubar=0');
myRef.focus()" href="javascript:void(0);"><IMG SRC="/images/insert.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Agregar nuevo"></a>&nbsp;|&nbsp;<a class="l_text" onClick="if (document.form2.cboSubCategoria.value != '-1') {  myRef = window.open('subcategorias_update.asp?ID=<%=iSubcategoriaID%>&CatID=<%=iCategoriaID%>','mywin',
'left=200,top=200,width=600,height=1,scrollbars=0,toolbar=0,resizable=0,menubar=0'); myRef.focus(); } else { alert('Por favor seleccione una opcion'); return false;}" href="javascript:void(0);"><IMG SRC="/images/update.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Actualizar"></a>&nbsp;|&nbsp;<a class="l_text" href="subcategorias_delete.asp?ID=<%=iSubcategoriaID%>&CatID=<%=iCategoriaID%>" onclick="if (document.form2.cboSubCategoria.value != '-1') { return confirm('¿Está seguro que desea eliminar el registro?'); } else { alert('Por favor seleccione una opcion'); return false;}"><IMG SRC="/images/delete.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Borrar"></a>&nbsp;<font class="m_text">*</font>
<%End If%>
</form>
</BODY>
</HTML>