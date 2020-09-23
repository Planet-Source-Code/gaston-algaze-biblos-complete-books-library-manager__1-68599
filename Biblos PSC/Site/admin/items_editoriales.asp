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
Dim oEditorial
Dim oEditorialAux
Dim iID, iEditorialID
Dim strXML, strSearch
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim oDom

iID = Server.HTMLEncode(Request.Querystring("ID"))
'If len(iID) = 0 Or isNumeric(iID) = False Then
'	iID = "sql_injection_attempt"
'End If

Set oEditorial = Server.CreateObject("Biblos_BR.cEditorial")
Set oEditorialAux = Server.CreateObject("Biblos_BR.cEditorial")
%>
<select name="cboEditorial" id="cboEditorial" style="width:150px;" onchange='location.href="items_editoriales.asp?ID="+ this.value;'>
<option value="-1" selected>-Seleccione-</option>
<%

Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
Set oRs = Server.CreateObject("ADODB.Recordset")

strSearch = ""
If len(iID) > 0 then

	strSearch = "id = " & iID
	if oEditorialAux.Search(session("userID"), strXML, cstr(strSearch), , , lErrNum, sErrDesc, sErrSource) then
		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				iEditorialID = oRs("ID")
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

if oEditorial.Search(session("userID"), strXML, "fecha_baja is NULL", , , lErrNum, sErrDesc, sErrSource) then

	Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
	Set oRs = Server.CreateObject("ADODB.Recordset")

	If oDOM.loadXML(strXML) Then
		Set oRs = RecordsetFromXMLDocument(oDOM)
		if not oRs.EOF then
			While Not oRs.EOF
				response.write "<option value=""" & oRs("ID") & """ " & IIf(oRs("ID") = cint(iEditorialID),"selected","") & " >" & oRs("Nombre") &  "</option>"
				oRs.MoveNext
			Wend
		end if
	End if
End if

%>
  </select>&nbsp;&nbsp;<a class="l_text" onClick="myRef = window.open('editoriales_insert.asp?ID=<%=iID%>','mywin',
'left=200,top=200,width=300,height=400,scrollbars=0,toolbar=0,resizable=0,menubar=0');
myRef.focus()" href="javascript:void(0);"><IMG SRC="/images/insert.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Agregar nuevo"></a>&nbsp;|&nbsp;<a class="l_text" onClick="if (document.form2.cboEditorial.value != '-1') {  myRef = window.open('editoriales_update.asp?ID=<%=iEditorialID%>','mywin',
'left=200,top=200,width=300,height=400,scrollbars=0,toolbar=0,resizable=0,menubar=0'); myRef.focus(); } else { alert('Por favor seleccione una opcion'); return false;}" href="javascript:void(0);"><IMG SRC="/images/update.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Actualizar"></a>&nbsp;|&nbsp;<a class="l_text" href="editoriales_delete.asp?ID=<%=iEditorialID%>" onclick="if (document.form2.cboEditorial.value != '-1') { return confirm('¿Está seguro que desea eliminar el registro?'); } else { alert('Por favor seleccione una opcion'); return false;}"><IMG SRC="/images/delete.gif" WIDTH="11" HEIGHT="12" BORDER="0" ALT="Borrar"></a>&nbsp;<font class="m_text">*</font>
</form>
</BODY>
</HTML>