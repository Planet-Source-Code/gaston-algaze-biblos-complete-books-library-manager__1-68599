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
Dim oCampo
Dim strCampo, strTabla, strCampoAux
Dim strXML
Dim oRs
Dim lErrNum, sErrDesc, sErrSource
Dim oDom

strTabla = Server.HTMLEncode(Request.Querystring("tabla"))
strCampo = Server.HTMLEncode(Request.Querystring("campo"))

Set oCampo = Server.CreateObject("Biblos_BR.cCampo")
%>
<select name="cboCampo" id="cboCampo" style="width:110px;">
<option value="-1" selected>-Seleccione-</option>
<%

Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
Set oRs = Server.CreateObject("ADODB.Recordset")

	if oCampo.SearchByTabla(session("user"), strTabla, strXML, lErrNum, sErrDesc, sErrSource) then

		Set oDOM = Server.CreateObject("MSXML2.DOMDocument")
		Set oRs = Server.CreateObject("ADODB.Recordset")

		If oDOM.loadXML(strXML) Then
			Set oRs = RecordsetFromXMLDocument(oDOM)
			if not oRs.EOF then
				While Not oRs.EOF
					if Instr(1,  oRs(0), "id_", 1) or Instr(1,  oRs(0), "_alta", 1) or Instr(1,  oRs(0), "_ult_act", 1) or Instr(1,  oRs(0), "_baja", 1) then 
						oRs.MoveNext
					else
						response.write "<option value=""" & oRs(0) & """ " & IIf(oRs(0) = strCampo,"selected","") & " >" & oRs(0) &  "</option>"
						oRs.MoveNext
					End if
				Wend
			end if
		End if

	End if

%>
  </select>
</form>
</BODY>
</HTML>