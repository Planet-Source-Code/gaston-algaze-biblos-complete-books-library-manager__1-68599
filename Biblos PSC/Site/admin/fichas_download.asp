<%Option Explicit%>
<%
Dim sXML 
dim oFicha
Dim oXMLFile
Dim oDOM
dim oStream
dim oXML

if len(request.querystring("id")) > 0 then

	Set oFicha = Server.CreateObject("Biblos_BR.cFicha")
	set oDOM = Server.CreateObject("MSXML2.DOMDocument")	
	set oStream = Server.CreateObject("ADODB.Stream")
	set oXML = Server.CreateObject("MSXML2.DOMDocument")

	oFicha.GetFile 1, sXML, "id = " & request.querystring("id")
	oFicha.Read 1, sXML

	oXML.LoadXML oFicha.Archivo

	' retrieve XML node with binary content
	set oXMLFile = oXML.selectSingleNode("root/file1")

	' open stream object and store XML node content into it   
	oStream.Type = 1  ' 1=adTypeBinary 
	oStream.open 
	oStream.Write oXMLFile.nodeTypedValue
	' save uploaded file
	oStream.SaveToFile "c:\temp\tempfile",2  ' 2=adSaveCreateOverWrite 
	
	Response.AddHeader "content-disposition","attachment; filename=" & oFicha.Archivo_nombre
	Response.ContentType = "application/octet-stream"
	Response.BinaryWrite oStream.read

	Response.Flush
	Response.End
	Response.Close

	oStream.close

	' destroy COM object   
	set oStream = Nothing 
	set oXML = Nothing
	Set oFicha = Nothing
	set oDOM = Nothing
end if%>