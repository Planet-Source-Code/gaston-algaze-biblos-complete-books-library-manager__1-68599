<%@ LANGUAGE=VBScript%>
<% Option Explicit
   Response.Expires = 0 
   
   ' define variables and COM objects
   dim oDOM
   dim oFicha
   dim l_node1

   Dim lErrNum, sErrDesc, sErrSource
   Dim strTitle
   Dim iPos, iPrevPos
   Dim strFilename

   ' create XMLDOM object and load it from request ASP object
   set oDOM = Server.CreateObject("MSXML2.DOMDocument")
  
   oDOM.load(request)

   Set oFicha = Server.CreateObject("Biblos_BR.cFicha")

   oFicha.Titulo = request.querystring("titulo")
   oFicha.UsuarioID = request.querystring("usuarioID")

   'Loose the path information and keep just the file name. 
	iPos = instr(1,request.querystring("filename"),"\")
	do while iPos > 0
		iPrevPos = iPos
		iPos = instr(iPrevPos + 1,request.querystring("filename"),"\")
	loop
	strFilename = right(request.querystring("filename"),len(request.querystring("filename")) - iPrevPos)

   oFicha.Archivo_Nombre = replace(strFilename,"\","\\")
   oFicha.Archivo = oDOM.xml

	If oFicha.Add(1, lErrNum, sErrDesc, sErrSource) Then
		Set oFicha = Nothing
		'Response.Write "<iframe id=""iFrameMsg"" name='iFrameMsg' FRAMEBORDER=""0"" SCROLLING='no' WIDTH=""85%""  HEIGHT=""30"" src=""mensajes.asp?error=Archivo subido exitosamente!&msgID=0""></iframe>"
		'response.redirect "/admin/fichas_list.asp?msgid=0&msg=Archivo subido exitosamente!"
		Response.Write "<iframe id=""iFrameMsg"" name='iFrameMsg' FRAMEBORDER=""0"" SCROLLING='no' WIDTH=""85%""  HEIGHT=""30"" src=""fichas_done.asp""></iframe>"
	Else
		response.write "<iframe id=""iFrameMsg"" name='iFrameMsg' FRAMEBORDER=""0"" SCROLLING='no' WIDTH=""85%""  HEIGHT=""30"" src=""mensajes.asp?msgid=1&error=" & sErrDesc & "&msg=0""></iframe>"
	End If

   ' destroy COM object    
   set oDOM = Nothing
   set oFicha = Nothing
   
%>

