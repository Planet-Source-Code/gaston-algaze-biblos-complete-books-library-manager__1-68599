VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrestamo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cPrestamo
'   Version Number - 1.0
'   Last Updated   - January 18th 2007 - 5:34 PM
'   Author         - Algaze, Gastón
'
'
'
'
'
'========================================================================
' This file contains trade secrets of CHMR No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of CHMR.
'========================================================================

'========================================================================
'Variables miembro
'========================================================================
Private mID As Integer
Private mFecha_Desde As String
Private mFecha_Hasta As String
Private mUsuarioID As Integer
Private mUsuario As String
Private mTitulo As String
Private mAutor As String
Private mBibliotecariaID As String
Private mItemID As String
Private mItemTipoID As String
Private mFecha_Devolucion As String
Private mTipo_PrestamoID As Integer

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get Usuario() As String
    Usuario = mUsuario
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    mUsuario = vNewValue
End Property

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Let Titulo(ByVal vNewValue As String)
    mTitulo = vNewValue
End Property

Public Property Get Autor() As String
    Autor = mAutor
End Property

Public Property Let Autor(ByVal vNewValue As String)
    mAutor = vNewValue
End Property

Public Property Get Fecha_Desde() As String
    Fecha_Desde = mFecha_Desde
End Property

Public Property Let Fecha_Desde(ByVal vNewValue As String)
    mFecha_Desde = vNewValue
End Property

Public Property Get Fecha_Hasta() As String
    Fecha_Hasta = mFecha_Hasta
End Property

Public Property Let Fecha_Hasta(ByVal vNewValue As String)
    mFecha_Hasta = vNewValue
End Property

Public Property Get UsuarioID() As Integer
    UsuarioID = mUsuarioID
End Property

Public Property Let UsuarioID(ByVal vNewValue As Integer)
    mUsuarioID = vNewValue
End Property

Public Property Get BibliotecariaID() As Integer
    BibliotecariaID = mBibliotecariaID
End Property

Public Property Let BibliotecariaID(ByVal vNewValue As Integer)
    mBibliotecariaID = vNewValue
End Property

Public Property Get ItemID() As Integer
    ItemID = mItemID
End Property

Public Property Let ItemID(ByVal vNewValue As Integer)
    mItemID = vNewValue
End Property

Public Property Get ItemTipoID() As Integer
    ItemTipoID = mItemTipoID
End Property

Public Property Let ItemTipoID(ByVal vNewValue As Integer)
    mItemTipoID = vNewValue
End Property

Public Property Get Fecha_Devolucion() As String
    Fecha_Devolucion = mFecha_Devolucion
End Property

Public Property Let Fecha_Devolucion(ByVal vNewValue As String)
    mFecha_Devolucion = vNewValue
End Property

Public Property Get Tipo_PrestamoID() As Integer
    Tipo_PrestamoID = mTipo_PrestamoID
End Property

Public Property Let Tipo_PrestamoID(ByVal vNewValue As Integer)
    mTipo_PrestamoID = vNewValue
End Property

'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 Then
        If Not IsNumeric(ID) And Format(Fecha_Desde, "yyyymmdd") < Format(Now, "yyyymmdd") Then
            ValidateMe = False
            Err.Raise -1, "cPrestamo.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        'If Len(Descripcion) > 255 Or Len(Titulo) > 30 Or _
        '   Len(Titulo) = 0 Then
        '    ValidateMe = False
        '    Err.Raise -1, "cPrestamo.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        'Else
            ValidateMe = True
        'End If
    End If
End Function

'========================================================================
'
'   Title          - Función Add()
'   Description    - Inserta un nuevo elemento
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cPrestamo.Add", "Longitud de parámetros incorrecta."
    End If

    oRs.fields.Append "fecha_desde", adBSTR
    oRs.fields.Append "fecha_hasta", adBSTR
    oRs.fields.Append "usuarioID", adBSTR
    oRs.fields.Append "bibliotecariaID", adBSTR
    oRs.fields.Append "ItemID", adBSTR
    oRs.fields.Append "Tipo_PrestamoID", adBSTR
    oRs.fields.Append "Usuario", adBSTR
    oRs.fields.Append "Titulo", adBSTR
    oRs.fields.Append "Autor", adBSTR
    oRs.fields.Append "ItemTipoID", adBSTR
    
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = Format(Fecha_Desde, "yyyymmdd")
    oRs(1) = Format(Fecha_Hasta, "yyyymmdd")
    oRs(2) = UsuarioID
    oRs(3) = BibliotecariaID
    oRs(4) = ItemID
    oRs(5) = Tipo_PrestamoID
    oRs(6) = Usuario
    oRs(7) = Titulo
    oRs(8) = Autor
    oRs(9) = ItemTipoID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    Add = oDataAccessTier.SavePrestamo(oDOM.xml, lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    Set oDOM = Nothing
        
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cPrestamo.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un archivo XML
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Read(iUserID As Variant, sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo ReadErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Read = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cPrestamo.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Fecha_Desde = oRs("Fecha_Desde")
            Fecha_Hasta = oRs("Fecha_Hasta")
            UsuarioID = oRs("UsuarioID")
            BibliotecariaID = oRs("BibliotecariaID")
            ItemID = oRs("ItemID")
            Fecha_Devolucion = IIf(IsNull(oRs("fecha_devolucion")), "", oRs("fecha_devolucion"))
            Tipo_PrestamoID = oRs("tipo_prestamoID")
            Usuario = oRs("nombre")
            Titulo = oRs("titulo")
            Autor = oRs("autor")
            ItemTipoID = oRs("ItemTipoID")
            
            Read = True
        End If
    Else
        Err.Raise -1, "cPrestamo.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cPrestamo.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    Search = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cPrestamo", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "itemID ", "Prestamos.id_item ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Prestamos.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Prestamos.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "usuarioID ", "Prestamos.id_usuario ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Prestamos.id_Prestamo ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            
            If oSecAgent.ValidateOrderBy(sOrderBy, "cPrestamo") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cPrestamo.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cPrestamo.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchPrestamo(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cPrestamo.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SearchForReport()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchForReport(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo SearchForReportErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    SearchForReport = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cPrestamo", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "itemID ", "Prestamos.id_item ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Prestamos.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "usuarioID ", "Prestamos.id_usuario ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Prestamos.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Prestamos.id_Prestamo ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            
            If oSecAgent.ValidateOrderBy(sOrderBy, "cPrestamo") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cPrestamo.SearchForReport", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cPrestamo.SearchForReport", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    SearchForReport = oDataAccessTier.SearchPrestamoForReport(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If SearchForReport Then
        sXML = strXML
        Exit Function
    Else
    
SearchForReportErr:
    
        ShowError "cPrestamo.SearchForReport", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    
            If Not ValidateMe Then
                Err.Raise -1, "cPrestamo.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "fecha_desde", adBSTR
            oRs.fields.Append "fecha_hasta", adBSTR
            oRs.fields.Append "usuarioID", adBSTR
            oRs.fields.Append "bibliotecariaID", adBSTR
            oRs.fields.Append "ItemID", adBSTR
            oRs.fields.Append "fecha_devolucion", adBSTR
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "tipo_prestamoID", adBSTR
            oRs.fields.Append "usuario", adBSTR
            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "itemtipoID", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Format(Fecha_Desde, "yyyymmdd")
            oRs(1) = Format(Fecha_Hasta, "yyyymmdd")
            oRs(2) = UsuarioID
            oRs(3) = BibliotecariaID
            oRs(4) = ItemID
            oRs(5) = Fecha_Devolucion
            oRs(6) = ID
            oRs(7) = Tipo_PrestamoID
            oRs(8) = Usuario
            oRs(9) = Titulo
            oRs(10) = Autor
            oRs(11) = ItemTipoID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdatePrestamo(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oRs = Nothing
            Set oDOM = Nothing
        
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cPrestamo.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oSecAgent As New cSecurityAgent
Dim oItem As New cItem
Dim strXML As String

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cPrestamo", opDelete) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cPrestamo.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    
    If Search(iUserID, strXML, "id = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            If oItem.Search(iUserID, strXML, "id = " & ItemID, , , lErrNum, sErrDesc, sErrSource) Then
                If oItem.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
                    oItem.Prestado = 0
                    If oDataAccessTier.DeletePrestamo(oDOM.xml, 0, lErrNum, sErrDesc, sErrSource) Then
                        Delete = oItem.UpdateStatus(iUserID, lErrNum, sErrDesc, sErrSource)
                    End If
                End If
            End If
        End If
    End If

    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cPrestamo.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function


