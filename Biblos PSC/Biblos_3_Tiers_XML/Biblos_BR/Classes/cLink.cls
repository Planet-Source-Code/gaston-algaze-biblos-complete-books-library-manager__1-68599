VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLink"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cLink
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 5:34 PM
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
Private mUsuarioID As Integer
Private mItemTipoID As Integer
Private mCategoriaID As Integer
Private mSubCategoriaID As Integer
Private mDescripcion As String
Private mLink As String
Private mFecha_Alta As String
Private mFecha_Ult_Act As String
Private mFecha_Baja As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get UsuarioID() As Integer
    UsuarioID = mUsuarioID
End Property

Public Property Let UsuarioID(ByVal vNewValue As Integer)
    mUsuarioID = vNewValue
End Property

Public Property Get CategoriaID() As Integer
    CategoriaID = mCategoriaID
End Property

Public Property Let CategoriaID(ByVal vNewValue As Integer)
    mCategoriaID = vNewValue
End Property

Public Property Get ItemTipoID() As Integer
    ItemTipoID = mItemTipoID
End Property

Public Property Let ItemTipoID(ByVal vNewValue As Integer)
    mItemTipoID = vNewValue
End Property

Public Property Get SubCategoriaID() As Integer
    SubCategoriaID = mSubCategoriaID
End Property

Public Property Let SubCategoriaID(ByVal vNewValue As Integer)
    mSubCategoriaID = vNewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    mDescripcion = vNewValue
End Property

Public Property Get Link() As String
    Link = mLink
End Property

Public Property Let Link(ByVal vNewValue As String)
    mLink = vNewValue
End Property

Public Property Get Fecha_Alta() As String
    Fecha_Alta = mFecha_Alta
End Property

Public Property Let Fecha_Alta(ByVal vNewValue As String)
    mFecha_Alta = vNewValue
End Property

Public Property Get Fecha_Ult_Act() As String
    Fecha_Ult_Act = mFecha_Ult_Act
End Property

Public Property Let Fecha_Ult_Act(ByVal vNewValue As String)
    mFecha_Ult_Act = vNewValue
End Property

Public Property Get Fecha_Baja() As String
    Fecha_Baja = mFecha_Baja
End Property

Public Property Let Fecha_Baja(ByVal vNewValue As String)
    mFecha_Baja = vNewValue
End Property

'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 Then
        If Len(Link) > 255 Or _
           Len(Link) = 0 Or Not IsNumeric(ID) Then
            ValidateMe = False
            Err.Raise -1, "cLink.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        If Len(Link) > 255 Or Len(Link) = 0 Then
            ValidateMe = False
            Err.Raise -1, "cLink.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    End If
End Function

'========================================================================
'
'   Title          - Función Add()
'   Description    - Inserta un nuevo elemento
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cLink
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cLink.Add", "Longitud de parámetros incorrecta."
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "link = '" & Link & "' AND usuarioID = " & UsuarioID & " AND descripcion = '" & Descripcion & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cLink
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cLink.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cLink.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "subcategoriaID", adBSTR
            oRs.fields.Append "usuarioID", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_mod", adBSTR
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "link", adBSTR
            oRs.fields.Append "categoriaID", adBSTR
            oRs.fields.Append "itemtipoID", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = SubCategoriaID
            oRs(1) = UsuarioID
            oRs(2) = Format(Date, "yyyymmdd")
            oRs(3) = Format(Date, "yyyymmdd")
            oRs(4) = Descripcion
            oRs(5) = Link
            oRs(6) = CategoriaID
            oRs(7) = ItemTipoID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveLink(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
            
        End If
    End If
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cLink.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un descripcion XML
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cLink.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Descripcion = oRs("descripcion")
            UsuarioID = oRs("UsuarioID")
            CategoriaID = oRs("categoriaID")
            SubCategoriaID = oRs("subcategoriaID")
            Link = oRs("link")
            ItemTipoID = oRs("ItemTipoID")
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))
            Read = True
        End If
    Else
        Err.Raise -1, "cLink.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cLink.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    Search = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cLink", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "descripcion ", "Links.descripcion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "usuarioID ", "Links.id_usuario ", , , vbTextCompare)
    sSearch = Replace(sSearch, "subcategoriaID ", "Links.id_subcategoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Links.id_Link ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cLink") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cLink.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cLink.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchLink(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cLink.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oItem As New cItem
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cLink.Update", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    
    Delete = oDataAccessTier.DeleteLink(oDOM.xml, 1, lErrNum, sErrDesc, sErrSource)
        
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cLink.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cLink
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND link = '" & Link & "' and descripcion = '" & Descripcion & "' AND usuarioID = '" & UsuarioID & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cLink
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cLink.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cLink.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            If Not ValidateMe Then
                Err.Raise -1, "cLink.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "link", adBSTR
            oRs.fields.Append "usuarioID", adBSTR
            oRs.fields.Append "fecha_ult_mod", adBSTR, , adFldIsNullable
            oRs.fields.Append "fecha_baja", adBSTR
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "categoriaID", adBSTR
            oRs.fields.Append "subcategoriaID", adBSTR
            oRs.fields.Append "itemtipoID", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Link
            oRs(2) = UsuarioID
            oRs(3) = Format(Date, "yyyymmdd")
            oRs(4) = Format(Date, "yyyymmdd")
            oRs(5) = ID
            oRs(6) = CategoriaID
            oRs(7) = SubCategoriaID
            oRs(8) = ItemTipoID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateLink(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cLink.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cLink.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Recover()
'   Description    - Recupera lógicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - April 2 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Recover(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo RecoverErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cLink
Dim oSecAgent As New cSecurityAgent

    Recover = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cLink", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
        
    Set oHelper = New cLink
        
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cLink.Recover", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & CStr(ID), , , lErrNum, sErrDesc, sErrSource) Then
        oHelper.Read strXML, lErrNum, sErrDesc, sErrSource
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND link = '" & oHelper.Link & "' AND usuarioID = '" & oHelper.UsuarioID & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cLink
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cLink.Recover", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cLink.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Recover = oDataAccessTier.RecoverLink(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Recover Then
        Exit Function
    Else
    
RecoverErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cLink.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function
