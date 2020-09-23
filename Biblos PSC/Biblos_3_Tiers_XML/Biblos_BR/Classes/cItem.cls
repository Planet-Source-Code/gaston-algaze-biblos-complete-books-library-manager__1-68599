VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cItem
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
Private mTitulo As String
Private mAutor As String
Private mISBN As String
Private mAnno As String
Private mEditorialID As Integer
Private mEditorial As String
Private mUbicacionID As Integer
Private mUbicacion As String
Private mCategoriaID As String
Private mCategoria As String
Private mSubCategoriaID As String
Private mSubcategoria As String
Private mFecha_Alta As String
Private mFecha_Ult_Act As String
Private mFecha_Baja As String
Private mPrestado As Integer
Private mItemTipoID As Integer
Private mItemTipo As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get Prestado() As Integer
    Prestado = mPrestado
End Property

Public Property Let Prestado(ByVal vNewValue As Integer)
    mPrestado = vNewValue
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

Public Property Get ISBN() As String
    ISBN = mISBN
End Property

Public Property Let ISBN(ByVal vNewValue As String)
    mISBN = vNewValue
End Property

Public Property Get Anno() As String
    Anno = mAnno
End Property

Public Property Let Anno(ByVal vNewValue As String)
    mAnno = vNewValue
End Property

Public Property Get EditorialID() As Integer
    EditorialID = mEditorialID
End Property

Public Property Let EditorialID(ByVal vNewValue As Integer)
    mEditorialID = vNewValue
End Property

Public Property Get Editorial() As String
    Editorial = mEditorial
End Property

Public Property Let Editorial(ByVal vNewValue As String)
    mEditorial = vNewValue
End Property

Public Property Get UbicacionID() As Integer
    UbicacionID = mUbicacionID
End Property

Public Property Let UbicacionID(ByVal vNewValue As Integer)
    mUbicacionID = vNewValue
End Property

Public Property Get Ubicacion() As String
    Ubicacion = mUbicacion
End Property

Public Property Let Ubicacion(ByVal vNewValue As String)
    mUbicacion = vNewValue
End Property

Public Property Get CategoriaID() As Integer
    CategoriaID = mCategoriaID
End Property

Public Property Let CategoriaID(ByVal vNewValue As Integer)
    mCategoriaID = vNewValue
End Property

Public Property Get Categoria() As String
    Categoria = mCategoria
End Property

Public Property Let Categoria(ByVal vNewValue As String)
    mCategoria = vNewValue
End Property

Public Property Get SubCategoriaID() As Integer
    SubCategoriaID = mSubCategoriaID
End Property

Public Property Let SubCategoriaID(ByVal vNewValue As Integer)
    mSubCategoriaID = vNewValue
End Property

Public Property Get Subcategoria() As String
    Subcategoria = mSubcategoria
End Property

Public Property Let Subcategoria(ByVal vNewValue As String)
    mSubcategoria = vNewValue
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

Public Property Get ItemTipoID() As Integer
    ItemTipoID = mItemTipoID
End Property

Public Property Let ItemTipoID(ByVal vNewValue As Integer)
    mItemTipoID = vNewValue
End Property

Public Property Get ItemTipo() As String
    ItemTipo = mItemTipo
End Property

Public Property Let ItemTipo(ByVal vNewValue As String)
    mItemTipo = vNewValue
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
        If Len(Titulo) > 30 Or Len(Titulo) = 0 Or _
        Len(Autor) > 255 Or Len(Autor) = 0 Or _
        Len(ISBN) > 20 Or Len(ISBN) = 0 Or _
        Len(Anno) > 4 Or Len(Anno) = 0 Or _
        Not IsNumeric(ID) Then
            ValidateMe = False
            Err.Raise -1, "cItem.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        If Len(Titulo) > 30 Or Len(Titulo) = 0 Or _
        Len(Autor) > 255 Or Len(Autor) = 0 Or _
        Len(ISBN) > 20 Or Len(ISBN) = 0 Or _
        Len(Anno) > 4 Or Len(Anno) = 0 Then
            ValidateMe = False
            Err.Raise -1, "cItem.ValidateMe", "Longitud o tipo de parámetros incorrecto."
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
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cItem
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cItem.Add", "Longitud de parámetros incorrecta."
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "titulo = '" & Titulo & "' AND autor = '" & Autor & "' AND fecha_baja IS NULL AND itemtipoID = " & ItemTipoID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cItem
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cItem.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor actualice el numero de copias."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cItem.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor actualice el numero de copias."
            End If
            
            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "ISBN", adBSTR
            oRs.fields.Append "anno", adBSTR
            oRs.fields.Append "editorialID", adBSTR
            oRs.fields.Append "editorial", adBSTR
            oRs.fields.Append "ubicacionID", adBSTR
            oRs.fields.Append "ubicacion", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "categoriaID", adBSTR
            oRs.fields.Append "categoria", adBSTR
            oRs.fields.Append "subcategoriaID", adBSTR
            oRs.fields.Append "subcategoria", adBSTR
            oRs.fields.Append "prestado", adBSTR
            oRs.fields.Append "ItemTipoID", adBSTR
            oRs.fields.Append "ItemTipo", adBSTR

            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Titulo
            oRs(1) = Autor
            oRs(2) = ISBN
            oRs(3) = Anno
            oRs(4) = EditorialID
            oRs(5) = Editorial
            oRs(6) = UbicacionID
            oRs(7) = Ubicacion
            oRs(8) = Format(Now(), "yyyymmdd")
            oRs(9) = Format(Now(), "yyyymmdd")
            oRs(10) = CategoriaID
            oRs(11) = Categoria
            oRs(12) = SubCategoriaID
            oRs(13) = Subcategoria
            oRs(14) = Prestado
            oRs(15) = ItemTipoID
            oRs(16) = ItemTipo
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveItem(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
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
    
        ShowError "cItem.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función AddCopy()
'   Description    - Inserta un nuevo elemento
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function AddCopy(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddCopyErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cItem
Dim oSecAgent As New cSecurityAgent

    AddCopy = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cItem
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            
            If lErrNum = "-2" Then
                Err.Raise -1, "cItem.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "ISBN", adBSTR
            oRs.fields.Append "anno", adBSTR
            oRs.fields.Append "editorialID", adBSTR
            oRs.fields.Append "editorial", adBSTR
            oRs.fields.Append "ubicacionID", adBSTR
            oRs.fields.Append "ubicacion", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "categoriaID", adBSTR
            oRs.fields.Append "categoria", adBSTR
            oRs.fields.Append "subcategoriaID", adBSTR
            oRs.fields.Append "subcategoria", adBSTR
            oRs.fields.Append "itemtipoID", adBSTR
            oRs.fields.Append "itemtipo", adBSTR
            

            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = oHelper.Titulo
            oRs(1) = oHelper.Autor
            oRs(2) = oHelper.ISBN
            oRs(3) = oHelper.Anno
            oRs(4) = oHelper.EditorialID
            oRs(5) = oHelper.Editorial
            oRs(6) = oHelper.UbicacionID
            oRs(7) = oHelper.Ubicacion
            oRs(8) = Format(Now(), "yyyymmdd")
            oRs(9) = Format(Now(), "yyyymmdd")
            oRs(10) = oHelper.CategoriaID
            oRs(11) = oHelper.Categoria
            oRs(12) = oHelper.SubCategoriaID
            oRs(13) = oHelper.Subcategoria
            oRs(14) = oHelper.ItemTipoID
            oRs(15) = oHelper.ItemTipo
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            AddCopy = oDataAccessTier.SaveItem(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If AddCopy Then
        Exit Function
    Else
    
AddCopyErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cItem.AddCopy", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
            ID = oRs("id")
            Titulo = oRs("titulo")
            Autor = oRs("autor")
            ISBN = oRs("isbn")
            Anno = oRs("año")
            Editorial = oRs("Editorial")
            EditorialID = oRs("EditorialID")
            Ubicacion = oRs("ubicacion")
            UbicacionID = oRs("ubicacionID")
            Categoria = oRs("Categoria")
            CategoriaID = oRs("CategoriaID")
            Subcategoria = oRs("Subcategoria")
            SubCategoriaID = oRs("SubcategoriaID")
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))
            Prestado = IIf(IsNull(oRs("prestado")), "0", oRs("prestado"))
            ItemTipoID = oRs("itemtipoID")
            ItemTipo = oRs("tipo")
            Read = True
    Else
        Err.Raise -1, "cItem.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cItem.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "All", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "titulo ", "Items.titulo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Items.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "ubicacionID ", "Items.id_ubicacion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "subcategoriaID ", "subcategorias.id_subcategoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "categoriaID ", "categorias.id_categoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "editorialID ", "Items.id_editorial ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Items.id_Item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Items.id_Item ", , , vbTextCompare)
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cItem") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cItem.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cItem.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchItem(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cItem.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SearchForBorrow()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchForBorrow(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    SearchForBorrow = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "All", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "titulo ", "Items.titulo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Items.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "ubicacionID ", "Items.id_ubicacion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "subcategoriaID ", "subCategorias.id_subCategoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "categoriaID ", "categorias.id_categoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "editorialID ", "Items.id_editorial ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Items.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Items.id_Item ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cItem") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cItem.SearchForBorrow", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cItem.SearchForBorrow", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    SearchForBorrow = oDataAccessTier.SearchItemForBorrow(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If SearchForBorrow Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cItem.SearchForBorrow", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SearchForReserve()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchForReserve(iUserID As Variant, sXML As Variant, sXMLData As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    SearchForReserve = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "All", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "titulo ", "Items.titulo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Items.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "ubicacionID ", "Items.id_ubicacion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "subcategoriaID ", "subCategorias.id_subCategoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "categoriaID ", "categorias.id_categoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "editorialID ", "Items.id_editorial ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Items.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Items.id_Item ", , , vbTextCompare)
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cItem") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cItem.SearchForReserve", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cItem.SearchForReserve", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    SearchForReserve = oDataAccessTier.SearchItemForReserve(strXML, CStr(sXMLData), sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If SearchForReserve Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cItem.SearchForReserve", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SearchForDelete()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchForDelete(iUserID As Variant, sXML As Variant, sXMLData As Variant, iAll As Integer, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    SearchForDelete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "All", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "titulo ", "Items.titulo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "Items.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "ubicacionID ", "Items.id_ubicacion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "subcategoriaID ", "subCategorias.id_subCategoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "categoriaID ", "categorias.id_categoria ", , , vbTextCompare)
    sSearch = Replace(sSearch, "editorialID ", "Items.id_editorial ", , , vbTextCompare)
    sSearch = Replace(sSearch, "itemtipoID ", "Items.id_item_tipo ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "Items.id_Item ", , , vbTextCompare)
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cItem") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cItem.SearchForDelete", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cItem.SearchForDelete", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    SearchForDelete = oDataAccessTier.SearchItemForDelete(strXML, CStr(sXMLData), iAll, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If SearchForDelete Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cItem.SearchForDelete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
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
Dim oReserva As New cReserva
Dim oPrestamo As New cPrestamo
Dim strXML As String
Dim strXMLAux As String
Dim i As Integer

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cItem.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then

            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "itemTipoID", adBSTR

            oRs.open
            
            oRs.AddNew

            oRs(0) = Titulo
            oRs(1) = Autor
            oRs(2) = ItemTipoID

            oRs.Update
            oRs.save oDOM, adPersistXML

            If SearchForDelete(iUserID, strXMLAux, oDOM.xml, 1, , , , lErrNum, sErrDesc, sErrSource) Then
                If Read(iUserID, strXMLAux, lErrNum, sErrDesc, sErrSource) Then
                    Err.Raise -1, "cItem.Delete", "La copia del item no se puede eliminar pues la misma esta prestada o reservada."
                Else
                    Delete = oDataAccessTier.DeleteItem(strXML, 1, lErrNum, sErrDesc, sErrSource)
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
    
        ShowError "cItem.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función DeleteCopy()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function DeleteCopy(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oSecAgent As New cSecurityAgent
Dim oReserva As New cReserva
Dim oPrestamo As New cPrestamo
Dim strXML As String
Dim i As Integer

    DeleteCopy = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cItem.DeleteCopy", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then

            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "itemTipoID", adBSTR

            oRs.open
            
            oRs.AddNew

            oRs(0) = Titulo
            oRs(1) = Autor
            oRs(2) = ItemTipoID

            oRs.Update
            oRs.save oDOM, adPersistXML

            If SearchForDelete(iUserID, strXML, oDOM.xml, 0, , , , lErrNum, sErrDesc, sErrSource) Then
                If Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
                    DeleteCopy = oDataAccessTier.DeleteItemCopy(strXML, 1, lErrNum, sErrDesc, sErrSource)
                Else
                    Err.Raise -1, "cItem.DeleteCopy", "La copia del item no se puede eliminar pues la misma esta prestada o reservada."
                End If
            End If
        End If
    End If
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If DeleteCopy Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cItem.DeleteCopy", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
Dim oHelper As cItem
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND titulo = '" & Titulo & "' AND autor = '" & Autor & "' AND fecha_baja IS NULL AND itemtipoID = " & ItemTipoID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cItem
        
            If Not ValidateMe Then
                Err.Raise -1, "cItem.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "autor", adBSTR
            oRs.fields.Append "ISBN", adBSTR
            oRs.fields.Append "anno", adBSTR
            oRs.fields.Append "editorial", adBSTR
            oRs.fields.Append "editorialID", adBSTR
            oRs.fields.Append "ubicacion", adBSTR
            oRs.fields.Append "ubicacionID", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "categoriaID", adBSTR
            oRs.fields.Append "categoria", adBSTR
            oRs.fields.Append "subcategoriaID", adBSTR
            oRs.fields.Append "subcategoria", adBSTR
            oRs.fields.Append "itemTipoID", adBSTR
            oRs.fields.Append "itemtipo", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Titulo
            oRs(1) = Autor
            oRs(2) = ISBN
            oRs(3) = Anno
            oRs(4) = Editorial
            oRs(5) = CStr(EditorialID)
            oRs(6) = Ubicacion
            oRs(7) = CStr(UbicacionID)
            oRs(8) = Format(Now(), "yyyymmdd")
            oRs(9) = ID
            oRs(10) = CStr(CategoriaID)
            oRs(11) = Categoria
            oRs(12) = CStr(SubCategoriaID)
            oRs(13) = Subcategoria
            oRs(14) = ItemTipoID
            oRs(15) = ItemTipo
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateItem(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing

    Else
        Err.Raise -1, "cItem.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cItem.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función UpdateStatus()
'   Description    - Actualiza el estado del Item.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function UpdateStatus(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateStatusErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    UpdateStatus = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cItem.Update", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "prestado", adBSTR
    oRs.fields.Append "ItemID", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = Prestado
    oRs(1) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    UpdateStatus = oDataAccessTier.UpdateStatusItem(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    
    If UpdateStatus Then
        Exit Function
    Else
    
UpdateStatusErr:
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cItem.UpdateStatus", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Recover()
'   Description    - Recupera lógicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Recover(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo RecoverErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cItem
Dim oSecAgent As New cSecurityAgent

    Recover = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cItem", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
        
    Set oHelper = New cItem
        
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cItem.Recover", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & CStr(ID), , , lErrNum, sErrDesc, sErrSource) Then
        oHelper.Read strXML, lErrNum, sErrDesc, sErrSource
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND titulo = '" & oHelper.Titulo & "' AND autor = '" & oHelper.Autor & "' AND itemtipoID = " & oHelper.ItemTipoID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cItem
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cItem.Recover", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos o actualice el número de copias."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cItem.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Recover = oDataAccessTier.RecoverItem(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
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
    
        ShowError "cItem.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

