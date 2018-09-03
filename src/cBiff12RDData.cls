VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================
'= Name:        cBiff12RDData
'= Programmer:  Andrew Davey
'= Purpose:     Improve built in collection object and allow implementation.
'===================================================================================
Option Explicit

Private mCol As Collection
#Const isDebug = False   ' Set to false on release build.

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set mCol = Nothing
End Sub

Private Sub displayError()
    #If isDebug Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "cBiff12RDData"
    #Else
        Err.Raise Err.Number, "cBiff12RDData", Err.Description
    #End If
End Sub

Public Sub Add(newObj As Object, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
    Dim bAfter As Boolean, bBefore As Boolean
    
    bAfter = Not IsMissing(After)
    bBefore = Not IsMissing(Before)
    
    If IsMissing(Key) Then
        If bAfter Then
            mCol.Add newObj, , After
            Exit Sub ' Just in case bBefore = True also.
        ElseIf bBefore Then
            mCol.Add newObj, , Before
        Else
            mCol.Add newObj
        End If
    Else
        If bAfter Then
            mCol.Add newObj, Key, After
            Exit Sub ' Just in case bBefore = True also.
        ElseIf bBefore Then
            mCol.Add newObj, Key, Before
        Else
            mCol.Add newObj, Key
        End If
    End If
End Sub

Public Sub Remove(Index As Variant)
    On Error Resume Next
    mCol.Remove Index
    If Err.Number Then displayError
End Sub

Public Sub Clear()
    If mCol Is Nothing Then Exit Sub    ' Shouldn't happen, but just in case.
    
    Dim i As Integer
    For i = mCol.Count To 1 Step -1
        mCol.Remove i
    Next i
End Sub

Public Property Get Item(Index As Variant) As Object
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol.Item(Index)
    If Err.Number Then displayError
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' Allows the nifty For..Each structure to work.
    Set NewEnum = mCol.[_NewEnum]   ' Calls the hidden function.
End Function

Public Property Get Count() As Long
    If mCol Is Nothing Then Exit Property    ' Shouldn't happen, but just in case.

    Count = mCol.Count
End Property

