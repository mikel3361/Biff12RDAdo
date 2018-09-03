VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cBiff12RDData
Private mCol As cBiff12RDData


Private Sub Class_Initialize()
    Set mCol = New cBiff12RDData
End Sub

Private Sub Class_Terminate()
    mCol.Clear
    Set mCol = Nothing
End Sub

Private Sub cBiff12RDData_Add(newObj As Object, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
    ' Don't delete me you stupid compiler!
    ' We don't want to use this method.
End Sub

Private Sub cBiff12RDData_Clear()
    mCol.Clear
End Sub

Private Property Get cBiff12RDData_Count() As Long
    cBiff12RDData_Count = mCol.Count
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Property Get cBiff12RDData_Item(Index As Variant) As Object
    Set cBiff12RDData_Item = mCol.Item(Index)
End Property

Private Function cBiff12RDData_NewEnum() As IUnknown
    Set cBiff12RDData_NewEnum = mCol.NewEnum
End Function

Private Sub cBiff12RDData_Remove(Index As Variant)
    mCol.Remove Index
End Sub

Public Sub Remove(Index As Variant)
    mCol.Remove Index
End Sub

' A rather simplied interface, just to be quick!
Public Sub Add(ByVal sFieldArray As String, ByVal sDataArray As String)
Attribute Add.VB_UserMemId = -4
    Dim oCols As cBiff12RDCols, oCell As cBiff12RDCell, lCols As Long
    Dim tmp1() As String, tmp2() As String, tmp3() As String
    Dim tmp4() As String, tmp5() As String
    If InStr(sFieldArray, vbTab) > 0 Then
      tmp1 = Split(sFieldArray, vbTab)
    Else
      ReDim tmp1(0) As String
      tmp1(0) = sFieldArray
    End If
    If InStr(sDataArray, vbTab) > 0 Then
      tmp2 = Split(sDataArray, vbTab)
    Else
      ReDim tmp2(0) As String
      tmp2(0) = sDataArray
    End If
    Set oCols = New cBiff12RDCols
    For lCols = 0 To UBound(tmp1)    'Fieldcount = Maximum
      tmp3 = Split(tmp1(lCols), "=")
      tmp4 = Split(tmp3(0), ",")     'FieldNumber, Fieldname
      tmp5 = Split(tmp3(1), ",")
      Set oCell = New cBiff12RDCell  'DataType, DataTypeNumber
      With oCell
        .FieldIndex = CLng(tmp4(0))
        .FieldName = tmp4(1)
        .FieldType = tmp5(0)
        .TypeNumber = CLng(tmp5(1))
        If lCols <= UBound(tmp2) Then
          .CellText = tmp2(lCols)
        Else
          .CellText = ""
        End If
      End With
      oCols.Add oCell
    Next
    mCol.Add oCols
End Sub

Public Sub Clear()
    mCol.Clear
End Sub

Public Function DataRowCount(ByVal Col As Long) As Long
  Dim oCols As cBiff12RDCols, cnt As Long, idx As Long
  cnt = 0
  If Col > 0 Then
    If Not mCol Is Nothing Then
      If mCol.Count > 0 Then   'Rows vorhanden
        For idx = 1 To mCol.Count
          Set oCols = mCol.Item(idx)
          If oCols.Count > 0 And Col <= oCols.Count Then
            If oCols.Cols(Col).IsDataBool Then
              If Not oCols.Cols(Col).IsValueEmpty And oCols.Cols(Col).IsValueBool Then cnt = cnt + 1
            ElseIf oCols.Cols(Col).IsDataDate Then
              If Not oCols.Cols(Col).IsValueEmpty And oCols.Cols(Col).IsValueDate Then cnt = cnt + 1
            ElseIf oCols.Cols(Col).IsDataNumeric Then
              If Not oCols.Cols(Col).IsValueEmpty And oCols.Cols(Col).IsValueNumeric Then cnt = cnt + 1
            Else
              If Not oCols.Cols(Col).IsValueEmpty Then cnt = cnt + 1
            End If
          End If
        Next
        Set oCols = Nothing
      End If
    End If
  End If
  DataRowCount = cnt
End Function

Public Property Get Rows(Index As Variant) As cBiff12RDCols
    Set Rows = mCol.Item(Index)
End Property

Public Function NewEnum() As IUnknown
    Set NewEnum = mCol.NewEnum
End Function


