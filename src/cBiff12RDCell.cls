VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sCellText    As String
Private m_sFieldName   As String
Private m_sFieldType   As String
Private m_lTypeNumber  As Long
Private m_lFieldIndex  As Long

Public Property Get IsDataBool() As Boolean
  IsDataBool = CBool(m_sFieldType = "BOL")
End Property

Public Property Get IsDataNumeric() As Boolean
  IsDataNumeric = CBool(m_sFieldType = "NUM")
End Property

Public Property Get IsDataDate() As Boolean
  IsDataDate = CBool(m_sFieldType = "DAT")
End Property

Public Property Get IsDataText() As Boolean
  IsDataText = CBool(m_sFieldType = "TXT")
End Property

Public Property Get IsValueBool() As Boolean
  IsValueBool = CBool(InStr("ja,yes,true,wahr,nein,no,false,falsch,", LCase(m_sCellText) & ",") > 0)
End Property

Public Property Get IsValueNumeric() As Boolean
  IsValueNumeric = IsNumeric(m_sCellText)
End Property

Public Property Get IsValueDate() As Boolean
  IsValueDate = IsDate(m_sCellText)
End Property

Public Property Get IsValueEmpty() As Boolean
  IsValueEmpty = CBool(Trim(m_sCellText) = "")
End Property

Public Property Get CellText() As String
    CellText = m_sCellText
End Property
Public Property Let CellText(NewVal As String)
    m_sCellText = NewVal
End Property

Public Property Get FieldName() As String
    FieldName = m_sFieldName
End Property
Public Property Let FieldName(NewVal As String)
    m_sFieldName = NewVal
End Property

Public Property Get FieldType() As String
    FieldType = m_sFieldType
End Property
Public Property Let FieldType(NewVal As String)
    m_sFieldType = NewVal
End Property

Public Property Get FieldIndex() As Long
    FieldIndex = m_lFieldIndex
End Property
Public Property Let FieldIndex(NewVal As Long)
    m_lFieldIndex = NewVal
End Property

Public Property Get TypeNumber() As Long
    TypeNumber = m_lTypeNumber
End Property
Public Property Let TypeNumber(NewVal As Long)
    m_lTypeNumber = NewVal
End Property

