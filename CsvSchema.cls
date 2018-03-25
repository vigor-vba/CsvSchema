VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#############################################################################
' CsvSchema: schema.ini creator.
' MIT License | github.com/vigor-vba/CsvSchema
'#############################################################################
Option Explicit
#Const ISDEBUG = False
Private Const MAJOR_VERSION As Long = 1
Private Const MINOR_VERSION As Long = 0
Private Const PATCH_VERSION As Long = 0

'-----------------------------------------------------------------------------
' Enums
'-----------------------------------------------------------------------------
' FormatType
Public Enum FormatType
    fmCSVDelimited
    fmTabDelimited
    fmCustomDelimited
End Enum

' CharsetType
Public Enum CharsetType
    chSJIS = 932    ' Shift-JIS
    chUTF8 = 65001  ' UTF-8
    chEUC = 51932   ' EUC(jp)
End Enum

' DataType
Public Enum DataType
    dtBit
    dtByte
    dtShort
    dtLong
    dtCurrency
    dtSingle
    dtDouble
    dtDateTime
    dtText
    dtMemo
    dtFloat
    dtInteger
    dtLongChar
End Enum

'-----------------------------------------------------------------------------
' Consts
'-----------------------------------------------------------------------------
Private Const SCHEMAFILE As String = "schema.ini"

'-----------------------------------------------------------------------------
' Fields
'-----------------------------------------------------------------------------
Private entry_      As String
Private fields_      As Collection
Private destination_ As String
Private autoSweep_   As Boolean
Private iniFileName  As String
Private iniCharset   As String
Private iniHeader    As String
Private iniFormat    As String


'*****************************************************************************
' Properties
'*****************************************************************************
' Version
Public Property Get Version() As String
    Version = MAJOR_VERSION & "." & MINOR_VERSION & "." & PATCH_VERSION
End Property


'*****************************************************************************
' Class_Initialize()
'*****************************************************************************
Private Sub Class_Initialize()
    entry_ = ""
    Set fields_ = New Collection
    autoSweep_ = False

    Call Me.ColNameHeader(True)
    Call Me.Format(FormatType.fmCSVDelimited)
End Sub


'*****************************************************************************
' Class_Terminate()
'*****************************************************************************
Private Sub Class_Terminate()
    If autoSweep_ Then
        Call Me.Sweep
    End If
End Sub


'*****************************************************************************
' [+] Entry()
'*****************************************************************************
Public Function Entry(FileName As String) As CsvSchema
    entry_ = Trim(FileName)
    iniFileName = "[" & entry_ & "]"

    Set Entry = Me
End Function


'*****************************************************************************
' [+] ColNameHeader()
'*****************************************************************************
Public Function ColNameHeader(Optional HasHeader As Boolean = True) As CsvSchema
    iniHeader = "ColNameHeader=" & IIf(HasHeader, "True", "False")

    Set ColNameHeader = Me
End Function


'*****************************************************************************
' [+] Format()
'*****************************************************************************
Public Function Format(ByVal FileFormat As FormatType, Optional ByVal CustomDelimiter As String) As CsvSchema
    Select Case FileFormat
        Case fmCSVDelimited
            iniFormat = "Format=CSVDelimited"
        Case fmTabDelimited
            iniFormat = "Format=TabDelimited"
        Case fmCustomDelimited
            iniFormat = "Format=Delimited(" & Left(Trim(CustomDelimiter), 1) & ")"
    End Select

    Set Format = Me
End Function


'*****************************************************************************
' [+] CharacterSet()
'*****************************************************************************
Public Function CharacterSet(Charset As CharsetType) As CsvSchema
    iniCharset = "CharacterSet=" & Charset

    Set CharacterSet = Me
End Function


'*****************************************************************************
' [+] AddField()
'*****************************************************************************
Public Function AddField(FieldName As String, Optional FieldType As DataType = DataType.dtText, Optional FieldWidth As Long = 0) As CsvSchema
    Dim fldToken As String

    ' Create field token
    fldToken = Trim(FieldName)
    Select Case FieldType
        ' Bit field
        Case DataType.dtBit
            fldToken = fldToken & " Bit"

        ' Byte field
        Case DataType.dtByte
            fldToken = fldToken & " Byte"

        ' Short field
        Case DataType.dtShort
            fldToken = fldToken & " Short"

        ' Long field
        Case DataType.dtLong
            fldToken = fldToken & " Long"

        ' Currency field
        Case DataType.dtCurrency
            fldToken = fldToken & " Currency"

        ' Single field
        Case DataType.dtSingle
            fldToken = fldToken & " Single"

        ' Double field
        Case DataType.dtDouble
            fldToken = fldToken & " Double"

        ' DateTime field
        Case DataType.dtDateTime
            fldToken = fldToken & " DateTime"

        ' Text field
        Case DataType.dtText
            fldToken = fldToken & " Text"
            If FieldWidth > 0 Then
                fldToken = fldToken & " Width " & FieldWidth
            End If

        ' Memo field
        Case DataType.dtMemo
            fldToken = fldToken & " Memo"
            If FieldWidth > 0 Then
                fldToken = fldToken & " Width " & FieldWidth
            End If

        ' Float field (same as Double)
        Case DataType.dtFloat
            fldToken = fldToken & " Float"

        ' Integer field (same as Short)
        Case DataType.dtInteger
            fldToken = fldToken & " Integer"

        ' LongChar field (same as Memo)
        Case DataType.dtLongChar
            fldToken = fldToken & " LongChar"
            If FieldWidth > 0 Then
                fldToken = fldToken & " Width " & FieldWidth
            End If

    End Select

    ' Add field token to fields_
    Call fields_.Add(fldToken)

    Set AddField = Me
End Function


'*****************************************************************************
' [+] Build()
'*****************************************************************************
Public Function Build() As String
    Dim val       As String
    Dim cntFields As Long
    Dim c         As Long

    ' FileName
    val = iniFileName

    ' ColNameHeader
    If Len(iniHeader) > 0 Then
        val = val & vbCrLf & iniHeader
    End If

    ' CharacterSet
    If Len(iniCharset) > 0 Then
        val = val & vbCrLf & iniCharset
    End If

    ' Format
    If Len(iniFormat) > 0 Then
        val = val & vbCrLf & iniFormat
    End If

    ' Fields
    cntFields = fields_.Count
    For c = 1 To cntFields
        val = val & vbCrLf & "Col" & c & "=" & fields_(c)
    Next c

    Build = val
End Function


'*****************************************************************************
' [+] Create()
'*****************************************************************************
Public Function Create(Optional AutoSweep As Boolean = False, Optional Destination As String) As CsvSchema
    Dim fso As Object ' Scripting.FileSystemObject

    autoSweep_ = AutoSweep
    destination_ = Trim(Destination)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Resolve destination
    If Len(destination_) > 0 Then
        destination_ = destination_ & "\" & SCHEMAFILE
    Else
        destination_ = fso.GetParentFolderName(entry_) & "\" & SCHEMAFILE
    End If

    ' Create schema file
    With fso.OpenTextFile(destination_, 2, True) ' 2:ForWriting
        Call .Write(Me.Build())
        .Close
    End With

    Set fso = Nothing

    Set Create = Me
End Function


'*****************************************************************************
' [+] Sweep()
'*****************************************************************************
Public Sub Sweep()
    On Error Resume Next
    Call Kill(destination_)
    On Error GoTo 0
End Sub
