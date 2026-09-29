VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSourceParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSourceParser
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsSourceParser content loading, output, and sanitization.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Core")
'@Tag("unit")

Private m_parser As clsSourceParser


Private Sub Class_Initialize()
    Set m_parser = New clsSourceParser
End Sub


Private Sub Class_Terminate()
    Set m_parser = Nothing
End Sub


Public Sub TestLoadString_ReturnsOutput()
    Dim strContent As String
    strContent = "Begin Form" & vbCrLf & "  Caption =""Test""" & vbCrLf & "End"
    m_parser.LoadString strContent, edbForm
    m_parser.Sanitize ectObjectDefinition
    TestAssert Len(m_parser.GetOutput) > 0, "GetOutput returns content"
End Sub


Public Sub TestHash_Deterministic()
    Dim strContent As String
    strContent = "Some content for hashing"
    m_parser.LoadString strContent, edbModule
    Dim strHash1 As String
    strHash1 = m_parser.Hash()
    m_parser.LoadString strContent, edbModule
    TestAssert m_parser.Hash() = strHash1, "same content produces same hash"
End Sub


Public Sub TestStripClassHeader()
    Dim strContent As String
    Dim strResult As String
    strContent = "VERSION 1.0 CLASS" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "  MultiUse = -1" & vbCrLf & _
        "END" & vbCrLf & _
        "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database"
    strResult = m_parser.StripClassHeader(strContent, False)
    TestAssert InStr(strResult, "VERSION 1.0 CLASS") = 0, "class header stripped"
    TestAssert InStr(strResult, "Option Compare Database") > 0, "code content preserved"
End Sub


Public Sub TestStripClassHeader_NameOnly()
    Dim strContent As String
    Dim strResult As String
    strContent = "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database" & vbCrLf & _
        "Option Explicit"
    strResult = m_parser.StripClassHeader(strContent, True)
    TestAssert InStr(strResult, "VB_Name") = 0, "VB_Name attribute stripped"
    TestAssert InStr(strResult, "Option Compare") > 0, "code preserved"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSanitizeXml_KeepsWhitespaceValues
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/29/2026
' Purpose   : A table data value that is only whitespace, or that has tabs at the start
'           : of a line, is data and has to come out of the sanitizer as it went in. Both
'           : formatters are covered: the XSLT one below 30,000 characters of input, and
'           : the one for large documents above, where a padding row pushes it over.
'---------------------------------------------------------------------------------------
'
Public Sub TestSanitizeXml_KeepsWhitespaceValues()

    Dim strOut As String
    Dim intPass As Integer
    Dim strPath As String

    For intPass = 1 To 2
        strPath = IIf(intPass = 1, "XSLT", "large")
        m_parser.LoadString WhitespaceDataXml(intPass = 2), edbTableData
        strOut = m_parser.Sanitize(ectXML)

        TestAssert InStr(1, strOut, "<Note> </Note>", vbBinaryCompare) > 0, _
            strPath & ": a single space is kept"
        TestAssert InStr(1, strOut, "<Note>x" & vbCrLf & vbTab & "y</Note>", vbBinaryCompare) > 0, _
            strPath & ": a tab at the start of a line is kept"
        TestAssert InStr(1, strOut, "<Note>" & vbTab & "z</Note>", vbBinaryCompare) > 0, _
            strPath & ": a tab at the start of the value is kept"
        TestAssert InStr(1, strOut, "<Note>  m  </Note>", vbBinaryCompare) > 0, _
            strPath & ": spaces around a value are kept"
        TestAssert InStr(1, strOut, vbCrLf & "  <tbl>" & vbCrLf & "    <ID>1</ID>" & vbCrLf, _
            vbBinaryCompare) > 0, strPath & ": rows and fields are still laid out"
    Next intPass

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSanitizeXml_LargeLayoutMatchesXslt
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/29/2026
' Purpose   : The two formatters exist only because XSLT runs out of memory on large
'           : documents. Whichever one runs, the same rows have to be laid out the same,
'           : or a table's file would change shape the day it crosses the size limit.
'---------------------------------------------------------------------------------------
'
Public Sub TestSanitizeXml_LargeLayoutMatchesXslt()

    Dim strSmall As String
    Dim strLarge As String
    Dim lngPad As Long

    m_parser.LoadString WhitespaceDataXml(False), edbTableData
    strSmall = m_parser.Sanitize(ectXML)
    m_parser.LoadString WhitespaceDataXml(True), edbTableData
    strLarge = m_parser.Sanitize(ectXML)

    ' Drop the padding row from the large output, and compare the rest.
    lngPad = InStr(1, strLarge, "  <tbl>" & vbCrLf & "    <ID>99</ID>", vbBinaryCompare)
    TestAssert lngPad > 0, "padding row found in the large output"
    If lngPad = 0 Then Exit Sub
    strLarge = Left$(strLarge, lngPad - 1) & _
        Mid$(strLarge, InStr(lngPad, strLarge, "</tbl>" & vbCrLf, vbBinaryCompare) + 8)
    TestAssert StrComp(strLarge, strSmall, vbBinaryCompare) = 0, _
        "large and small documents share one layout"

End Sub


'---------------------------------------------------------------------------------------
' Function  : WhitespaceDataXml
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/29/2026
' Purpose   : Table data as Application.ExportXML writes it -- one element per line and
'           : no indentation -- with values that are made of or start with whitespace.
'---------------------------------------------------------------------------------------
'
Private Function WhitespaceDataXml(blnLarge As Boolean) As String

    Dim cXml As clsConcat

    Set cXml = New clsConcat
    cXml.AppendOnAdd = vbCrLf
    cXml.Add "<?xml version=""1.0"" encoding=""UTF-8""?>"
    cXml.Add "<dataroot xmlns:od=""urn:schemas-microsoft-com:officedata"" generated=""2026-09-29T10:00:00"">"
    cXml.Add "<tbl>", vbCrLf, "<ID>1</ID>", vbCrLf, "<Note> </Note>", vbCrLf, "</tbl>"
    cXml.Add "<tbl>", vbCrLf, "<ID>2</ID>", vbCrLf, "<Note>x", vbCrLf, vbTab, "y</Note>", vbCrLf, "</tbl>"
    cXml.Add "<tbl>", vbCrLf, "<ID>3</ID>", vbCrLf, "<Note>", vbTab, "z</Note>", vbCrLf, "</tbl>"
    cXml.Add "<tbl>", vbCrLf, "<ID>4</ID>", vbCrLf, "<Note>  m  </Note>", vbCrLf, "</tbl>"
    If blnLarge Then
        cXml.Add "<tbl>", vbCrLf, "<ID>99</ID>", vbCrLf, "<Note>", String$(30000, "p"), "</Note>", vbCrLf, "</tbl>"
    End If
    cXml.Add "</dataroot>"

    WhitespaceDataXml = cXml.GetStr

End Function
