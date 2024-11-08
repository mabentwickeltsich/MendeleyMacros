VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GAUG_RegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-23                                                                                                                **
'**                                                                                                                                           **
'**  Class GAUG_RegExp                                                                                                                        **
'**                                                                                                                                           **
'**  This class is a very simplified wrapper for the original class RegExp provided by Microsoft VBScript Regular Expressions 5.5.            **
'**  It is used to remove the dependency on Microsoft VBScript Regular Expressions 5.5 on Windows                                             **
'**     and to be able to port Mendeley Macros to macOS.                                                                                      **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************



'private properties (standard from "VBScript.RegExp" for compatibility)
Private strPattern As String
Private blnIgnoreCase As Boolean
Private blnGlobalSearch As Boolean

'private object for regular expressions
Private objRegularExpression As Object



Private Sub Class_Initialize()
    'creates the object for the regular expressions
    Set objRegularExpression = CreateObject("VBScript.RegExp")
    'initializes the properties
    strPattern = ""
    blnIgnoreCase = False
    blnGlobalSearch = True

    'initializes the object for regular expressions with new properties
    objRegularExpression.pattern = strPattern
    objRegularExpression.ignoreCase = blnIgnoreCase
    objRegularExpression.Global = blnGlobalSearch
End Sub

Private Sub Class_Terminate()
    'frees memory
    Set objRegularExpression = Nothing
End Sub



'for Pattern
Public Property Get pattern() As String
   pattern = strPattern
End Property
Public Property Let pattern(strNewPattern As String)
    'sets the pattern
    strPattern = strNewPattern
    objRegularExpression.pattern = strPattern
End Property

'for IgnoreCase
Public Property Get ignoreCase() As Boolean
    ignoreCase = blnIgnoreCase
End Property
Public Property Let ignoreCase(blnNewIgnoreCase As Boolean)
    'sets case insensitivity
    blnIgnoreCase = blnNewIgnoreCase
    objRegularExpression.ignoreCase = blnIgnoreCase
End Property

'for GlobalSearch
Public Property Get GlobalSearch() As Boolean
    GlobalSearch = blnGlobalSearch
End Property
Public Property Let GlobalSearch(blnNewGlobalSearch As Boolean)
    'sets global applicability
    blnGlobalSearch = blnNewGlobalSearch
    objRegularExpression.Global = blnGlobalSearch
End Property



Public Function Test(strString As String) As Boolean
    'tests and returns the result
    Test = objRegularExpression.Test(strString)
End Function

Public Function Execute(strString As String) As Collection
    'used to store the original collection of matches
    Dim allMatches As Object
    Dim currentMatch As Object

    'used to store the new collection of matches
    Dim colAllCustomMatches As New Collection
    Dim customMatch As GAUG_Match

    'executes the RegExp pattern against the input string
    Set allMatches = objRegularExpression.Execute(strString)

    'if matches exist
    If allMatches.Count > 0 Then
        For Each currentMatch In allMatches
            'creates new custom match object and populates new custom match object
            Set customMatch = New GAUG_Match
            customMatch.Initialize lngNewFirstIndex:=currentMatch.FirstIndex, lngNewLength:=currentMatch.Length, strNewValue:=currentMatch.value
            'adds the new custom match to the collection
            colAllCustomMatches.Add customMatch
        Next currentMatch
    End If

    'clears the original collection of matches
    Set allMatches = Nothing

    'returns the collection of matches
    Set Execute = colAllCustomMatches
End Function

