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
'**  This class is a very simplified wrapper for the regular expression engine VBA StaticRegex.                                               **
'**  It is used to be able to port Mendeley Macros to macOS.                                                                                  **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************



'private properties (standard from "VBScript.RegExp" for compatibility)
Private strPattern As String
Private blnIgnoreCase As Boolean
Private blnGlobalSearch As Boolean

'private object for regular expressions
Private objRegularExpression As srRegexTy



Private Sub Class_Initialize()
    'initializes the properties
    strPattern = ""
    blnIgnoreCase = False
    blnGlobalSearch = True

    'initializes the object for regular expressions with new properties
    InitializeRegex
End Sub

Private Sub Class_Terminate()

End Sub



'for Pattern
Public Property Get pattern() As String
   pattern = strPattern
End Property
Public Property Let pattern(strNewPattern As String)
    'sets the pattern
    strPattern = strNewPattern
    'initializes the object for regular expressions with new property
    'InitializeRegex
End Property

'for IgnoreCase
Public Property Get ignoreCase() As Boolean
    ignoreCase = blnIgnoreCase
End Property
Public Property Let ignoreCase(blnNewIgnoreCase As Boolean)
    'sets case insensitivity
    blnIgnoreCase = blnNewIgnoreCase
    'initializes the object for regular expressions with new property
    'InitializeRegex
End Property

'for GlobalSearch
Public Property Get GlobalSearch() As Boolean
    GlobalSearch = blnGlobalSearch
End Property
Public Property Let GlobalSearch(blnNewGlobalSearch As Boolean)
    'sets global applicability
    blnGlobalSearch = blnNewGlobalSearch
End Property



Private Sub InitializeRegex()
    'initializes the object for regular expressions
    srInitializeRegex objRegularExpression, strPattern, blnIgnoreCase
End Sub



Public Function Test(strString As String) As Boolean
    'initializes the object for regular expressions with current properties
    InitializeRegex

    'tests and returns the result
    Test = srTest(objRegularExpression, strString)
End Function

Public Function Execute(strString As String) As Collection
    'used to store the new collection of matches
    Dim colAllCustomMatches As New Collection
    Dim customMatch As GAUG_Match

    'needed by Static Regex engine by Sihlfall
    Dim srMatcherState As srMatcherStateTy

    'initializes the object for regular expressions with current properties
    InitializeRegex

    'if global search is active, we get all matches
    If blnGlobalSearch Then
        'gets all matches (if any)
        Do While srMatchNext(srMatcherState, objRegularExpression, strString)
            'creates new custom match object
            Set customMatch = New GAUG_Match
            'populates new custom match object
            With srMatcherState.captures.entireMatch
                If .Length > 0 Then
                    customMatch.Initialize lngNewFirstIndex:=.start - 1, lngNewLength:=.Length, strNewValue:=Mid$(strString, .start, .Length)
                Else
                    customMatch.Initialize lngNewFirstIndex:=0, lngNewLength:=0, strNewValue:=vbNullString
                End If
            End With
            'adds the new custom match to the collection
            colAllCustomMatches.Add customMatch
        Loop

    'if global search is not active, we get only one match
    Else
        'gets the first match (if any)
        If srMatch(srMatcherState, objRegularExpression, strString) Then
            'creates new custom match object
            Set customMatch = New GAUG_Match
            'populates new custom match object
            With srMatcherState.captures.entireMatch
                If .Length > 0 Then
                    customMatch.Initialize lngNewFirstIndex:=.start, lngNewLength:=.Length, strNewValue:=Mid$(strString, .start, .Length)
                Else
                    customMatch.Initialize lngNewFirstIndex:=0, lngNewLength:=0, strNewValue:=vbNullString
                End If
            End With
            'adds the new custom match to the collection
            colAllCustomMatches.Add customMatch
        End If

    End If 'if global search is active, we get all matches

    'returns the collection of matches
    Set Execute = colAllCustomMatches
End Function

