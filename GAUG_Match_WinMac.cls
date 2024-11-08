VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GAUG_Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-23                                                                                                                **
'**                                                                                                                                           **
'**  Class GAUG_Match                                                                                                                         **
'**                                                                                                                                           **
'**  This class is a very simplified wrapper for the original class Match provided by Microsoft VBScript Regular Expressions 5.5.             **
'**  It is used to remove the dependency on Microsoft VBScript Regular Expressions 5.5 on Windows                                             **
'**     and to be able to port Mendeley Macros to macOS.                                                                                      **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************



'private properties (standard from "VBScript.Match" for compatibility)
Private lngFirstIndex As Long
Private lngLength As Long
Private strValue As String



Private Sub Class_Initialize()
    'initializes the values
    lngFirstIndex = 0
    lngLength = 0
    strValue = ""
End Sub

Private Sub Class_Terminate()

End Sub



'for FirstIndex
Public Property Get FirstIndex() As Long
   FirstIndex = lngFirstIndex
End Property

'for Length
Public Property Get Length() As Long
   Length = lngLength
End Property

'for Value
Public Property Get value() As String
   value = strValue
End Property



Public Sub Initialize(lngNewFirstIndex As Long, lngNewLength As Long, strNewValue As String)
    'initializes the values
    lngFirstIndex = lngNewFirstIndex
    lngLength = lngNewLength
    strValue = strNewValue
End Sub

