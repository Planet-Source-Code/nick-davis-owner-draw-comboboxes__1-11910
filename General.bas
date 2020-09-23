Attribute VB_Name = "General"
Option Explicit

' ==================================================================
' Filename:     General.bas
' Description:  General application functions
' ------------------------------------------------------------------
' Created by:   Nick Davis      Date: Sep-00
' Updated by:                   Date:
' ------------------------------------------------------------------
' Notes: General functions and structures required for the app.
'
'
'===================================================================

' ************************************************
'           API STRUCTURE DEFINITIONS
' ************************************************

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
    CtrlType As Long
    CtrlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hWndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

' ************************************************
'       USER DEFINED STRUCTURE DEFINITIONS
' ************************************************

Public Type FONTSTRUCT
    FontName As String
    FontWeight As Integer
    FontStyle As Integer
    FontUnderline As Integer
End Type

' ************************************************
'       USER DEFINED COLOUR PALETTE ENTRIES
' ************************************************

Public Const llRed = 181
Public Const llGreen = 31
Public Const llBlue = 6
Public Const llYellow = 211

' This function splits a message parameter into a Hi and a Lo Word
' This is useful for processing messages
Public Sub SplitParam(ByVal param As Long, ByRef hiWord As Integer, ByRef loWord As Integer)
   
    CopyMemory loWord, param, Len(loWord)
    hiWord = (param / (2 ^ 16))

End Sub
' This function takes two integers and makes a long word to pass as parameters
Public Function MakeLongWord(ByVal hiWord As Integer, ByVal loWord As Integer) As Long
    
    MakeLongWord = (hiWord * (2 ^ 16)) + loWord

End Function
