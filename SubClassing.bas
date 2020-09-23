Attribute VB_Name = "SubClassing"
Option Explicit

' ==================================================================
' Filename:     SubClassing.bas
' Description:  Implements sub classing for a form / object
' ------------------------------------------------------------------
' Created by:   Nick Davis      Date: Sep-00
' Updated by:                   Date:
' ------------------------------------------------------------------
' Notes:
' The functions in this module allow the windows messages sent to
' the application to be intercepted and processed.
'===================================================================


Sub SubClass(objSubClass As Object, ptrObject As ISubClass)
' objSubClass = object to be subclassed
' ptrObject = object that contains message handler we want to use

Dim llOldWindowProc As Long
Dim llObjPtr As Long
    
    'Get a pointer to the object that contains the message handler...
    llObjPtr = ObjPtr(ptrObject)
    
    'Change the window procedure for the object we are subclassing to our procedure...
    llOldWindowProc = SetWindowLong(objSubClass.hwnd, GWL_WNDPROC, AddressOf NewWindowProcL)
    
    'Store the information we will need later
    Call SetProp(objSubClass.hwnd, "ObjPtr", llObjPtr)
    Call SetProp(objSubClass.hwnd, "OldWindowProc", llOldWindowProc)
    
End Sub

Sub UnSubClass(objSubClass As Object)
    
Dim llOldWindowProc As Long
    
    'get the old window procedure address from user data area
    llOldWindowProc = GetProp(objSubClass.hwnd, "OldWindowProc")
    
    If llOldWindowProc Then
        Call SetWindowLong(objSubClass.hwnd, GWL_WNDPROC, llOldWindowProc)
    End If
    
    'TidyUp
    Call RemoveProp(objSubClass.hwnd, "ObjPtr")
    Call RemoveProp(objSubClass.hwnd, "OldWindowProc")
    
End Sub

Function NewWindowProcL(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim llWindowAddress As Long
Dim MessageHandler As ISubClass
Dim llOldWindowProc As Long
    
    'Get a pointer to the object that contains the message handler...
    llWindowAddress = GetProp(hwnd, "ObjPtr")
    llOldWindowProc = GetProp(hwnd, "OldWindowProc")
    
    ' Copy the pointer over our local object so the
    ' local object now points to object that has the message handler...
    CopyMemory MessageHandler, llWindowAddress, 4
    NewWindowProcL = MessageHandler.WindowProcL(llOldWindowProc, hwnd, Msg, wParam, lParam)
    CopyMemory MessageHandler, 0&, 4
    
End Function
