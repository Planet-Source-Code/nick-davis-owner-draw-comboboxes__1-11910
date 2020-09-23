VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Owner Draw Combo Demo"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraText 
      Caption         =   "Font Style"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Frame fraFill 
      Caption         =   "Fill Style"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame fraLine 
      Caption         =   "Line Style"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Frame fraBitmap 
      Caption         =   "Bitmap Example"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame fraColour 
      Caption         =   "Colour"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin ComctlLib.ImageList imglst_Bitmaps 
      Left            =   2040
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":021E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":043C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":065A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0878
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0A96
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISubClass

' ==================================================================
' Filename:     frmMain.frm
' Description:  OwnerDraw Combo dialog
' ------------------------------------------------------------------
' Created by:   Nicholas Davis      Date: 04-Oct-00
' Updated by:
' ------------------------------------------------------------------
' Notes:
'
'
'===================================================================

' IDs for Colour combo boxes
Private Const ID_COLOUR = 30001

' IDs for Line style combo boxes
Private Const ID_LINESTYLE = 30101

' IDs for Fill style combo boxes
Private Const ID_FILLSTYLE = 30201

' IDs for Text style combo boxes
Private Const ID_TEXTSTYLE = 30301

Private Const ID_BITMAP = 30401

' Store ComboBox window handles
Dim hWndColour As Long
Dim hWndLineStyle As Long
Dim hWndFillStyle As Long
Dim hWndTextStyle As Long
Dim hWndBitmap As Long

' Flag to start processing messages
Dim bFormInitialised As Boolean


Private Sub Form_Load()
SubClass fraColour, Me        ' Sub class the frames to trap Windows
SubClass fraLine, Me          ' messages for the combo boxes
SubClass fraFill, Me
SubClass fraText, Me
SubClass fraBitmap, Me
    
    SetComboColours     ' Set the combo colours based on the palette file
    SetComboFonts       ' Set the combo fonts based on the font file
    
    ' Create the Colour combo boxes
    Call CreateColourCombo(fraColour.hWnd, hWndColour, ID_COLOUR, "Colour", 15, 120)
    
    ' Create the Line style combo boxes
    Call CreateLineCombo(fraLine.hWnd, hWndLineStyle, ID_LINESTYLE, "Line Style", 15, 120)
    
    ' Create the Fill style combo boxes
    Call CreateFillCombo(fraFill.hWnd, hWndFillStyle, ID_FILLSTYLE, "Fill Style", 15, 120)
    
    'Create the Font style combo boxes
    Call CreateTextCombo(fraText.hWnd, hWndTextStyle, ID_TEXTSTYLE, "Text Style", 15, 120)

    'Create the Bitmap style combo boxes
    Call CreateBitmapCombo(fraBitmap.hWnd, hWndBitmap, ID_BITMAP, "Bitmap", 15, 120)

' Finished loading, start manually processing combo box messages
bFormInitialised = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnSubClass Me

    ' Destroy the ComboBoxes we created
    DestroyWindow hWndColour
    DestroyWindow hWndLineStyle
    DestroyWindow hWndFillStyle
    DestroyWindow hWndTextStyle
    DestroyWindow hWndBitmap
        
End Sub

' This is our WinProc function to process messages (namely the DRAWITEM messages for
' drawing our owner draw combo boxes
Private Function ISubClass_WindowProcL(ByVal ptrOldWindowProc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' For some reason (which I cant figure out) the app (along with the IDE)
    ' crashes if you try to process this message before the form has finshed loading
    If bFormInitialised Then
    
        Select Case Msg
            Case WM_DRAWITEM
                
                ' If the message is correctly processed by this function quit the WinProc
                DrawComboItems hWnd, wParam, lParam
                ISubClass_WindowProcL = 0
                Exit Function
            
            ' Trap the command when the combo item is changed
            Case WM_COMMAND
                Dim loWord As Integer, hiWord As Integer
                    
                    ' Get the hiWord and loWord values
                    Call SplitParam(wParam, hiWord, loWord)
                    ' The hiWord is the command, is this notification
                    ' that a different combo item has been selected?
                    
                    If hiWord = CBN_SELCHANGE Then
                        ' The loWord contains the control ID
                        Select Case loWord
                            
                            Case ID_COLOUR
                                cboColour_Change
                            Case ID_LINESTYLE
                                cboLineStyle_Change
                            Case ID_FILLSTYLE
                                cboFillStyle_Change
                            Case ID_TEXTSTYLE
                                cboTextStyle_Change
                            
                        End Select
                    End If
        
        End Select
    
    End If
    
    'direct messages to correct WinProc
    ISubClass_WindowProcL = CallWindowProc(ptrOldWindowProc, hWnd, Msg, wParam, lParam)

End Function

Private Sub cboColour_Change()
' Process whatever you want
'Dim curIndex As Long

    ' Find the current selected index and assign
    'curIndex = SendMessage(hWndColour, CB_GETCURSEL, 0, 0)

End Sub

Private Sub cboLineStyle_Change()
' Process whatever you want
'Dim curIndex As Long

    ' Find the current selected index and assign
    'curIndex = SendMessage(hWndLineStyle, CB_GETCURSEL, 0, 0)

End Sub

Private Sub cboFillStyle_Change()
' Process whatever you want
'Dim curIndex As Long

    ' Find the current selected index and assign
    'curIndex = SendMessage(hWndFillStyle, CB_GETCURSEL, 0, 0)

End Sub

Private Sub cboTextStyle_Change()
' Process whatever you want
'Dim curIndex As Long

    ' Find the current selected index and assign
    'curIndex = SendMessage(hWndTextStyle, CB_GETCURSEL, 0, 0)

End Sub
