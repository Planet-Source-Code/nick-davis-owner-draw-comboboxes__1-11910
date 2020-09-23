Attribute VB_Name = "ComboBox"
Option Explicit

' ==================================================================
' Filename:     ComboBox.bas
' Description:  Ownerdraw combo box functions
' ------------------------------------------------------------------
' Created by:   Nicholas Davis      Date: 06-Sep-00
' Updated by:
' ------------------------------------------------------------------
' Notes:
'
'
' ==================================================================

' Dynamic arrays to store colour and font data
Public colours() As Long
Public fonts() As FONTSTRUCT

' Store for the colours in the combo box.
Public cboColours() As Long

Public Function DrawComboItems(hWnd As Long, wParam As Long, lParam As Long) As Boolean
    
Dim dsItem As DRAWITEMSTRUCT        ' Structure passed as lParam
Dim bStr() As Byte                  ' Byte array to hold a string
Dim hBrush, hPen, hOldPen As Long   ' Drawing pen/brush handles
Dim hFont, hOldFont As Long         ' Font handles
Dim buffLen As Integer              ' Length of string buffer
Dim rcRect As RECT                  ' Highlighting bound rectangle
Dim yPos As Integer                 ' Position to draw lines
                    
    ' Copy the pointer to the DRAWITEMSTRUCT in the lParam
    '  to a local variable ready to process the data
    CopyMemory dsItem, ByVal lParam, Len(dsItem)
                
    ' Choose the action to take based on the Control ID of the combo box
        
    If dsItem.CtrlID < 30100 Then   ' Colour combo box
        ' Colour the combo item according to the colour
        ' reference stored in the itemData member
        hBrush = CreateSolidBrush(dsItem.itemData)
        FillRect dsItem.hdc, dsItem.rcItem, hBrush
        DeleteObject hBrush
            
    ElseIf dsItem.CtrlID < 30200 Then  ' Edge/Line style
        ' Set the background colour
        SetBkColor dsItem.hdc, RGB(255, 255, 255)
        ' Fill the combo item with white
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        FillRect dsItem.hdc, dsItem.rcItem, hBrush
        DeleteObject hBrush
        ' Create the pen
        hPen = CreatePen(dsItem.itemData, 1, RGB(0, 0, 0))
        hOldPen = SelectObject(dsItem.hdc, hPen)
        ' Draw the line. For some reason when a line width > 1 is specified
        ' just a solid line is drawn
        For yPos = 9 To 10 Step 1
            MoveToEx dsItem.hdc, dsItem.rcItem.Left, dsItem.rcItem.Top + yPos, 0&
            LineTo dsItem.hdc, dsItem.rcItem.Right, dsItem.rcItem.Top + yPos
        Next yPos
        ' Set the old Pen back and delete this one
        SelectObject dsItem.hdc, hOldPen
        DeleteObject hPen
            
    ElseIf dsItem.CtrlID < 30300 Then ' Fill Style
        ' Set the background colour
        SetBkColor dsItem.hdc, RGB(255, 255, 255)
        ' Either the fill style is the user defined HS_NOFILL, HS_SOLID
        ' or a correct Windows brush style
        If dsItem.itemData = HS_NOFILL Then
            ' Fill the combo item with white
            hBrush = CreateSolidBrush(RGB(255, 255, 255))
            FillRect dsItem.hdc, dsItem.rcItem, hBrush
            DeleteObject hBrush
        ElseIf dsItem.itemData = HS_SOLID Then
            ' Draw a filled rectangle
            hBrush = CreateSolidBrush(RGB(80, 80, 80))
            FillRect dsItem.hdc, dsItem.rcItem, hBrush
            DeleteObject hBrush
        Else
            ' Draw a hatch pattern
            hBrush = CreateHatchBrush(dsItem.itemData, RGB(0, 0, 0))
            FillRect dsItem.hdc, dsItem.rcItem, hBrush
            DeleteObject hBrush
        End If
        
    ElseIf dsItem.CtrlID < 30400 Then   ' Font style
        ' Set the background colour to white unless we're drawing the edit box
        ' with the focus. I have defined ODS_FOCUSONEDITBOX myself as it is
        ' not documented or defined in windows.h??
        If Not dsItem.itemState = ODS_FOCUSONEDITBOX Then SetBkColor dsItem.hdc, RGB(255, 255, 255)
        ' Fill the combo item with white
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        FillRect dsItem.hdc, dsItem.rcItem, hBrush
        DeleteObject hBrush
        
        ' Get the length of the item's string, store it in our
        ' byte array so we can write the text to the item on screen
        buffLen = Len(fonts(dsItem.itemID).FontName)
        bStr() = StrConv(fonts(dsItem.itemID).FontName, vbFromUnicode)
        ' Set the bStr array to a fixed length to help the CreateFont function find the
        ' font name properly
        ReDim Preserve bStr(buffLen)
        
        ' Create the font and select it into the device context
        hFont = CreateFont(15, 0, 0, 0, fonts(dsItem.itemID).FontWeight, fonts(dsItem.itemID).FontStyle, fonts(dsItem.itemID).FontUnderline, 0, ANSI_CHARSET, 0, 0, 0, 0, VarPtr(bStr(0)))
        hOldFont = SelectObject(dsItem.hdc, hFont)
                        
        ' Draw the text
        TextOutBStr dsItem.hdc, dsItem.rcItem.Left + 5, dsItem.rcItem.Top, bStr(0), buffLen
        ' Reset the font
        SelectObject dsItem.hdc, hOldFont
        DeleteObject hFont
        DeleteObject hOldFont
    
    ElseIf dsItem.CtrlID < 30500 Then   ' Bitmap combo
        ' Fill the combo item with white
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        FillRect dsItem.hdc, dsItem.rcItem, hBrush
        DeleteObject hBrush
    
        ' Draw the image from the bitmaps stored in the image list control
        frmMain.imglst_Bitmaps.ListImages(dsItem.itemData).Draw dsItem.hdc, (dsItem.rcItem.Left * Screen.TwipsPerPixelX) + (15 * Screen.TwipsPerPixelX), (dsItem.rcItem.Top * Screen.TwipsPerPixelY) + (3 * Screen.TwipsPerPixelY)
        
    End If
        
    ' Highlight the item as it receives focus. Again I have defined
    ' ODS_FOCUSITEM, since I cant find it documented or defined in windows.h??
    If dsItem.itemAction = ODA_SELECT And dsItem.itemState = ODS_FOCUSITEM Then
            
        ' Make the rectangle sit 1 pixel inside the item bounds
        With rcRect
            .Bottom = dsItem.rcItem.Bottom - 1
            .Left = dsItem.rcItem.Left + 1
            .Right = dsItem.rcItem.Right - 1
            .Top = dsItem.rcItem.Top + 1
        End With
            
        hBrush = CreateSolidBrush(RGB(0, 0, 0))
        ' Draw two rectangles, one inside the other to give a thicker frame
        FrameRect dsItem.hdc, dsItem.rcItem, hBrush
        FrameRect dsItem.hdc, rcRect, hBrush
        DeleteObject hBrush

    End If
        
    ' Finished processing
    DrawComboItems = True

End Function

Public Sub CreateColourCombo(ByVal hWndParent As Long, ByRef hWndCombo As Long, ByVal nID As Integer, ByVal szWinName As String, ByVal nTop As Integer, ByVal nWidth As Integer)

Dim clrIndex As Integer     ' Variable to create colour reference
Dim nIndex As Integer       ' Item index

    ' Create the combobox as 'Owner Draw' to generate
    ' WM_DRAWITEM messages used in WinProc function
    hWndCombo = CreateWindowEx(0, "COMBOBOX", szWinName, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST _
    Or CBS_OWNERDRAWFIXED, 10, nTop, nWidth, 500, hWndParent, nID, App.hInstance, 0)

    ' Add combo Items. Combo created without CBS_HASSTRINGS constant so
    ' the lParam is stored as item data rather than a string
    
    For clrIndex = 0 To UBound(cboColours) Step 1
        ' Add the RGB value stored in cboColours array
        nIndex = SendMessage(hWndCombo, CB_ADDSTRING, 0&, cboColours(clrIndex))
    Next
    
    ' Select the first combo item into the edit box portion of the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0&, 0&

End Sub

Public Sub CreateLineCombo(ByVal hWndParent As Long, ByRef hWndCombo As Long, ByVal nID As Integer, ByVal szWinName As String, ByVal nTop As Integer, ByVal nWidth As Integer)

Dim llColour As Long        ' Variable to store colour reference
Dim PS_STYLE As Integer     ' Variable to create colour reference
Dim nIndex As Integer       ' Item index

    ' Create the combobox as 'Owner Draw' to generate WM_DRAWITEM messages
    ' used in WinProc function
    hWndCombo = CreateWindowEx(0, "COMBOBOX", szWinName, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST _
    Or CBS_OWNERDRAWFIXED, 10, nTop, nWidth, 300, hWndParent, nID, App.hInstance, 0)
 
    ' Add combo Items. Combo created without CBS_HASSTRINGS constant so
    ' the lParam is stored as item data rather than a string
    
    For PS_STYLE = 0 To 4 Step 1
        nIndex = SendMessage(hWndCombo, CB_ADDSTRING, 0&, PS_STYLE)
        ' Set the item height a little bigger (just to look nice)
        SendMessage hWndCombo, CB_SETITEMHEIGHT, nIndex, (SendMessage(hWndCombo, CB_GETITEMHEIGHT, nIndex, 0) * 1.075)
    Next PS_STYLE
    
    ' Select the first combo item into the edit box portion of the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0&, 0&

End Sub

' Create a combo box to show the available fill styles
Public Sub CreateFillCombo(ByVal hWndParent As Long, ByRef hWndCombo As Long, ByVal nID As Integer, ByVal szWinName As String, ByVal nTop As Integer, ByVal nWidth As Integer)

Dim llColour As Long        ' Variable to store colour reference
Dim HS_STYLE As Integer     ' Variable to create colour reference
Dim nIndex As Integer       ' Item index

    ' Create the combobox as 'Owner Draw' to generate
    ' WM_DRAWITEM messages used in WinProc function
    hWndCombo = CreateWindowEx(0, "COMBOBOX", szWinName, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST _
    Or CBS_OWNERDRAWFIXED, 10, nTop, nWidth, 280, hWndParent, nID, App.hInstance, 0)

    ' Add combo Items. Combo created without CBS_HASSTRINGS constant
    ' so the lParam is stored as item data rather than a string.
    
    ' Create the no fill and solid styles first...
    Call SendMessage(hWndCombo, CB_ADDSTRING, 0&, HS_NOFILL)
    Call SendMessage(hWndCombo, CB_ADDSTRING, 0&, HS_SOLID)
    ' ...then loop to create the other brush styles
    For HS_STYLE = 0 To 5 Step 1
        nIndex = SendMessage(hWndCombo, CB_ADDSTRING, 0&, HS_STYLE)
        ' Set the item height a little bigger (just to look nice)
        SendMessage hWndCombo, CB_SETITEMHEIGHT, nIndex, (SendMessage(hWndCombo, CB_GETITEMHEIGHT, nIndex, 0) * 1.075)
    Next HS_STYLE
    
    ' Select the first combo item into the edit box portion of the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0&, 0&

End Sub

' Create a combo box to show the selection of available fonts
Public Sub CreateTextCombo(ByVal hWndParent As Long, ByRef hWndCombo As Long, ByVal nID As Integer, ByVal szWinName As String, ByVal nTop As Integer, ByVal nWidth As Integer)

Dim bStr() As Byte          ' Byte Array to store string data
Dim nIndex As Integer       ' Item index
Dim nCount As Integer       ' loop variable

    ' Create the combobox as 'Owner Draw' to generate WM_DRAWITEM messages
    ' used in WinProc function
    hWndCombo = CreateWindowEx(0, "COMBOBOX", szWinName, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST _
    Or CBS_OWNERDRAWFIXED, 10, nTop, nWidth, 280, hWndParent, nID, App.hInstance, 0)
    
    ' Store nothing as item data, we'll query the item index and cross
    ' reference this with the values stored in the fonts array when we
    ' want to draw the item
    For nCount = 0 To UBound(fonts) Step 1
        SendMessage hWndCombo, CB_ADDSTRING, 0&, 0&
    Next nCount
    
    ' Set the width of the combo's list box (to fit the full font name)
    SendMessage hWndCombo, CB_SETDROPPEDWIDTH, (SendMessage(hWndCombo, CB_GETDROPPEDWIDTH, nIndex, 0) * 1.5), 0
    
    ' Select the first combo item into the edit box portion of the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0&, 0&

End Sub

Public Sub CreateBitmapCombo(ByVal hWndParent As Long, ByRef hWndCombo As Long, ByVal nID As Integer, ByVal szWinName As String, ByVal nTop As Integer, ByVal nWidth As Integer)

Dim nIndex As Integer       ' Item index
Dim nCount As Integer       ' loop variable

    ' Create the combobox as 'Owner Draw' to generate WM_DRAWITEM messages
    ' used in WinProc function
    hWndCombo = CreateWindowEx(0, "COMBOBOX", szWinName, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST _
    Or CBS_OWNERDRAWFIXED, 10, nTop, nWidth, 280, hWndParent, nID, App.hInstance, 0)
    
    ' Store the various styles as integer values that can be drawn as styles
    ' when the combo box is processed
    For nCount = 1 To 6 Step 1
        nIndex = SendMessage(hWndCombo, CB_ADDSTRING, 0&, nCount)
        ' Set the item height a little bigger (just to look nice)
        SendMessage hWndCombo, CB_SETITEMHEIGHT, nIndex, (SendMessage(hWndCombo, CB_GETITEMHEIGHT, nIndex, 0) * 1.075)
    Next nCount

    ' Select the first combo item into the edit box portion of the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0&, 0&

End Sub

' Sets up the RGB colour values for the combo boxes
Public Sub SetComboColours()

Dim FILE$, a, fso               ' File manipulation variables
Dim fOutput As String           ' File output string buffer
Dim maxEntries As Integer       ' Max number of palette entries
Dim strlen As Integer           ' Length of string buffer
Dim nColourEntries As Integer   ' Number of colours stored
Dim nValue As Integer           ' Used to determine which value is being extracted from a line
Dim Char As String              ' Single character
Dim nIndex As Integer           ' Looping variable

' Palette values
Dim palIndex As Integer, rVal As Integer, gVal As Integer, bVal As Integer
        
    ' Create a filesystemobject and open as a text file the
    ' palette definition file
    Set fso = CreateObject("Scripting.FileSystemObject")
    FILE$ = App.Path & "\tslcolours.dat"
    Set a = fso.OpenTextFile(FILE$, 1)
        
    a.readline                  ' Skip the first two lines to
    a.readline                  ' find the number of palette entries
    maxEntries = a.readline     ' Store the maximum number of palette entries
        
    ' Loop until we get to the end of the file
    Do While Not a.AtEndOfStream
        
        ' Get the line and find its length
        fOutput = a.readline
        strlen = Len(fOutput)
         
        ' Reset all the variable values
        nValue = 1
        palIndex = 0
        rVal = 0
        gVal = 0
        bVal = 0
            
        ' Extract the palette entries from the line
        For nIndex = 1 To strlen Step 1
            Char = Mid(fOutput, nIndex, 1)
            If Not Char = ";" Then
                Select Case nValue
                    ' Is this the palette, R, G or B value?
                    Case 1
                        palIndex = palIndex & Char
                    Case 2
                        rVal = rVal & Char
                    Case 3
                        gVal = gVal & Char
                    Case 4
                        bVal = bVal & Char
                End Select
            Else
                ' Semi-colon. Ignore this character, start storing the next value
                nValue = nValue + 1
            End If
        Next
                
        ' Get the first and last palette entries, plus about 24 or so inbetween
        ' Including the pure Red, Green, Blue plus yellow
        If palIndex = 1 Or palIndex = maxEntries _
        Or palIndex = llRed Or palIndex = llGreen _
        Or palIndex = llBlue Or palIndex = llYellow _
        Or (palIndex Mod 10 = 0) Then
            ' Store the RGB values for the combo boxes to use, plus store
            ' the corresponding palette entry value for use by the editor
            ReDim Preserve cboColours(nColourEntries)
            ReDim Preserve colours(nColourEntries)
            cboColours(nColourEntries) = RGB(rVal, gVal, bVal)
            colours(nColourEntries) = palIndex
            If Not palIndex = maxEntries Then nColourEntries = nColourEntries + 1
        End If
        
    Loop
        
    ' Close the file
    a.Close
    
End Sub


' Sets up the font array values for the combo boxes
Public Sub SetComboFonts()

Dim FILE$, a, fso               ' File manipulation variables
Dim fOutput As String           ' File output string buffer
Dim maxEntries As Integer       ' Max number of palette entries
Dim strlen As Integer           ' Length of string buffer
Dim nValue As Integer           ' Used to determine which value is being extracted from a line
Dim nFontEntries As Integer     ' Number of font entries
Dim Char As String              ' Single character
Dim nIndex As Integer           ' Looping variable

' font values
Dim fontIndex As Integer, FontName As String, FontWeight As Integer, FontStyle As Integer, FontUnderline As Integer
        
    ' Create a filesystemobject and open as a text file the
    ' palette definition file
    Set fso = CreateObject("Scripting.FileSystemObject")
    FILE$ = App.Path & "\tslfonts.dat"
    Set a = fso.OpenTextFile(FILE$, 1)
        
    a.readline                  ' Skip the first two lines to
    a.readline                  ' find the number of font entries
    maxEntries = a.readline     ' Store the maximum number of font entries
        
    ' Loop until we get to the end of the file
    Do While Not a.AtEndOfStream
        
        ' Get the line and find its length
        fOutput = a.readline
        strlen = Len(fOutput)
         
        ' Reset all the variable values
        nValue = 1
        fontIndex = 0
        FontName = ""
        FontWeight = 0
        FontStyle = 0
        FontUnderline = 0
                    
        ' Extract the font data entries from the line
        For nIndex = 1 To strlen Step 1
            Char = Mid(fOutput, nIndex, 1)
            If Not Char = ";" Then
                Select Case nValue
                    ' Is this the font index or the font name?
                    Case 1
                        fontIndex = fontIndex & Char
                    Case 3
                        FontName = FontName & Char
                    Case 4
                        FontWeight = FontWeight & Char
                    Case 5
                        FontStyle = FontStyle & Char
                    Case 6
                        FontUnderline = FontUnderline & Char
                End Select
            Else
                ' Semi-colon. Ignore this character, start storing the next value
                nValue = nValue + 1
            End If
        Next
                
        ReDim Preserve fonts(nFontEntries)
        fonts(nFontEntries).FontName = FontName
        fonts(nFontEntries).FontWeight = FontWeight
        fonts(nFontEntries).FontStyle = FontStyle
        fonts(nFontEntries).FontUnderline = FontUnderline
        If Not fontIndex = maxEntries Then nFontEntries = nFontEntries + 1
        
    Loop
        
    ' Close the file
    a.Close
    
End Sub
