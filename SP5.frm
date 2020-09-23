VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form SP5 
   Caption         =   "SJOTS Perfect 5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "SP5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "Img_Bar"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.ComboBox Size 
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox Fonts 
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1920
   End
   Begin MSComDlg.CommonDialog CommD 
      Left            =   2160
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SP5.frx":0442
   End
   Begin ComctlLib.ImageList Img_Bar 
      Left            =   2160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":061D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":072F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0841
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0953
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0C89
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0D9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SP5.frx":0EAD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Stripe1 
         Caption         =   "-"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Stripe4 
         Caption         =   "-"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu CloseDoc 
         Caption         =   "&Close"
      End
      Begin VB.Menu Stripe2 
         Caption         =   "-"
      End
      Begin VB.Menu PrintDoc 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Stripe3 
         Caption         =   "-"
      End
      Begin VB.Menu Quit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Cut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Layout 
      Caption         =   "&Layout"
      Begin VB.Menu CFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu FColor 
         Caption         =   "Font &Color..."
      End
   End
   Begin VB.Menu Search 
      Caption         =   "&Search"
      Begin VB.Menu SearchFor 
         Caption         =   "Search &For..."
      End
      Begin VB.Menu SNext 
         Caption         =   "&Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu Ask 
      Caption         =   "&?"
      Begin VB.Menu About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "SP5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Dim FNEXT As Integer
Dim Found As String
Dim FName As String
Dim Changed As Boolean

'The following declarations are for the
'printing operation.
Private Type Rect
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Type CharRange
cpMin As Long
cpMax As Long
End Type

Private Type FormatRange
hdc As Long
hdcTarget As Long
rc As Rect
rcPage As Rect
chrg As CharRange
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, _
lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
Dim LeftOffset As Long, TopOffset As Long
Dim LeftMargin As Long, TopMargin As Long
Dim RightMargin As Long, BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawTo As Rect
Dim rcPage As Rect
Dim TextLength As Long
Dim NextCharPosition As Long
Dim r As Long
Printer.Print Space(1)
Printer.ScaleMode = vbTwips
LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
PHYSICALOFFSETX), vbPixels, vbTwips)
TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
PHYSICALOFFSETY), vbPixels, vbTwips)
LeftMargin = LeftMarginWidth - LeftOffset
TopMargin = TopMarginHeight - TopOffset
RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

rcPage.Left = 0
rcPage.Top = 0
rcPage.Right = Printer.ScaleWidth
rcPage.Bottom = Printer.ScaleHeight
rcDrawTo.Left = LeftMargin
rcDrawTo.Top = TopMargin
rcDrawTo.Right = RightMargin
rcDrawTo.Bottom = BottomMargin

fr.hdc = Printer.hdc
fr.hdcTarget = Printer.hdc
fr.rc = rcDrawTo
fr.rcPage = rcPage
fr.chrg.cpMin = 0
fr.chrg.cpMax = -1
TextLength = Len(RTF.Text)
Do
NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
If NextCharPosition >= TextLength Then Exit Do
fr.chrg.cpMin = NextCharPosition
Printer.NewPage
Printer.Print Space(1)
fr.hdc = Printer.hdc
fr.hdcTarget = Printer.hdc
Loop
Printer.EndDoc
r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub

Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub CloseDoc_Click()
If Changed Then
DetectChange
Text.Text = ""
CommD.filename = ""
Changed = False
SP5.Caption = "SJOTS Perfect 5"
Menu 2
Exit Sub
Else
Text.Text = ""
CommD.filename = ""
Changed = False
SP5.Caption = "SJOTS Perfect 5"
Menu 2
End If
End Sub

Private Sub Copy_Click()
Clipboard.Clear
Clipboard.SetText Text.SelText, vbCFRTF
End Sub

Private Sub Cut_Click()
Clipboard.SetText Text.SelText, vbCFRTF
Text.SelText = ""
End Sub

Private Sub FColor_Click()
With CommD
.CancelError = True
On Error GoTo NoFC
.Flags = &H0 + &H1
.ShowColor
Kleur = .Color
End With
Text.SelColor = Kleur

NoFC:
Exit Sub
End Sub

Private Sub CFont_Click()
With CommD
.CancelError = True
On Error GoTo GeenFont
.Flags = &H0 + &H1
.ShowFont
End With
With Text
.SelFontName = CommD.FontName
.SelFontSize = CommD.FontSize
.SelBold = CommD.FontBold
.SelItalic = CommD.FontItalic
End With

GeenFont:
Exit Sub

End Sub

Private Sub Fonts_Click()
Text.SelFontName = Fonts.Text
End Sub

Private Sub Form_Load()
Menu 2
Dim i
Dim K
    For i = 0 To Screen.FontCount - 1
        Fonts.AddItem Screen.Fonts(i)
    Next i
    
    Size.AddItem "1"
    Size.AddItem "2"
    Size.AddItem "3"
    Size.AddItem "4"
    Size.AddItem "5"
    Size.AddItem "6"
    Size.AddItem "7"
    Size.AddItem "8"
    Size.AddItem "9"
    Size.AddItem "10"
    Size.AddItem "11"
    Size.AddItem "12"
    For K = 14 To 100 Step 2
    Size.AddItem K
    Next K
    SNext.Enabled = False
  'the following code adds bitmaps to the menu
  'very cool solution --> by Swift
  
  Dim m%
  Dim hMenu, hSubMenu, menuID, x
  hMenu = GetMenu(hwnd)
  hSubMenu = GetSubMenu(hMenu, 0) '1 for "Next" Menu, etc.
  menuID = GetMenuItemID(hSubMenu, 0)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(1).Picture, Img_Bar.ListImages(1).Picture)
  menuID = GetMenuItemID(hSubMenu, 2)
  x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(2).Picture, Img_Bar.ListImages(2).Picture)
  menuID = GetMenuItemID(hSubMenu, 5)
    x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(3).Picture, Img_Bar.ListImages(3).Picture)
    menuID = GetMenuItemID(hSubMenu, 8)
    x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(4).Picture, Img_Bar.ListImages(4).Picture)
      
hSubMenu = GetSubMenu(hMenu, 1) '1 for "Next" Menu, etc.
menuID = GetMenuItemID(hSubMenu, 0)
x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(5).Picture, Img_Bar.ListImages(5).Picture)
menuID = GetMenuItemID(hSubMenu, 1)
x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(6).Picture, Img_Bar.ListImages(6).Picture)
menuID = GetMenuItemID(hSubMenu, 2)
x = SetMenuItemBitmaps(hMenu, menuID, &H4, Img_Bar.ListImages(7).Picture, Img_Bar.ListImages(7).Picture)
End Sub

Private Sub Form_Resize()
Text.Height = SP5.ScaleHeight - 255
Text.Width = SP5.ScaleWidth
End Sub

Private Sub New_Click()
InitNew
End Sub

Private Sub Next_Click()
SearchFor_Click
End Sub

Private Sub Open_Click()
With CommD
.CancelError = True
On Error GoTo Open_ET
.DialogTitle = "Open a file..."
.Filter = Filter
.ShowOpen
FName = .filename
Text.LoadFile (FName)
SP5.Caption = "SJOTS Perfect 5 - [" & FName & "]"
Menu 1
Changed = False 'otherwise the program would think the
                'contents had changed. Ofcourse they are
                'changed, but I only want them marked as
                'changed if you type in something.
Exit Sub

Open_ET:
SP5.Caption = "SJOTS Perfect 5"
Exit Sub
End With
End Sub

Private Sub Paste_Click()
Text.Text = Text.Text & Clipboard.GetText(vbCFRTF)
End Sub

Private Sub PrintDoc_Click()
PrintRTF Text, 1440, 1440, 1440, 1440
End Sub

Private Sub Quit_Click()
DetectChange
Unload Me
End Sub

Private Sub Save_Click()
If CommD.filename = "" Then
SaveAs_Click
Else
FName = CommD.filename
Text.SaveFile (FName)
Changed = False
End If
End Sub

Private Sub SaveAs_Click()
With CommD
.CancelError = True
On Error GoTo Save_ET
.DialogTitle = "Save a file as..."
.Filter = "Rich Text Files (*.RTF)|*.RTF"
.ShowSave
FName = .filename
End With
Text.SaveFile (FName)
Changed = False
SP5.Caption = "SJOTS Perfect 5 - [" & FName & "]"
Exit Sub

Save_ET:
Exit Sub
End Sub

Private Sub SearchFor_Click()
    Static Where As Integer
    Dim SearchFor As String
    FNEXT = 1
    SearchFor = InputBox("Type in what word or phrase you are looking for...", "SJOTS Perfect 5")
    Where = InStr(FNEXT, Text.Text, SearchFor) ' Find string in text.
    If Where Then   ' If found,
        Text.SelStart = Where - 1  ' set selection start and
        Text.SelLength = Len(SearchFor)   ' set selection length.
        FNEXT = Val(Where + Text.SelLength)
        Found = SearchFor
        SNext.Enabled = True
Else
        NotFound = MsgBox("'" & SearchFor & "'" & " wasn't found.", vbCritical + vbOKOnly, "SJOTS Perfect 5")
        SNext.Enabled = False
    End If

End Sub

Private Sub Size_Click()
Text.SelFontSize = Size.Text
End Sub

Private Sub SNext_Click()
FindNext
End Sub

Private Sub Text_Change()
Changed = True
End Sub

Function DetectChange()
If Changed Then
   Ask = MsgBox("Do you want to save the changes made in " & Text.filename & " ?", vbYesNoCancel + vbInformation, "SJOTS Perfect 5")
  Select Case Ask
   Case vbCancel 'this option is nessecary in case you don't
   Exit Function 'want to start a new document or quit the program.
   
   Case vbNo     'Don't save any changes and set the changed
   Changed = False ' boolean to false.
   Exit Function
   
   Case vbYes    'Start the save as dialog and reset the
   SaveAs_Click  'changed boolean to false.
   Changed = False
   Exit Function
  End Select
End If
End Function

Function InitNew()
If Changed Then
   Ask = MsgBox("Do you want to save the changes made in " & Text.filename & " ?", vbYesNoCancel + vbInformation, "SJOTS Perfect 5")
  Select Case Ask
   Case vbCancel
   Exit Function
   
   Case vbNo
   Text.Text = ""
   CommD.filename = ""
   SP5.Caption = "SJOTS Perfect 5"
   Menu 1
   Changed = False
   Exit Function
   
   Case vbYes
   SaveAs_Click
   Text.Text = ""
   CommD.filename = ""
   SP5.Caption = "SJOTS Perfect 5"
   Menu 1
   Changed = False
   Exit Function
  End Select
Else
   Text.Text = ""
   CommD.filename = ""
   SP5.Caption = "SJOTS Perfect 5"
   Menu 1
End If
End Function

Function Menu(Choice As Integer)
Select Case Choice
Case 1
Text.Enabled = True
Save.Enabled = True
SaveAs.Enabled = True
CloseDoc.Enabled = True
PrintDoc.Enabled = True
Copy.Enabled = True
Cut.Enabled = True
Paste.Enabled = True
Fonts.Enabled = True
Size.Enabled = True
SearchFor.Enabled = True
CFont.Enabled = True
FColor.Enabled = True
For i = 3 To 13
Toolbar1.Buttons(i).Enabled = True
Next i
Exit Function

Case 2
Text.Enabled = False
Save.Enabled = False
SaveAs.Enabled = False
CloseDoc.Enabled = False
PrintDoc.Enabled = False
Copy.Enabled = False
Cut.Enabled = False
Paste.Enabled = False
Fonts.Enabled = False
Size.Enabled = False
SearchFor.Enabled = False
SNext.Enabled = False
CFont.Enabled = False
FColor.Enabled = False
For i = 3 To 13
Toolbar1.Buttons(i).Enabled = False
Next i
End Select
End Function

Private Sub Timer1_Timer()
If Text.SelBold = True Then
Toolbar1.Buttons("Bold").Value = tbrPressed
Else
Toolbar1.Buttons("Bold").Value = tbrUnpressed
End If
If Text.SelItalic Then
Toolbar1.Buttons("Italic").Value = tbrPressed
Else
Toolbar1.Buttons("Italic").Value = tbrUnpressed
End If
If Text.SelUnderline Then
Toolbar1.Buttons("Underline").Value = tbrPressed
Else
Toolbar1.Buttons("Underline").Value = tbrUnpressed
End If
On Error GoTo NoText
Fonts.Text = Text.SelFontName
Size.Text = Text.SelFontSize

NoText:
Exit Sub

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "New"
New_Click
Case "Open"
Open_Click
Case "Save"
SaveAs_Click
Case "Print"
PrintDoc_Click
Case "Copy"
Copy_Click
Case "Cut"
Cut_Click
Case "Paste"
Paste_Click
Case "Bold"
If Button.Value = tbrPressed Then
Text.SelBold = True
Else
Text.SelBold = False
End If
Case "Italic"
If Button.Value = tbrPressed Then
Text.SelItalic = True
Else
Text.SelItalic = False
End If
Case "Underline"
If Button.Value = tbrPressed Then
Text.SelUnderline = True
Else
Text.SelUnderline = False
End If
End Select
End Sub

Function FindNext()
Static Where As Integer
Dim SearchFor As String
    SearchFor = Found
    Where = InStr(FNEXT, Text.Text, SearchFor) ' Find string in text.
    If Where Then   ' If found,
        Text.SelStart = Where - 1  ' set selection start and
        Text.SelLength = Len(SearchFor)   ' set selection length.
        FNEXT = Val(Where + Text.SelLength)
        Found = SearchFor
        SNext.Enabled = True
Else
        NotFound = MsgBox("'" & SearchFor & "'" & " wasn't found.", vbCritical + vbOKOnly, "SJOTS Perfect 5")
        SNext.Enabled = False
    End If

End Function
