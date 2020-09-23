VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Easy Note-Pad"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2430
      Left            =   0
      TabIndex        =   2
      Top             =   2625
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   4286
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create a new file."
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open existing file."
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save current file"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print current file."
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut selected text."
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy to clipboard."
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste from clipboard."
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete selected text."
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold font."
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic font."
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline font."
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find text in document."
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Description     =   "Align text left."
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Align text center."
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Description     =   "Align text right."
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6165
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6853
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "9/4/99"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "6:24 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   1575
      Top             =   3885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1575
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":050F
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0621
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0733
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0845
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0957
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A69
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B7B
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C8D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D9F
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EB1
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FC3
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10D5
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E7
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12F9
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140B
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151D
            Key             =   "Justify"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFind 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find Next"
         Height          =   330
         Left            =   4635
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   330
         Left            =   3735
         TabIndex        =   5
         Top             =   225
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   45
         TabIndex        =   4
         Top             =   225
         Width           =   3570
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditTimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFonts 
      Caption         =   "F&onts"
      Begin VB.Menu mnuFontsBold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuFontsItalic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuFontsUnderline 
         Caption         =   "&Underline"
      End
      Begin VB.Menu mnuFontsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontsFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuFontsColor 
         Caption         =   "&Color..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DocChanged As Boolean
Dim docname As String

Private Sub cmdFind_Click()
Dim textfound As Integer

    ' Enables the cmdFindNext
cmdFindNext.Enabled = True

    ' Finds the text in the search box and highlights it,
    ' then sets the focus on the richtextbox so the selected
    ' text is editable.
RichTextBox1.Find (Text1.Text)
RichTextBox1.SetFocus

    ' The richtextbox1.find method returns an integer
    ' value of -1 if the searched for text is not found.
    ' If this is true then it displays a message box.
textfound = RichTextBox1.Find(Text1.Text)
If textfound = -1 Then
MsgBox "End of Document" & vbCr & "Text Not Found", vbInformation, _
    App.Title
End If

End Sub

Private Sub cmdFindNext_Click()

mnuEditFindNext_Click

End Sub

Private Sub Form_Activate()

    ' Updates toolbar and menus.
ChangeToolBar
ChangeMenus

End Sub

Private Sub Form_Load()
Dim LineWidth As Long

    ' Updates toolbar and menus.
ChangeToolBar
ChangeMenus



    ' Tell the RichTextBox to base it's display off of the printer,
    ' and gives a 1 inch border all around.
LineWidth = WYSIWYG_RTF(RichTextBox1, 1440, 1440)

    ' Set the form width to match the line width of printer.
    ' Plus a little to keep the scroll bar from showin up.
Me.Width = LineWidth + 300
   
    ' Center the form in middle of screen.
Me.Move (Screen.Height - Me.Height) / 2, (Screen.Width - Me.Width) / 2

docname = " (Untitled)"
Me.Caption = App.Title & docname

    ' Disables the Find and FindNext buttons
cmdFind.Enabled = False
cmdFindNext.Enabled = False

End Sub

Private Sub Form_Resize()
    
    ' Position the RTF on form,
    ' taking into consideration the size of Toolbar and Status bar.
RichTextBox1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690

End Sub

Private Sub Form_Unload(cancel As Integer)

    ' Checks to see if document has changed since last save,
    ' and if it has gives a message box with a chance to save.
If DocChanged Then
    
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmMain.Caption)
    
    Case vbYes
        mnuFileSave_Click
    Case vbNo
        Unload frmMain
    Case vbCancel
        cancel = True
    
    End Select

End If

End Sub

Private Sub mnuEditDelete_Click()

DeleteSelectedText

    'Updates toolbar and menus.
ChangeToolBar
ChangeMenus

End Sub

Private Sub mnuEditFind_Click()

    ' This routine shows the Find Toolbar.
    ' And resizes the richtextbox depending on whether or not the
    ' standard toolbar is shown.
If tbFind.Visible = False Then
    tbFind.Visible = True
        If tbToolBar.Visible = False Then
            RichTextBox1.Move 0, 630, Me.ScaleWidth, Me.ScaleHeight - 900
        Else
            RichTextBox1.Move 0, 1050, Me.ScaleWidth, Me.ScaleHeight - 1320
        End If
Else
    ' This routine hides the Find Toolbar.
    ' And resizes the richtextbox dependin on whether or not the
    ' standard toolbar is shown.
    tbFind.Visible = False
        If tbToolBar.Visible = False Then
            RichTextBox1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 270
        Else
            RichTextBox1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690
        End If
End If

End Sub

Private Sub mnuEditFindNext_Click()

    ' Set the focus so the selected text can be directly edited.
RichTextBox1.SetFocus

    ' Finds the next instance of the word, starting from
    ' the selected text.
RichTextBox1.Find (Text1.Text), RichTextBox1.SelStart + 1

End Sub

Private Sub mnuEditSelectAll_Click()

    ' Selects all the text in the current document.
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)

    ' Updates the toolbar and menus.
ChangeToolBar
ChangeMenus

End Sub

Private Sub mnuEditTimeDate_Click()
Dim Text As String
Dim SelStart As Long

    ' Deletes the selected text, if any, and gets ready to insert
    ' the time and date string.
DeleteSelectedText

If RichTextBox1.SelLength > 0 Then
End If

    ' Inserts time and date.
Text = RichTextBox1.Text
SelStart = RichTextBox1.SelStart
RichTextBox1.Text = Left(Text, SelStart) & Now & _
        Right(Text, Len(Text) - SelStart)

    ' Resets cursor to original position.
RichTextBox1.SelStart = SelStart

    ' Updates toolbar and menus.
ChangeToolBar
ChangeMenus

End Sub

Private Sub mnuFontsBold_Click()
    
    ' Sets the Bold property and makes the menu context sensative.
If RichTextBox1.SelBold Then
    RichTextBox1.SelBold = False
    mnuFontsBold.Checked = False
Else
    RichTextBox1.SelBold = True
    mnuFontsBold.Checked = True
End If

End Sub

Private Sub mnuFontsColor_Click()

    ' Shows the color dialogue box and sets the current color.
CDL1.Flags = cdlCCFullOpen
CDL1.ShowColor

RichTextBox1.SelColor = CDL1.Color

End Sub

Private Sub mnuFontsFont_Click()

    ' Shows the Font dialogue box and sets the current font.
CDL1.Flags = cdlCFBoth Or cdlCFEffects
CDL1.ShowFont

With RichTextBox1
    .SelFontName = CDL1.FontName
    .SelFontSize = CDL1.FontSize
    .SelBold = CDL1.FontBold
    .SelItalic = CDL1.FontItalic
    .SelStrikeThru = CDL1.FontStrikethru
    .SelUnderline = CDL1.FontUnderline
    .SelColor = CDL1.Color
End With

End Sub

Private Sub mnuFontsItalic_Click()

    ' Sets the italic property and makes the menu context sensative.
If RichTextBox1.SelItalic Then
    RichTextBox1.SelItalic = False
    mnuFontsItalic.Checked = False
Else
    RichTextBox1.SelItalic = True
    mnuFontsItalic.Checked = True
End If
            
End Sub

Private Sub mnuFontsUnderline_Click()

    ' Sets the underline property and makes the menu context sensative.
If RichTextBox1.SelUnderline Then
    RichTextBox1.SelUnderline = False
    mnuFontsUnderline.Checked = False
Else
    RichTextBox1.SelUnderline = True
    mnuFontsUnderline.Checked = True
End If

End Sub

Private Sub RichTextBox1_Change()

    ' Changes the docchanged value to true for saving purposes.
DocChanged = True

    ' Updates the toolbar and menus.
ChangeToolBar
ChangeMenus

End Sub

Private Sub RichTextBox1_Click()

    ' Updates the toolbar and menus just in case.
ChangeToolBar
ChangeMenus

End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)

 ' Updates the toolbar and menus just in case.
ChangeMenus
ChangeToolBar

End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

' Updates the toolbar and menus just in case.
ChangeMenus
ChangeToolBar

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error Resume Next
    ' Just tells the toolbar buttons what to do...
    Select Case Button.Key
        
        Case "New"
            mnuFileNew_Click
        
        Case "Open"
            mnuFileOpen_Click
        
        Case "Save"
            mnuFileSave_Click
        
        Case "Print"
            mnuFilePrint_Click
        
        Case "Cut"
            mnuEditCut_Click
        
        Case "Copy"
            mnuEditCopy_Click
        
        Case "Paste"
            mnuEditPaste_Click
        
        Case "Delete"
            mnuEditDelete_Click
        
        Case "Bold"
            
            mnuFontsBold_Click
        
        Case "Italic"
            
            mnuFontsItalic_Click
                        
        Case "Underline"
            
            mnuFontsUnderline_Click
            
        Case "Find"
            
            mnuEditFind_Click
                    
        Case "Align Left"
            
            ' Aligns text to left margin.
            RichTextBox1.SelAlignment = rtfLeft
        
        Case "Center"
            
            'Aligns text to center.
            RichTextBox1.SelAlignment = rtfCenter
        
        Case "Align Right"
            
            'Aligns text to right margin.
            RichTextBox1.SelAlignment = rtfRight
End Select

End Sub

Private Sub mnuHelpAbout_Click()
    
    ' Simple message box for the About.
MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub mnuViewStatusBar_Click()
    
    ' Shows or hides the status bar as needed.
    ' And makes the menu context sensative.
mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
sbStatusBar.Visible = mnuViewStatusBar.Checked

End Sub

Private Sub mnuViewToolbar_Click()
    
    ' Shows or hides the toolbar as needed.
    ' And makes the menu context sensative.
mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
tbToolBar.Visible = mnuViewToolbar.Checked

    ' This resizes the richtextbox depending on the state of the toolbar.
If tbToolBar.Visible = False Then
    RichTextBox1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 270
Else
    RichTextBox1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690
End If

End Sub

Private Sub mnuEditPaste_Click()
    
Dim Text As String
Dim ClipboardText As String
Dim SelStart As Long
    
If Clipboard.GetFormat(vbCFText) Then
    
    ' Replace selected text. (if any)
    If RichTextBox1.SelLength > 0 Then
        DeleteSelectedText
    End If
    
    ' Move text we need to a variable.
    Text = RichTextBox1.Text
    SelStart = RichTextBox1.SelStart
    ClipboardText = Clipboard.GetText
    
    ' Gather new text string and replace text box
    ' contents with it.
    RichTextBox1.Text = Left(Text, SelStart) & _
            ClipboardText & Right(Text, Len(Text) - SelStart)
    
    ' Restore the cursor position.
    RichTextBox1.SelStart = SelStart

Else
    
    ChangeMenus
    ChangeToolBar
End If

End Sub

Private Sub mnuEditCopy_Click()
    
    CopytoClipBoard
    ChangeMenus
    ChangeToolBar
    
End Sub

Private Sub mnuEditCut_Click()
    
    CopytoClipBoard
    DeleteSelectedText
    ChangeMenus
    ChangeToolBar
    
End Sub

Private Sub mnuFileExit_Click()
    
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
Dim bcancel As Boolean
Dim ncopy As Integer
On Error GoTo errorhandler

bcancel = False

CDL1.Flags = cdlPDHidePrintToFile Or _
        cdlPDNoSelection Or cdlPDNoPageNums _
        Or cdlPDCollate

CDL1.CancelError = True
CDL1.PrinterDefault = True
CDL1.Copies = 1
CDL1.ShowPrinter

If bcancel = False Then
    PrintRTF RichTextBox1, 1440, 1440, 1440, 1440
    For ncopy = 1 To CDL1.Copies
    Next ncopy
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
bcancel = True
Resume Next
End If

End Sub

Private Sub mnuFileSaveAs_Click()
Dim cancel As Boolean
On Error GoTo errorhandler
cancel = False

CDL1.DefaultExt = ".txt"
CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
CDL1.CancelError = True
CDL1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

CDL1.ShowSave

If Not cancel Then
    If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
        RichTextBox1.SaveFile CDL1.FileName, rtfRTF
    Else
        RichTextBox1.SaveFile CDL1.FileName, rtfText
    End If
    RichTextBox1.FileName = CDL1.FileName
    docname = CDL1.FileName
    Me.Caption = App.Title & " " & docname
    DocChanged = False
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
    cancel = True
    Resume Next
End If

End Sub

Private Sub mnuFileSave_Click()

If docname = " (Untitled)" Then
    mnuFileSaveAs_Click
Else
    If UCase(Right(RichTextBox1.FileName, 3)) = "RTF" Then
        RichTextBox1.SaveFile RichTextBox1.FileName, rtfRTF
    Else
        RichTextBox1.SaveFile RichTextBox1.FileName, rtfText
    End If
    DocChanged = False
End If

End Sub

Private Sub mnuFileOpen_Click()
Dim cancel As Boolean
On Error GoTo errorhandler
cancel = False

CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files|*.*"
CDL1.CancelError = True
CDL1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
CDL1.ShowOpen

If Not cancel Then
    If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
        RichTextBox1.LoadFile CDL1.FileName, rtfRTF
    Else
        RichTextBox1.LoadFile CDL1.FileName, rtfText
    End If
        RichTextBox1.FileName = CDL1.FileName
        docname = RichTextBox1.FileName
        Me.Caption = App.Title & " " & docname
        DocChanged = False
End If
Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
    cancel = True
    Resume Next
End If
End

End Sub

Private Sub mnuFileNew_Click()
Dim cancel As Integer

If DocChanged = False Then
    RichTextBox1.Text = ""
Else
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmMain.Caption)
    
    Case vbYes
        mnuFileSave_Click
    Case vbNo
        RichTextBox1.Text = ""
    Case vbCancel
        cancel = True
    
    End Select
End If

End Sub


Public Sub ChangeMenus()

    ' Makes the menus and toolbar context sensative.
mnuFileSave.Enabled = DocChanged
mnuFileSaveAs.Enabled = DocChanged
mnuEditCopy.Enabled = False
mnuEditCut.Enabled = False
mnuEditDelete.Enabled = False
mnuEditPaste.Enabled = False
tbToolBar.Buttons("Save").Enabled = DocChanged
    
If RichTextBox1.SelLength > 0 Then
    mnuEditCut.Enabled = True
    mnuEditCopy.Enabled = True
    mnuEditDelete.Enabled = True
    tbToolBar.Buttons("Cut").Enabled = True
    tbToolBar.Buttons("Copy").Enabled = True
    tbToolBar.Buttons("Delete").Enabled = True
Else
    mnuEditCut.Enabled = False
    mnuEditCopy.Enabled = False
    mnuEditDelete.Enabled = False
    tbToolBar.Buttons("Cut").Enabled = False
    tbToolBar.Buttons("Copy").Enabled = False
    tbToolBar.Buttons("Delete").Enabled = False
    
End If

If Clipboard.GetFormat(vbCFText) Then
    mnuEditPaste.Enabled = True
    tbToolBar.Buttons("Paste").Enabled = True
Else
    mnuEditPaste.Enabled = False
    tbToolBar.Buttons("Paste").Enabled = True
End If

If RichTextBox1.SelBold Then
    mnuFontsBold.Checked = True
Else
    mnuFontsBold.Checked = False
End If

If RichTextBox1.SelItalic Then
    mnuFontsItalic.Checked = True
Else
    mnuFontsItalic.Checked = False
End If

If RichTextBox1.SelUnderline Then
    mnuFontsUnderline.Checked = True
Else
    mnuFontsUnderline.Checked = False
End If

End Sub

Private Sub CopytoClipBoard()

    ' Copies selected text to clipboard.
Clipboard.SetText RichTextBox1.SelText

End Sub

Public Sub DeleteSelectedText()

    ' Sets the value of the selected text to nothing.
RichTextBox1.SelText = ""

End Sub

Public Sub ChangeToolBar()

    ' Makes portions of the tool bar context sensative.
If RichTextBox1.SelBold Then
    tbToolBar.Buttons("Bold").Value = tbrPressed
Else
    tbToolBar.Buttons("Bold").Value = tbrUnpressed
End If

If RichTextBox1.SelItalic Then
    tbToolBar.Buttons("Italic").Value = tbrPressed
Else
    tbToolBar.Buttons("Italic").Value = tbrUnpressed
End If

If RichTextBox1.SelUnderline Then
    tbToolBar.Buttons("Underline").Value = tbrPressed
Else
    tbToolBar.Buttons("Underline").Value = tbrUnpressed
End If

End Sub

Private Sub Text1_Change()

    ' Enables the find button when the search text is entered.
cmdFind.Enabled = True

End Sub
