VERSION 5.00
Object = "{69C832A0-68F4-452F-9B16-837E157288D9}#1.0#0"; "styler_button.ocx"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Registration Form"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10410
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   10335
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3240
      TabIndex        =   21
      Text            =   "Group"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      TabIndex        =   20
      Text            =   "Doctor's Name"
      Top             =   10080
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      TabIndex        =   19
      Text            =   "Treatment Department"
      Top             =   9240
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      TabIndex        =   18
      Text            =   "Select Your Gender"
      Top             =   6720
      Width           =   3975
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   855
      Left            =   11760
      TabIndex        =   17
      Top             =   10800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Register"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   49152
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   65280
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648384
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Styler_button.StylerButton StylerButton1 
      Height          =   735
      Left            =   11760
      TabIndex        =   16
      Top             =   8040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      Caption         =   "UPLOAD PICTURE"
      ForeColor       =   16711680
      CaptionDisableColor=   13153946
      CaptionEffectColor=   49152
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   11760
      ScaleHeight     =   5595
      ScaleWidth      =   4875
      TabIndex        =   15
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   10920
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Bed No."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   10920
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Registration"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Dim filename As String


Private Sub Combo2_Click()
    s = Combo2.Text
    
    Combo3.Clear
    Call ConnectDB
    rs3.CursorLocation = adUseClient
    rs3.Open "Select * from Doctors where Department='" & s & "'", Cn, adOpenDynamic, adLockOptimistic
    'rs3.Open "Select * from Doctors where Username='" & UID & "' and  Department='" & s & "'", Cn, adOpenDynamic, adLockOptimistic
    
    While rs3.EOF = False
        Me.Combo3.AddItem rs3.Fields(1)
        rs3.MoveNext
    Wend
    rs3.Close
End Sub

Private Sub Form_Activate()
    MDIForm1.Visible = False
    
    Call ConnectDB
    
    
    Rs2.CursorLocation = adUseClient
    Rs2.Open "Select * from Departments", Cn, adOpenDynamic, adLockOptimistic
    
    While Rs2.EOF = False
        Me.Combo2.AddItem Rs2.Fields(0)
        Rs2.MoveNext
    Wend
    Rs2.Close
    
    Me.Text1.SetFocus
    
    filename = CStr(App.Path) & "/Noprofilepic.jpg"
    
End Sub

Private Sub Form_Load()
    
    Me.WindowState = vbMaximized
    Me.Image1.Picture = LoadPicture(App.Path & "/Green.jpg")
    Me.Label1.ZOrder 0
    Me.Label2.ZOrder 0
    Me.Label3.ZOrder 0
    
    Combo1.AddItem "Male"
    Combo1.AddItem "Female"
    Combo1.AddItem "Something Else"
    
    With Me.Combo4
        .AddItem "O+"
        .AddItem "AB+"
        .AddItem "A+"
        .AddItem "B+"
        .AddItem "O-"
        .AddItem "AB-"
        .AddItem "A-"
        .AddItem "B-"
    End With
        
        
End Sub

Private Sub Form_Resize()
    Me.Image1.Height = Me.ScaleHeight
    Me.Image1.Width = Me.ScaleWidth
    Me.Image1.Top = 0
    Me.Image1.Left = 0
    
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
    If Not Rs.State <> 1 Then 'rs is open
        Rs.Close
    End If
    MDIForm1.Visible = True
End Sub

Private Sub StylerButton1_Click()
    Me.CommonDialog1.ShowOpen
    Me.CommonDialog1.Filter = "Jpeg|*.jpg"
    filename = Me.CommonDialog1.filename
    
    If InStr(1, filename, ".jpg") = 0 Then
        MsgBox "Only JPG files are allowed as Profile Image", vbCritical + vbOKOnly, "Invalid"
        filename = CStr(App.Path) & "/Noprofilepic.jpg"
        Exit Sub
    End If
    
    Me.Picture1.Picture = LoadPicture(filename)
    
    'resizing the picture so that it fits into the picturebox

    Picture1.ScaleMode = 3

    Picture1.AutoRedraw = True

    Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, _
    Picture1.Picture.Width / 26.46, _
    Picture1.Picture.Height / 26.46

    Picture1.Picture = Picture1.Image
    
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'Accepts only numeric input
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
      KeyAscii = 0
      Beep
    End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    'Accepts only numeric input
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
      KeyAscii = 0
      Beep
    End Select
End Sub

Private Sub UserButtonz1_Click()
Dim a As Integer
a = Checkit
    If a = 1 Then
        'database entry
        Call ConnectDB
        Rs.CursorLocation = adUseClient
        Rs.Open "Select * from Patients", Cn, adOpenDynamic, adLockOptimistic
        
        With Rs
            .AddNew
            .Fields(1) = Trim(Text1)
            .Fields(2) = Text2
            .Fields(3) = Text3
            .Fields(4) = UID
            .Fields(5) = Combo3.Text
            .Fields(6) = Trim(Text9)
            .Fields(7) = "No"
            .Fields(8) = Combo1.Text
            .Fields(9) = Combo2.Text
            .Fields(10) = Combo4.Text
            .Fields(11) = Trim(Text6)
            .Fields(12) = filename
            .Fields(13) = CStr(App.Path) & "\Noprescription.pdf"
            .Update
            .MoveLast
            MsgBox "Patient has been Successfully Registered with PID: " & .Fields(0), vbInformation + vbOKOnly, "SUCCESS"
            .Close
        End With
    ElseIf a = 2 Then
        Exit Sub
    End If
    
    Unload Me

End Sub

Function Checkit() As Integer
    If Trim(Text1) = "" Or Trim(Text2) = "" Or Trim(Text3) = "" Or Combo1.Text = "Select Your Gender" Or Combo4.Text = "Group" Or Trim(Text6) = "" Or Trim(Text9) = "" Or Combo2.Text = "Treatment Department" Or Combo3.Text = "Doctor's Name" Then
        MsgBox "Required Fields are empty !", vbCritical + vbOKOnly, "Error"
        Checkit = 2
        Exit Function
    End If
    
    Call ConnectDB
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from Patients where Bed_Number='" & Trim(Text9) & "' and Release='No'", Cn, adOpenDynamic, adLockOptimistic
    
    If Rs.RecordCount <> 0 Then
        MsgBox "This Bed already Booked !", vbCritical + vbOKOnly, "Bed Already Booked"
        Rs.Close
        Text9 = ""
        Checkit = 2
        Exit Function
    End If
    
    Rs.Close
    Checkit = 1
        
End Function

