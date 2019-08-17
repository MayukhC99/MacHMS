VERSION 5.00
Object = "{69C832A0-68F4-452F-9B16-837E157288D9}#1.0#0"; "styler_button.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   882
      TabMaxWidth     =   3528
      BackColor       =   16777215
      ForeColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sign up"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text5"
      Tab(0).Control(1)=   "StylerButton1"
      Tab(0).Control(2)=   "Text4"
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(6)=   "Image4"
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(8)=   "Image1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Log in"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "StylerButton2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   9
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   -74760
         TabIndex        =   6
         Top             =   4440
         Width           =   3255
      End
      Begin Styler_button.StylerButton StylerButton1 
         Height          =   495
         Left            =   -74400
         TabIndex        =   7
         Top             =   5520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Sign Up !"
         ForeColor       =   33023
         CaptionDisableColor=   12236471
         CaptionEffectColor=   8454016
         IconDisableColor=   12236471
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconTransparentColour=   65280
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   5
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   4
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   3
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   2
         Top             =   1560
         Width           =   3255
      End
      Begin Styler_button.StylerButton StylerButton2 
         Height          =   495
         Left            =   600
         TabIndex        =   10
         Top             =   4320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Login"
         ForeColor       =   33023
         CaptionDisableColor=   12236471
         CaptionEffectColor=   8454016
         IconDisableColor=   12236471
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconTransparentColour=   65280
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   -75000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login !"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Image Image3 
         Height          =   975
         Left            =   -120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4215
      End
      Begin VB.Image Image2 
         Height          =   5535
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sign Up to Get Started !"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   5535
         Left            =   -70920
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim first As Boolean
'Dim h As Integer
'Dim w As Integer
'Dim c As Integer
Dim SRs As New ADODB.Recordset

Private Sub Form_Load()
   ' first = True
   ' c = 1
    Form1.Image1.Picture = LoadPicture(App.Path & "/sign up.jpg")
    Form1.Image2.Picture = LoadPicture(App.Path & "/login.jpg")
    Form1.Image3.Picture = LoadPicture(App.Path & "/white.jpg")
    Form1.Image4.Picture = LoadPicture(App.Path & "/white.jpg")
    Me.SSTab1.Width = Me.ScaleWidth
    Me.SSTab1.Height = Me.ScaleHeight

   
    
    Me.Text1.ToolTipText = "Username"
    Me.Text1.ForeColor = vbGrayText
    Me.Text1 = Me.Text1.ToolTipText
    
    Me.Text2.ToolTipText = "Name"
    Me.Text2.ForeColor = vbGrayText
    Me.Text2 = Me.Text2.ToolTipText
    
    Me.Text3.ToolTipText = "Hospital Name"
    Me.Text3.ForeColor = vbGrayText
    Me.Text3 = Me.Text3.ToolTipText
    
    Me.Text4.ToolTipText = "Designation"
    Me.Text4.ForeColor = vbGrayText
    Me.Text4 = Me.Text4.ToolTipText
    
    Me.Text5.ToolTipText = "Password"
    Me.Text5.ForeColor = vbGrayText
    Me.Text5 = Me.Text5.ToolTipText
    
    Me.Text6.ToolTipText = "Username"
    Me.Text6.ForeColor = vbGrayText
    Me.Text6 = Me.Text6.ToolTipText
    
    Me.Text7.ToolTipText = "Password"
    Me.Text7.ForeColor = vbGrayText
    Me.Text7 = Me.Text7.ToolTipText
    
    
'    GetLocation Me
'    ResizeForm Me
'
'    Me.Label1.Font = SetFontSize()
'    Me.Label2.Font = SetFontSize()
'    Me.Text1.Font = SetFontSize()
'    Me.Text2.Font = SetFontSize()
'    Me.Text3.Font = SetFontSize()
'    Me.Text4.Font = SetFontSize()
'    Me.Text5.Font = SetFontSize()
'    Me.Text6.Font = SetFontSize()
'    Me.Text7.Font = SetFontSize()

'    Me.Text6 = "Gennedy99"
'    Me.Text7 = "1234"
    
End Sub

Private Sub Form_Resize()
'    ResizeControls Me
'
'    Me.Label1.Font = SetFontSize()
'    Me.Label2.Font = SetFontSize()
'    Me.Text1.Font = SetFontSize()
'    Me.Text2.Font = SetFontSize()
'    Me.Text3.Font = SetFontSize()
'    Me.Text4.Font = SetFontSize()
'    Me.Text5.Font = SetFontSize()
'    Me.Text6.Font = SetFontSize()
'    Me.Text7.Font = SetFontSize()
        'sizing all controls in form1
        With Form1.SSTab1
        .Height = Form1.ScaleHeight + 100
        .Width = Form1.ScaleWidth
        .Top = 0
        .Left = 0
        End With

        With Form1.Image1
         '.Top = .Top * X
         .Height = Form1.SSTab1.Height - 450
         .Width = Form1.SSTab1.Width / 1.7
         '.Left = .Left * Y
        End With
        With Form1.Image2
         '.Top = .Top *x
         .Height = Form1.SSTab1.Height - 450
         .Width = Form1.SSTab1.Width / 1.7
         '.Left = .Left * Y
        End With

        'Form1.Image3.Width = (Form1.ScaleWidth - Form1.Image2.Width)
        'Form1.Image3.Height = (Form1.ScaleHeight - Form1.Image2.Height)
        Form1.Image3.Height = Form1.SSTab1.Height
        Form1.Image4.Height = Form1.SSTab1.Height
    'If first = False Then
     ' With Form1
      '      .Text1.Height = .Text1.Height * (Form1.SSTab1.Height / h)
       '     .Text1.Width = .Text1.Width * (Form1.SSTab1.Width / w)
        '    .Text1.Top = .Text1.Top * (Form1.SSTab1.Width / w)
        '    .Text1.Left = .Text1.Left * (Form1.SSTab1.Height / h)
            '.Text1.Height = .Text1.Height * (Form1.SSTab1.Height / h)
            '.Text1.Width = .Text1.Width * (Form1.SSTab1.Width / w)
    '    End With
    'End If
    'If c = 2 Then
    '    first = False
    'Else
    '    c = c + 1
   ' End If
    
   ' h = Form1.SSTab1.Height
   ' w = Form1.SSTab1.Width
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With MDIForm1
        .signup.Enabled = True
        .adminlogin.Enabled = True
        .hospitallogin.Enabled = True
    End With
End Sub

Private Sub StylerButton1_Click()
    On Error GoTo similar
    If errorcheck = True Then
        Exit Sub
    Else
        Call ConnectDB
        
        'SRs.Close 'better to close it
        SRs.Open "Select * from signup", Cn, adOpenDynamic, adLockOptimistic
        With SRs
            .AddNew
            .Fields(0) = Trim(Text1)
            .Fields(1) = Trim(Text2)
            .Fields(2) = Trim(Text3)
            .Fields(3) = Date
            .Fields(4) = Time
            .Fields(5) = Trim(Text4)
            .Fields(6) = Text5
            .Update
            .Close
        End With
        Cn.Close
        
        Call postsignup
        Unload Me
    End If
    Exit Sub
similar:
    MsgBox "The Username already exists", vbInformation + vbOKCancel, "Username !"
    Me.Text1.ToolTipText = "Username"
    Me.Text1.ForeColor = vbGrayText
    Me.Text1 = Me.Text1.ToolTipText
    
    Me.Text2.ToolTipText = "Name"
    Me.Text2.ForeColor = vbGrayText
    Me.Text2 = Me.Text2.ToolTipText
    
    Me.Text3.ToolTipText = "Hospital Name"
    Me.Text3.ForeColor = vbGrayText
    Me.Text3 = Me.Text3.ToolTipText
    
    Me.Text4.ToolTipText = "Designation"
    Me.Text4.ForeColor = vbGrayText
    Me.Text4 = Me.Text4.ToolTipText
    
    Me.Text5.ToolTipText = "Password"
    Me.Text5.ForeColor = vbGrayText
    Me.Text5 = Me.Text5.ToolTipText
    Text5.FontName = "tahoma"
    Text5.FontSize = 12
    Text5.PasswordChar = ""
    
    Exit Sub
End Sub

Private Sub StylerButton2_Click()
    ud = Text6
    psd = Text7
    
    
'    Cm.ActiveConnection = Cn
'    Cm.CommandText = "Select * from Signup"
'    SRs.ActiveConnection = Cn

    Call ConnectDB
    
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from signup where Username='" & Text6 & "' and Password='" & Text7 & "'", Cn, adOpenDynamic, adLockOptimistic 'opening with Rs to make login global
    
'    MsgBox Rs.RecordCount 'the recordset will not count unless cursor locaiton is set to aduseclient
    
    
    If Rs.RecordCount <> 0 Then
        UserName = Rs.Fields(1) 'taking the full name
        UID = ud                    'taking the userID
        MDIForm1.StatusBar1.Panels.Item(1).Text = UserName
        
        SRs.Open "Select * from LoginDetails", Cn, adOpenDynamic, adLockOptimistic

        SRs.AddNew
        SRs!UserName = UID
        SRs!login_date = Date
        SRs!login_time = Time
        SRs!sessionactive = 1
        SRs.Update
        SRs.Close
        
'    Cn.Execute ("INSERT INTO LOGINDETAILS (UserName,Login_Date, Login_Time, SessionActive) VALUES ('" & UID & "'," & Date & ", '" & Time & "',1)")
        
        Call switchon
        Call displayit
        Rs.Close
        Cn.Close
        Unload Me
    Else
        Rs.Close
        MsgBox "Wrong Username or Password !", vbExclamation + vbOKOnly, "Check Again"
        
        Me.Text6.ForeColor = vbGrayText
        Me.Text6 = Me.Text6.ToolTipText
    
        Me.Text7.ForeColor = vbGrayText
        Me.Text7 = Me.Text7.ToolTipText
        Text7.FontName = "tahoma"
        Text7.FontSize = 12
        Text7.PasswordChar = ""
    End If
End Sub

Private Sub Text1_GotFocus()
    With Text1
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .Refresh
        End If
    End With
End Sub

Private Sub Text1_LostFocus()
    With Text1
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .Refresh
        End If
    End With
End Sub

Private Sub Text2_GotFocus()
    With Text2
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .Refresh
        End If
    End With
End Sub

Private Sub Text2_LostFocus()
    With Text2
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .Refresh
        End If
    End With
End Sub

Private Sub Text3_GotFocus()
    With Text3
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .Refresh
        End If
    End With
End Sub

Private Sub Text3_LostFocus()
    With Text3
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .Refresh
        End If
    End With
End Sub

Private Sub Text4_GotFocus()
    With Text4
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .Refresh
        End If
    End With
End Sub

Private Sub Text4_LostFocus()
    With Text4
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .Refresh
        End If
    End With
End Sub

Private Sub Text5_GotFocus() ' the password part
    With Text5
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .FontName = "Wingdings 2"
            .FontSize = 14
            .PasswordChar = Chr(151)
            .Refresh
        End If
    End With
End Sub

Private Sub Text5_LostFocus() ' the password part
    With Text5
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .FontName = "tahoma"
            .FontSize = 12
            .PasswordChar = ""
            .Refresh
        End If
    End With
End Sub


Private Sub Text7_GotFocus() ' the password part
    With Text7
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .FontName = "Wingdings 2"
            .FontSize = 14
            .PasswordChar = Chr(151)
            .Refresh
        End If
    End With
End Sub

Private Sub Text7_LostFocus() ' the password part
    With Text7
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .FontName = "tahoma"
            .FontSize = 12
            .PasswordChar = ""
            .Refresh
        End If
    End With
End Sub

Private Sub Text6_GotFocus()
    With Text6
        If .Text = .ToolTipText Then
            .Text = ""
            .ForeColor = vbBlack
            .Refresh
        End If
    End With
End Sub

Private Sub Text6_LostFocus()
    With Text6
        If Trim(.Text) = "" Then
            .ForeColor = vbGrayText
            .Text = .ToolTipText
            .Refresh
        End If
    End With
End Sub


Function errorcheck() As Boolean
    If Trim(Text1) = "" Or Trim(Text2) = "" Or Trim(Text3) = "" Or Trim(Text4) = "" Or Trim(Text1) = Text1.ToolTipText Or Trim(Text2) = Text2.ToolTipText Or Trim(Text3) = Text3.ToolTipText Or Trim(Text4) = Text4.ToolTipText Or Trim(Text5) = "" Or Trim(Text5) = Text5.ToolTipText Then
        MsgBox "Fill Up all the Fields", vbCritical + vbOKOnly, "Error !"
        With Form1
        .Text1 = .Text1.ToolTipText
        .Text1.ForeColor = vbGrayText
        .Text2 = .Text2.ToolTipText
        .Text2.ForeColor = vbGrayText
        .Text3 = .Text3.ToolTipText
        .Text3.ForeColor = vbGrayText
        .Text4 = .Text4.ToolTipText
        .Text4.ForeColor = vbGrayText
        .Text5 = .Text5.ToolTipText
        .Text5.PasswordChar = ""
        .Text5.FontName = "tahoma"
        .Text5.FontSize = 12
        .Text5.ForeColor = vbGrayText
        End With
        errorcheck = True
        Exit Function
    End If
    Dim s, f As String
    Dim letter As String
    s = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789_"
    f = Trim(Text1)
    For i = 1 To Len(f)
        letter = Mid(f, i, 1) ' you can't use f(i) in vb6 as it does not accepts string as an array
        If InStr(s, letter) = False Then
            MsgBox "Username can only contain normal" & vbCrLf & " characters,underscore and numbers", vbCritical + vbOKOnly, "Error !"
            Form1.Text1 = Form1.Text1.ToolTipText
            Form1.Text1.ForeColor = vbGrayText
            errorcheck = True
            Exit Function
        End If
    Next
    
    If Len(Text5) < 4 Or Mid(Text5, 1, 1) = " " Then
        MsgBox "Password length must be greater than 3 and must not begin with space", vbCritical + vbOKOnly, "Password Error !"
        errorcheck = True
        Exit Function
    End If
    errorcheck = False
End Function


Function postsignup()
    MsgBox "You have successfully Signed Up :)", vbInformation + vbOKOnly, "Success !"
End Function


