VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form Form5 
   Caption         =   "Confirmation"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form5"
   ScaleHeight     =   7125
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   20
      Top             =   9720
      Width           =   3975
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   19
      Top             =   8880
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   15
      Top             =   6360
      Width           =   1215
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   4
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   2
      Top             =   10560
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   11160
      ScaleHeight     =   5595
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   1800
      Width           =   4935
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   855
      Left            =   11160
      TabIndex        =   0
      Top             =   10440
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
      Caption         =   "Confirm Release"
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
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uploaded Picture"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   11160
      TabIndex        =   21
      Top             =   7560
      Width           =   4935
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
      Left            =   240
      TabIndex        =   18
      Top             =   8880
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
      Left            =   240
      TabIndex        =   17
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   720
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmation"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   120
      Width           =   7575
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
      Left            =   240
      TabIndex        =   13
      Top             =   1800
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
      Left            =   240
      TabIndex        =   12
      Top             =   2640
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
      Left            =   240
      TabIndex        =   11
      Top             =   5520
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
      Left            =   240
      TabIndex        =   10
      Top             =   6360
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
      Left            =   240
      TabIndex        =   9
      Top             =   7200
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
      Left            =   240
      TabIndex        =   8
      Top             =   8040
      Width           =   2175
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
      Left            =   240
      TabIndex        =   7
      Top             =   10560
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    MDIForm1.Visible = False
    
    With Rs
    Text1 = .Fields(1)
    Text2 = .Fields(2)
    Text3 = .Fields(3)
    Text4 = .Fields(8)
    Text5 = .Fields(10)
    Text6 = .Fields(11)
    Text7 = .Fields(9)
    Text8 = .Fields(5)
    Text9 = .Fields(6)
    Me.Picture1.Picture = LoadPicture(.Fields(12))
    End With
    
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

Private Sub Form_Load()
    Me.Image1.Picture = LoadPicture(App.Path & "/Green.jpg")
    Me.WindowState = vbMaximized
    
    Me.Label1.ZOrder 0
    Me.Label10.ZOrder 0
    Me.Label2.ZOrder 0
    Me.Label3.ZOrder 0
    Me.Label4.ZOrder 0
    Me.Label5.ZOrder 0
    Me.Label6.ZOrder 0
    Me.Label7.ZOrder 0
    Me.Label8.ZOrder 0
    Me.Label9.ZOrder 0
    Me.Label11.ZOrder 0
    
    Me.Text1.Enabled = False
    Me.Text2.Enabled = False
    Me.Text3.Enabled = False
    Me.Text4.Enabled = False
    Me.Text5.Enabled = False
    Me.Text6.Enabled = False
    Me.Text7.Enabled = False
    Me.Text8.Enabled = False
    Me.Text9.Enabled = False
    
    
    
End Sub

Private Sub Form_Resize()
    With Me.Image1
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Rs.State = 1 Then 'if recordset is open
        Rs.Close
    End If
    MDIForm1.Visible = True
End Sub

Private Sub UserButtonz1_Click()
   Rs.Fields(7) = "Yes"
   Rs.Update
   Rs.Close
   Unload Me
End Sub
