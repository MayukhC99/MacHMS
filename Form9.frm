VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prescription"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9300
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
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
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
      Width           =   4215
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
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   4215
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   735
      Left            =   6720
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEXT"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16711935
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16777088
      ColorButtonUp   =   16761087
      ColorButtonDown =   16777088
      BorderBrightness=   0
      ColorBright     =   16761087
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Prescription"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bed No. :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "U_Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    MDIForm1.Visible = False
    Me.Text1.SetFocus
End Sub

Private Sub Form_Load()
    Me.Image1.Picture = LoadPicture(App.Path & "/purple.jpg")
    
    Me.Label1.ZOrder 0
    Me.Label2.ZOrder 0
    Me.Label3.ZOrder 0
    Me.Label4.ZOrder 0
    
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
    MDIForm1.Visible = True
End Sub

Private Sub UserButtonz1_Click()
'On Error GoTo invalid
Dim take As String
    Call ConnectDB
    
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from Patients where PID=" & Val(Trim(Text1)) & " and Bed_Number='" & Trim(Text2) & "' and U_Name='" & Trim(Text3) & "' and Release='No'", Cn, adOpenDynamic, adLockOptimistic
    
    If Rs.RecordCount < 1 Then
        Rs.Close
        Cn.Close
        MsgBox "No Patient Record Found !", vbInformation + vbOKOnly, "Patient NOT Found"
        Unload Me
        Exit Sub
    End If
    
    take = Rs.Fields(13)
    f = CStr(App.Path) & "\Noprescription.pdf"
    If take = f Then
        Form8.UserButtonz1.Caption = "Upload Prescription"
    Else
        Form8.UserButtonz1.Caption = "View Prescription"
    End If
    
    Me.Visible = False
    Form8.Visible = True
    Unload Me
    
    Exit Sub
'invalid: MsgBox "Invalid Input", vbCritical + vbOKOnly, "Error"
'        Rs.Close
    
End Sub
