VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Patient Information System"
   ClientHeight    =   11460
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17010
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   11535
      Left            =   0
      ScaleHeight     =   11475
      ScaleWidth      =   16950
      TabIndex        =   0
      Top             =   0
      Width           =   17010
      Begin glxpbuttonz.UserButtonz UserButtonz6 
         Height          =   975
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "UserButtonz6"
         IconHighLiteColor=   16744576
         CaptionHighLiteColor=   0
         Picture         =   "MDIForm1.frx":0000
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin glxpbuttonz.UserButtonz UserButtonz1 
         Height          =   1455
         Left            =   5880
         TabIndex        =   3
         Top             =   2160
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Patient Management"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   495
         Left            =   -120
         TabIndex        =   1
         Top             =   6840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   4410
               MinWidth        =   4410
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               TextSave        =   "5/1/2019"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               TextSave        =   "9:26 AM"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               TextSave        =   "NUM"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   1
               Enabled         =   0   'False
               TextSave        =   "CAPS"
            EndProperty
         EndProperty
      End
      Begin glxpbuttonz.UserButtonz UserButtonz2 
         Height          =   1455
         Left            =   5880
         TabIndex        =   4
         Top             =   4080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Doctor Management"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin glxpbuttonz.UserButtonz UserButtonz3 
         Height          =   1455
         Left            =   5880
         TabIndex        =   5
         Top             =   6000
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Billing"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin glxpbuttonz.UserButtonz UserButtonz4 
         Height          =   1455
         Left            =   5880
         TabIndex        =   6
         Top             =   7920
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Booking"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin glxpbuttonz.UserButtonz UserButtonz5 
         Height          =   1455
         Left            =   5880
         TabIndex        =   7
         Top             =   9840
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Charts and Graphs"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Style           =   1
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dashboard"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   7200
         TabIndex        =   2
         Top             =   720
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   120
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   255
      End
   End
   Begin VB.Menu login 
      Caption         =   "Login"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu signup 
         Caption         =   "Signup"
      End
      Begin VB.Menu adminlogin 
         Caption         =   "Admin Login"
      End
      Begin VB.Menu hospitallogin 
         Caption         =   "Hospital Login"
      End
      Begin VB.Menu Logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu patientmanagement 
      Caption         =   "Patient Management"
      Begin VB.Menu details 
         Caption         =   "Details"
      End
      Begin VB.Menu registration 
         Caption         =   "Registration"
      End
      Begin VB.Menu release 
         Caption         =   "Release"
      End
      Begin VB.Menu reports 
         Caption         =   "Reports"
      End
      Begin VB.Menu prescription 
         Caption         =   "Prescription"
      End
   End
   Begin VB.Menu doctormanagement 
      Caption         =   "Doctor Management"
      Begin VB.Menu availability 
         Caption         =   "Availability"
      End
      Begin VB.Menu doctorwisepatient 
         Caption         =   "Doctor wise patient"
      End
      Begin VB.Menu routinechart 
         Caption         =   "Routine Chart"
      End
   End
   Begin VB.Menu billing 
      Caption         =   "Billing"
      Begin VB.Menu Issuebill 
         Caption         =   "Issue Bill"
      End
      Begin VB.Menu totalbillingchart 
         Caption         =   "Total Billing Chart"
      End
   End
   Begin VB.Menu booking 
      Caption         =   "Booking"
      Begin VB.Menu remainingbeds 
         Caption         =   "Remaining Beds"
      End
      Begin VB.Menu totalbookedbeds 
         Caption         =   "Total Booked Beds"
      End
   End
   Begin VB.Menu chartsandgraphs 
      Caption         =   "Charts and Graphs"
      Begin VB.Menu patientbookinggraph 
         Caption         =   "Patient-Booking Graph"
      End
      Begin VB.Menu patientdiseasegraph 
         Caption         =   "Patient-Disease Graph"
      End
      Begin VB.Menu doctorpatientgraph 
         Caption         =   "Doctor-Patient Graph"
      End
      Begin VB.Menu patientdepartmentgraph 
         Caption         =   "Patient-Department Graph"
      End
      Begin VB.Menu doctordepartmentgraph 
         Caption         =   "Doctor-Department Graph"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub availability_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub details_Click()
    'code here
    Form2.Show

'        Form2.height = Me.height / 2
'        Form2.width = Me.width / 2
'        Form2.Top = Me.Top + (Me.height / 3.5)
'        Form2.Left = Me.Left + (Me.width / 4)
End Sub

Private Sub doctordepartmentgraph_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub doctorpatientgraph_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub doctorwisepatient_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub exit_Click()
    ConnectDB
    Cn.Execute ("update LOGINDETAILS set Logout_Date= " & Date & ", Logout_Time='" & Time & "', SessionActive=0 where UserName='" & UID & "' and SessionActive=1 ")
    Set Rs = Nothing
    Set Cn = Nothing
    End
End Sub

Private Sub hospitallogin_Click()
    Form1.Show vbModeless, Me 'to show it always on top of the main form
    Form1.SSTab1.Tab = 1
    
'    Form1.Height = Me.Height / 2
'    Form1.Width = Me.Width / 2
'    Form1.Top = Me.Top + (Me.Height / 3.5)
'    Form1.Left = Me.Left + (Me.Width / 4)
    
End Sub

Private Sub Issuebill_Click()
    Form10.Show
End Sub

Private Sub Logout_Click()
    ConnectDB
    Cn.Execute ("update LOGINDETAILS set Logout_Date= " & Date & ", Logout_Time='" & Time & "', SessionActive=0 where UserName='" & UID & "' and SessionActive=1 ")
    Set Rs = Nothing
    Set Cn = Nothing
    
    Call switchoff
End Sub

Private Sub MDIForm_Load()
    AlterTable = "ALTER TABLE Patients ADD ( BedNo Number(4))"
    CreateTable = "CREATE TABLE BenDetails (BedNO NUMBER(4)PRIMARY KEY, DeptName VARCHAR(20) NOT NULL, Available NUMBER , Remarks VARCHAR(30))"
    
    Me.WindowState = vbMaximized 'maximizing the window
    'Me.Picture1.Picture = LoadPicture(App.Path & "/wallpaper.jpg")
    Me.Image1.Picture = LoadPicture(App.Path & "/wallpaper.jpg")
    Me.StatusBar1.Visible = False           'change it after the work is completed
    
    Unload Form2
    Call switchoff

End Sub

Private Sub MDIForm_Resize()
    
    Me.Picture1.Height = Screen.Height
    Me.Image1.Height = Me.Height
    Me.Image1.Width = Me.Width
    Image1.Top = 0
    Image1.Left = 0
'    Me.Picture1.AutoSize = False
'    'resizing the picture so that it fits into the picturebox
'    Picture1.ScaleMode = 3
'    Picture1.AutoRedraw = True
'    'underscore in vb is used to tell that the code is not finished
'    'and it continues into the next line.It is used to make a single line code
'   'go multiple line to make it more readable
'    Picture1.PaintPicture Picture1.Picture, _
'        0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
'        0, 0, _
'        Picture1.Picture.Width / 26.46, _
'        Picture1.Picture.Height / 26.46
'    Picture1.Picture = Picture1.Image
'     Me.Picture1.Height = Me.Height

    'resizing form1
    Form1.Height = Me.Height / 2
    Form1.Width = Me.Width / 2
    Form1.Top = Me.Top + (Me.Height / 3.5)
    Form1.Left = Me.Left + (Me.Width / 4)
    'Form1.ZOrder 0 'ZOrder is the position, 0 being at the top and 1 being behind other form
    
    
    
'    Form2.height = Me.height / 2
'    Form2.width = Me.width / 2
'    Form2.Top = Me.Top + (Me.height / 3.5)
'    Form2.Left = Me.Left + (Me.width / 4)
    
    
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
ConnectDB
Cn.Execute ("update LOGINDETAILS set Logout_Date= " & Date & ", Logout_Time='" & Time & "', SessionActive=0 where UserName='" & UID & "' and SessionActive=1 ")
Set Rs = Nothing
Set Cn = Nothing
End
End Sub

Private Sub patientbookinggraph_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub patientdepartmentgraph_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub patientdiseasegraph_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub prescription_Click()
    Form9.Show vbModeless, Me
    Me.Visible = False
End Sub

Private Sub registration_Click()
    Form3.Show
End Sub

Private Sub release_Click()
     Form4.Show vbModeless, Me 'to show it always on top of the main form
    
    Form4.Height = Me.Height / 2
    Form4.Width = Me.Width / 2
    Form4.Top = Me.Top + (Me.Height / 3.5)
    Form4.Left = Me.Left + (Me.Width / 4)
        
    'call invisibleall
End Sub

Private Sub remainingbeds_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub reports_Click()
    Form7.Show vbModeless, Me 'to show it always on top of the main form
    Me.Visible = False
End Sub

Private Sub routinechart_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub signup_Click()
    Form1.Show vbModeless, Me 'to show it always on top of the main form
    Form1.SSTab1.Tab = 0
   
'    Form1.Height = Me.Height / 2
'    Form1.Width = Me.Width / 2
'    Form1.Top = Me.Top + (Me.Height / 3.5)
'    Form1.Left = Me.Left + (Me.Width / 4)
End Sub

Private Sub totalbillingchart_Click()
    Me.Hide
    Form12.Show
End Sub

Private Sub totalbookedbeds_Click()
    MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
End Sub

Private Sub UserButtonz1_Click()
    If UserButtonz1.Caption = "Patient Management" Then
        Me.UserButtonz6.Visible = True
        Call Visibleall
        Me.UserButtonz1.Caption = "Details"
        Me.UserButtonz2.Caption = "Registration"
        Me.UserButtonz3.Caption = "Release"
        Me.UserButtonz4.Caption = "Reports"
        Me.UserButtonz5.Caption = "Prescription"
        
    ElseIf UserButtonz1.Caption = "Details" Then
        Form2.Show

'        Form2.height = Me.height / 2
'        Form2.width = Me.width / 2
'        Form2.Top = Me.Top + (Me.height / 3.5)
'        Form2.Left = Me.Left + (Me.width / 4)
    ElseIf Me.UserButtonz1.Caption = "Patient-Booking Graph" Then
        MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"

    End If
    
End Sub

Private Sub UserButtonz2_Click()
    If UserButtonz2.Caption = "Doctor Management" Then
        Me.UserButtonz6.Visible = True
        Call Visibleall
        Me.UserButtonz2.Caption = "Availability"
        Me.UserButtonz3.Caption = "Doctor Wise Patient"
        Me.UserButtonz4.Caption = "Routine Chart"
        
        Me.UserButtonz1.Visible = False
        Me.UserButtonz5.Visible = False
    
    ElseIf UserButtonz2.Caption = "Registration" Then
        Form3.Show
    ElseIf UserButtonz2.Caption = "Issue Bill" Then
        Form10.Show
    ElseIf UserButtonz2.Caption = "Availability" Then
        MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    ElseIf Me.UserButtonz2.Caption = "Remaining Beds" Then
        MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    ElseIf Me.UserButtonz2.Caption = "Patient-Disease Graph" Then
        MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    End If
    
End Sub

Private Sub UserButtonz3_Click()
    If UserButtonz3.Caption = "Billing" Then
        Me.UserButtonz6.Visible = True
        Call Visibleall
        Me.UserButtonz2.Caption = "Issue Bill"
        Me.UserButtonz3.Caption = "Total Billing Chart"
        
        Me.UserButtonz1.Visible = False
        Me.UserButtonz4.Visible = False
        Me.UserButtonz5.Visible = False
    ElseIf Me.UserButtonz3.Caption = "Release" Then
        Form4.Show vbModeless, Me 'to show it always on top of the main form
    
        Form4.Height = Me.Height / 2
        Form4.Width = Me.Width / 2
        Form4.Top = Me.Top + (Me.Height / 3.5)
        Form4.Left = Me.Left + (Me.Width / 4)
        
        'call invisibleall
    ElseIf Me.UserButtonz3.Caption = "Doctor Wise Patient" Then
         MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    ElseIf Me.UserButtonz3.Caption = "Total Booked Beds" Then
         MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    ElseIf Me.UserButtonz3.Caption = "Total Billing Chart" Then
        Me.Hide
        Form12.Show
    End If
End Sub

Private Sub UserButtonz4_Click()
    If UserButtonz4.Caption = "Booking" Then
        Me.UserButtonz6.Visible = True
        Call Visibleall
        Me.UserButtonz2.Caption = "Remaining Beds"
        Me.UserButtonz3.Caption = "Total Booked Beds"
        
        Me.UserButtonz1.Visible = False
        Me.UserButtonz4.Visible = False
        Me.UserButtonz5.Visible = False
    ElseIf UserButtonz4.Caption = "Reports" Then
        Form7.Show vbModeless, Me 'to show it always on top of the main form
        Me.Visible = False
    ElseIf Me.UserButtonz4.Caption = "Routine Chart" Then
        MsgBox "Under Construction.......", vbInformation + vbOKOnly, "Under Construction"
    End If
End Sub

Private Sub UserButtonz5_Click()
    If UserButtonz5.Caption = "Charts and Graphs" Then
        Me.UserButtonz6.Visible = True
        Call Visibleall
        Me.UserButtonz1.Caption = "Patient-Booking Graph"
        Me.UserButtonz2.Caption = "Patient-Disease Graph"
        Me.UserButtonz3.Caption = "Doctor-Patient Graph"
        Me.UserButtonz4.Caption = "Patient-Department Graph"
        Me.UserButtonz5.Caption = "Doctor-Department Graph"
    ElseIf UserButtonz5.Caption = "Prescription" Then
        Form9.Show vbModeless, Me
        Me.Visible = False
    End If
End Sub

Private Sub UserButtonz6_Click()
    Call Backall
    Call Visibleall
    MDIForm1.UserButtonz6.Visible = False
End Sub

Public Function Visibleall()
    Me.UserButtonz1.Visible = True
    Me.UserButtonz2.Visible = True
    Me.UserButtonz3.Visible = True
    Me.UserButtonz4.Visible = True
    Me.UserButtonz5.Visible = True
End Function

Public Function Backall()
     With MDIForm1
        .UserButtonz1.Caption = "Patient Management"
        .UserButtonz2.Caption = "Doctor Management"
        .UserButtonz3.Caption = "Billing"
        .UserButtonz4.Caption = "Booking"
        .UserButtonz5.Caption = "Charts and Graphs"
    End With
End Function
