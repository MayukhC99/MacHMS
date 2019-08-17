VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PID"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Generate Bill"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DataEnvironment1.rsCommand1.Open "Select * from Billing where PID=" & Val(Me.Text1)
    
    If DataEnvironment1.rsCommand1.RecordCount < 1 Then
        MsgBox "Record Not Found", vbInformation + vbOKOnly, "Record"
        DataEnvironment1.rsCommand1.Close
        Exit Sub
    End If
    
    Call ConnectDB
    If Rs.State = 1 Then
        Rs.Close
    End If
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from Patients where PID=" & Val(Me.Text1), Cn, adOpenDynamic, adLockOptimistic
    picloc = CStr(Rs.Fields(12))
    'MsgBox DataReport1.Sections(3).Controls(16).Name
    Set DataReport3.Sections(3).Controls("Image1").Picture = LoadPicture(picloc)
'    DataReport3.Refresh
    DataReport3.Sections(3).Controls("Label11").Caption = Rs.Fields(1)
    DataReport3.Sections(3).Controls("Label12").Caption = Rs.Fields(2)
    DataReport3.Sections(3).Controls("Label13").Caption = Rs.Fields(3)
    Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from Billing where PID=" & Val(Me.Text1), Cn, adOpenDynamic, adLockOptimistic
    Dim a As Currency
    Dim total As Currency
    a = Val(Rs.Fields(2)) + Val(Rs.Fields(3)) + Val(Rs.Fields(4)) + Val(Rs.Fields(5)) + Val(Rs.Fields(6))
    total = a - ((Val(Rs.Fields(9)) / 100) * a)
    Rs.Close
    DataReport3.Sections(3).Controls("Label16").Caption = CStr(total)
    
    
    DataReport3.Refresh
    DataReport3.Show
    DataReport3.WindowState = vbMaximized
    
    Unload Me
            
End Sub

