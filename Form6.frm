VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PID"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   FillColor       =   &H000080FF&
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Generate Report"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PID :"
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DataEnvironment1.rsCommand2.Open "Select * from Patients where PID=" & Val(Me.Text1)
    
    If DataEnvironment1.rsCommand2.RecordCount < 1 Then
        MsgBox "Record Not Found", vbInformation + vbOKOnly, "Record"
        DataEnvironment1.rsCommand2.Close
        Exit Sub
    End If
    
    picloc = DataEnvironment1.rsCommand2.Fields(12)
    'MsgBox DataReport1.Sections(3).Controls(16).Name
    Set DataReport2.Sections(3).Controls("Image1").Picture = LoadPicture(picloc)
    
    DataReport2.Refresh
    DataReport2.Show
    DataReport2.WindowState = vbMaximized
    Unload Me
            
End Sub

