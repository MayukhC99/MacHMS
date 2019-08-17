VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PID"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7035
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7035
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
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PID :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
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
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Call ConnectDB
    If Rs.State <> 1 Then 'so the database is closed
        Rs.CursorLocation = adUseClient
        Rs.Open "Select * from Patients where PID=" & Val(Trim(Text1)), Cn, adOpenDynamic, adLockOptimistic
    End If
    
    If Rs.RecordCount < 1 Then
        MsgBox "Patient record not found", vbInformation + vbOKOnly, "Records"
        Rs.Close
        Text1 = ""
        Text1.SetFocus
        Exit Sub
    End If
    
    MDIForm1.Visible = False
'    MdCost = Rnd(500) * 10000
'    OpCost = Rnd(500) * 300000
'    DocCost = Rnd(500) * 20000
'    BedCost = Rnd(500) * 3000
'    OthCost = Rnd(500) * 10000
'    Form11.Text1 = MdCost
'    Form11.Text2 = OpCost
'    Form11.Text3 = DocCost
'    Form11.Text4 = BedCost
'    Form11.Text5 = OthCost
    
    
    
    Form11.Show
    Unload Me
    
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

