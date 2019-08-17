VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form Form11 
   Caption         =   "Issue Bill"
   ClientHeight    =   12060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form11"
   ScaleHeight     =   12060
   ScaleWidth      =   17490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
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
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Text            =   "<Select>"
      Top             =   6600
      Width           =   2175
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   975
      Left            =   7680
      TabIndex        =   8
      Top             =   12240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Finalise Bill"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16744576
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16711680
      ColorButtonUp   =   15309136
      ColorButtonDown =   16711680
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   10920
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   5
      Top             =   8760
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Do You have Mediclaim ?"
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   6840
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Charges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Charges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Operational Cost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Cost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Bill"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment  %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   10920
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   -360
      X2              =   18960
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   9840
      TabIndex        =   4
      Top             =   8760
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   8760
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   18960
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
'    MsgBox Blankchecking
    If Blankchecking = 1 Then
        Text1.SetFocus
        Exit Sub
    End If
        
    If Combo1.Text = "Yes" Then
        Text6.Enabled = False
        Text7.Enabled = False
        Label8.Enabled = True
        Label9.Enabled = True
        Text6 = Val(Text1) + Val(Text2) + Val(Text3) + Val(Text4)
        Text7 = Val(Text5)
        With Me
            .Label2.Enabled = False
            .Text1.Enabled = False
            .Label3.Enabled = False
            .Text2.Enabled = False
            .Label4.Enabled = False
            .Text3.Enabled = False
            .Label5.Enabled = False
            .Text4.Enabled = False
            .Label6.Enabled = False
            .Text5.Enabled = False

        End With
        'Text6.SetFocus
    Else
         With Me
            .Text6.Enabled = False
            .Text7.Enabled = False
            .Label8.Enabled = False
            .Label9.Enabled = False
            .Label2.Enabled = True
            .Text1.Enabled = True
            .Label3.Enabled = True
            .Text2.Enabled = True
            .Label4.Enabled = True
            .Text3.Enabled = True
            .Label5.Enabled = True
            .Text4.Enabled = True
            .Label6.Enabled = True
            .Text5.Enabled = True
            .Text6 = 0
            .Text7 = 0 'Val(Text1) + Val(Text2) + Val(Text3) + Val(Text4) + Val(Text5)
            .Text7.Enabled = False
        End With
    End If
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    
    Me.Image1.Picture = LoadPicture(App.Path & "\billing.jpeg")
    Me.Label1.ZOrder 0
        
    Me.Combo1.AddItem "Yes"
    Me.Combo1.AddItem "No"
    
    With Me
'        .Text1.Enabled = False
'        .Text2.Enabled = False
'        .Text3.Enabled = False
'        .Text4.Enabled = False
'        .Text5.Enabled = False
        
        .Text6.Enabled = False
        .Text7.Enabled = False
        .Label8.Enabled = False
        .Label9.Enabled = False
    End With
    
End Sub

Private Sub Form_Resize()
    With Me.Image1
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
        .Top = 0
        .Left = 0
    End With
    
    Label1.Left = (Me.ScaleWidth / 2) - 2500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.Visible = True
End Sub

Private Sub UserButtonz1_Click()
    On Error GoTo errorcheck
     If Blankchecking = 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    If Trim(Text8) = "" Or Combo1.Text = "<Select>" Then
        MsgBox "Insufficient Information", vbCritical + vbOKOnly, "Error"
        Combo1.Text = "No"
        Exit Sub
    End If
    
    Dim p As Currency
    p = Rs.Fields(0)
    
    Rs.Close
    Call ConnectDB
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * from Billing", Cn, adOpenDynamic, adLockOptimistic
    
    'Rs.MoveLast
    With Rs
        .AddNew
        .Fields(0) = p
        .Fields(1) = Me.Combo1.Text
        .Fields(2) = Val(Text1)
        .Fields(3) = Val(Text2)
        .Fields(4) = Val(Text3)
        .Fields(5) = Val(Text4)
        .Fields(6) = Val(Text5)
        .Fields(7) = Val(Text6)
        .Fields(8) = Val(Text7)
        .Fields(9) = Val(Text8)
        
        .Update
        .Close
        MsgBox "The Bill has been issued to PID: " & Val(p), vbInformation + vbOKOnly, "Bill Issued"
        
    End With
    
    Unload Me
    Exit Sub
    
errorcheck: MsgBox "Invalid input !", vbCritical + vbOKOnly, "Error"
            If Rs.State = 1 Then
                Rs.Close
            End If
            Combo1.Text = "No"
            Text1 = ""
            Text2 = ""
            Text3 = ""
            Text4 = ""
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text8 = ""
        With Me
            .Text6.Enabled = False
            .Text7.Enabled = False
            .Label8.Enabled = False
            .Label9.Enabled = False
            .Label2.Enabled = True
            .Text1.Enabled = True
            .Label3.Enabled = True
            .Text2.Enabled = True
            .Label4.Enabled = True
            .Text3.Enabled = True
            .Label5.Enabled = True
            .Text4.Enabled = True
            .Label6.Enabled = True
            .Text5.Enabled = True
        End With
        Combo1.Text = "No"
        Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
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
Private Sub Text4_KeyPress(KeyAscii As Integer)
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
Private Sub Text5_KeyPress(KeyAscii As Integer)
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
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
Private Sub Text8_KeyPress(KeyAscii As Integer)
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

Function Blankchecking() As Integer
    If Trim(Text1) = "" Or Trim(Text2) = "" Or Trim(Text3) = "" Or Trim(Text4) = "" Or Trim(Text5) = "" Then
        MsgBox "Insufficient Information", vbCritical + vbOKOnly, "Error"
        Blankchecking = 1
        Exit Function
    End If
        Blankchecking = 0
        Exit Function
End Function

