VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Patient Details"
   ClientHeight    =   9825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   ScaleHeight     =   9825
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   10695
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18865
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   22
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Text            =   "Select a type"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    Me.Text1.ToolTipText = Me.Combo1.Text
    Me.Text1.ForeColor = vbGrayText
    Me.Text1 = Me.Text1.ToolTipText
    
    
    If Me.Text1.ToolTipText = "Released" Then
        Rs.Close
        Rs.CursorLocation = adUseClient
        Rs.Open "Select * from Patients where U_Name='" & UID & "' and Release='Yes'", Cn, adOpenDynamic, adLockOptimistic
        Set Me.DataGrid1.DataSource = Rs
        Me.DataGrid1.ReBind
        Me.DataGrid1.Refresh
    ElseIf Me.Text1.ToolTipText = "Not Yet Released" Then
        Rs.Close
        Rs.CursorLocation = adUseClient
        Rs.Open "Select * from Patients where U_Name='" & UID & "' and Release='No'", Cn, adOpenDynamic, adLockOptimistic
        Set Me.DataGrid1.DataSource = Rs
        Me.DataGrid1.ReBind
        Me.DataGrid1.Refresh
    End If
    
    
End Sub

Private Sub Form_Activate()
    MDIForm1.Visible = False
    Me.WindowState = vbMaximized
    
    Call ConnectDB
'Connect DB is defined to connect the database of this project
    Rs.CursorLocation = adUseClient
Rs.Open "SELECT * from Patients where U_Name='" & UID & "'", Cn, adOpenDynamic, adLockOptimistic
    
    Set Me.DataGrid1.DataSource = Rs
    
End Sub

Private Sub Form_Load()
    With Me.Combo1
        .AddItem "Name"
        .AddItem "Patient ID"
        .AddItem "Doctor's Name"
        .AddItem "Bed Number"
        .AddItem "Released"
        .AddItem "Not Yet Released"
    End With
    
    Me.Text1.ToolTipText = "Name"
    Me.Text1.ForeColor = vbGrayText
    Me.Text1 = Me.Text1.ToolTipText
    
    Me.Image1.Picture = LoadPicture(App.Path & "/blue.jpg")
    Me.Label1.ZOrder 0
    
    DataGrid1.ScrollBars = dbgBoth
    DataGrid1.EditActive = False
    
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    a = Trim(Me.Text1)
    Rs.Close
'searches the database for a particualr patient and his/her name,ID,doctor name or bed number
    Rs.CursorLocation = adUseClient
    If Me.Text1.ToolTipText = "Name" Then
Rs.Open "Select * from Patients where U_Name='" & UID & "' and  Patient_Name like'" & a & "%'", Cn, adOpenDynamic, adLockOptimistic
    ElseIf Me.Text1.ToolTipText = "Patient ID" Then
        Rs.Open "Select * from Patients where U_Name='" & UID & "' and PID like'" & a & "%'", Cn, adOpenDynamic, adLockOptimistic
    ElseIf Me.Text1.ToolTipText = "Doctor's Name" Then
Rs.Open "Select * from Patients where U_Name='" & UID & "' and  Doctor like'" & a & "%'", Cn, adOpenDynamic, adLockOptimistic
    ElseIf Me.Text1.ToolTipText = "Bed Number" Then
Rs.Open "SELECT * from Patients where U_Name='" & UID & "' and Bed_Number like'" & a & "%'", Cn, adOpenDynamic, adLockOptimistic
    ElseIf Me.Text1.ToolTipText = "Released" Then
Rs.Open "SELECT * from Patients where U_Name='" & UID & "' and Release='Yes'", Cn, adOpenDynamic, adLockOptimistic
    ElseIf Me.Text1.ToolTipText = "Not Yet Released" Then
Rs.Open "SELECT * from Patients where U_Name='" & UID & "' and Release='No'", Cn, adOpenDynamic, adLockOptimistic
    End If
    Set Me.DataGrid1.DataSource = Rs
    Me.DataGrid1.ReBind
    Me.DataGrid1.Refresh
End Sub
