Attribute VB_Name = "Module1"

'API declaration to create default pdf viewer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public UserName As String
Public UID As String

Public Cn As New ADODB.Connection
Public Cm As New ADODB.Command
Public Rs As New ADODB.Recordset





Public Function ConnectDB()
    DataLocation = App.Path & "\Database\Signup.accdb"
    If Cn.State <> 1 Then ' if the database is not open
        Cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DataLocation & ";Persist Security Info=False"
    
        Cn.Open
'    Rs.Open "Select * from Signup", Cn, adOpenDynamic, adLockBatchOptimistic
'    Cm.ActiveConnection = Cn
    End If
    'cm.Execute
End Function


Public Function switchoff()
With MDIForm1
    .signup.Visible = True
    .adminlogin.Visible = True
    .hospitallogin.Visible = True
    .Logout.Visible = False
    
    
    .details.Enabled = False
    .registration.Enabled = False
    .release.Enabled = False
    .reports.Enabled = False
    .prescription.Enabled = False
       
    .availability.Enabled = False
    .doctorwisepatient.Enabled = False
    .routinechart.Enabled = False
        
    .Issuebill.Enabled = False
    .totalbillingchart.Enabled = False
    
    .remainingbeds.Enabled = False
    .totalbookedbeds.Enabled = False
        
    .patientbookinggraph.Enabled = False
    .patientdepartmentgraph.Enabled = False
    .patientdiseasegraph.Enabled = False
    .doctordepartmentgraph.Enabled = False
    .doctorpatientgraph.Enabled = False
    
    .Label1.Visible = False
    .UserButtonz1.Visible = False
    .UserButtonz2.Visible = False
    .UserButtonz3.Visible = False
    .UserButtonz4.Visible = False
    .UserButtonz5.Visible = False
    
    .UserButtonz6.Visible = False
    
End With
End Function

Public Function switchon() 'called in login but the first 3 lines of code not working
With MDIForm1
    .signup.Visible = False
    .adminlogin.Visible = False
    .hospitallogin.Visible = False
    .Logout.Visible = True
    
    .details.Enabled = True
    .registration.Enabled = True
    .release.Enabled = True
    .reports.Enabled = True
    .prescription.Enabled = True
    
    .availability.Enabled = True
    .doctorwisepatient.Enabled = True
    .routinechart.Enabled = True
    
    .Issuebill.Enabled = True
    .totalbillingchart.Enabled = True
    
    .remainingbeds.Enabled = True
    .totalbookedbeds.Enabled = True
    
    .patientbookinggraph.Enabled = True
    .patientdepartmentgraph.Enabled = True
    .patientdiseasegraph.Enabled = True
    .doctordepartmentgraph.Enabled = True
    .doctorpatientgraph.Enabled = True
    
    .Label1.Visible = True
    .UserButtonz1.Visible = True
    .UserButtonz2.Visible = True
    .UserButtonz3.Visible = True
    .UserButtonz4.Visible = True
    .UserButtonz5.Visible = True
End With
End Function

Function displayit()
    
    'dashboard code here
        
    
End Function







