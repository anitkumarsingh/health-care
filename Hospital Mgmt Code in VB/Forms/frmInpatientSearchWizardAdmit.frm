VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInpatientSearchWizardAdmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inpatient Search Wizard"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
   Begin VB.ComboBox cboSearchType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInpatientSearchWizardAdmit.frx":0000
      Left            =   2160
      List            =   "frmInpatientSearchWizardAdmit.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdInpatientsInfoTable 
      Height          =   4815
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Inpatients Information Table"
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
   Begin VB.Label lblSearchText 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   1335
      Width           =   1215
   End
   Begin VB.Shape shpSearchFrame 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   960
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Inpatient Search Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   15
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmInpatientSearchWizardAdmit.frx":0063
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   14
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Health Care Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   13
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmInpatientSearchWizardAdmit.frx":0105
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   0
      Picture         =   "frmInpatientSearchWizardAdmit.frx":01A3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Inpatients Search Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   12
      Top             =   240
      Width           =   3255
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   9120
      X2              =   9120
      Y1              =   2160
      Y2              =   7440
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   2160
      Y2              =   7440
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   360
      X2              =   9120
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   360
      X2              =   9120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Durdans Hospital Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   11
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Specialization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   7
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label lblSearchFor 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1335
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   1080
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1335
      Width           =   855
   End
End
Attribute VB_Name = "frmInpatientSearchWizardAdmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean


Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
    'Calling the clearAllFields procedure in the Inpatients Admission Form
    frmAdmitPatient.clearAllFields
    
    'Calling the Form_Load procedure in the Inpatients Admission Form
    frmAdmitPatient.Form_Load
    
End Sub

Private Sub dgrdInpatientsInfoTable_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call Inpatients_Maintenance    'Calling the Inpatients_Maintenance Procedure to interact with the recordset
    
    Set dgrdInpatientsInfoTable.DataSource = rsInpatientMaintenance  'Setting the DataSource of the DataGrid
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsInpatientMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[PatientID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[Surname] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[NICNumber] Like '" & txtSearch.Text & "%" & "'"
                Case 4:
                    .Filter = "[AccountType] Like '" & txtSearch.Text & "%" & "'"
                Case 5:
                    .Filter = "[CompanyID] Like '" & txtSearch.Text & "%" & "'"
                Case 6:
                    .Filter = "[CompanyName] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
        
        Set dgrdInpatientsInfoTable.DataSource = rsInpatientMaintenance  'Setting the DataSource of the DataGrid
    
    Else
    
        Form_Load   'Calling the Form_Load procedure
        
    End If
    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsInpatientMaintenance.RecordCount > 0 Then
    
        Call Inpatients_Admission
        
        With rsInpatientsAdmission
        
            .MoveFirst
            
            While .EOF = False
            
                If .Fields(1).Value = rsInpatientMaintenance.Fields(0).Value Then
                
                    'Entering the values in the particular record into the fields on the interface
                    frmAdmitPatient.txtAdmissionID.Text = .Fields(0).Value
                    frmAdmitPatient.txtPatientID.Text = .Fields(1).Value
                    frmAdmitPatient.txtGuardianID.Text = .Fields(2).Value
                    frmAdmitPatient.txtAdmissionDate.Text = .Fields(3).Value
                    frmAdmitPatient.txtAdmissionTime.Text = .Fields(4).Value
                    frmAdmitPatient.cboPatientStatus.Text = .Fields(5).Value
                    frmAdmitPatient.txtReasonForStatus.Text = .Fields(6).Value
                    frmAdmitPatient.txtReferredDoctorID.Text = .Fields(7).Value
                    frmAdmitPatient.txtReferredDoctorName.Text = .Fields(8).Value
                    frmAdmitPatient.txtAssignedDoctorID.Text = .Fields(9).Value
                    frmAdmitPatient.txtAssignedDoctorName.Text = .Fields(10).Value
                    frmAdmitPatient.txtDepartmentID.Text = .Fields(11).Value
                    frmAdmitPatient.txtDepartmentName.Text = .Fields(12).Value
                    frmAdmitPatient.txtWardID.Text = .Fields(13).Value
                    frmAdmitPatient.txtWardNo.Text = .Fields(14).Value
                    frmAdmitPatient.txtRoomID.Text = .Fields(15).Value
                    frmAdmitPatient.txtAdditionalNotes.Text = .Fields(16).Value
    
                    Unload Me
                    
                    Exit Sub
                
                End If
                
                .MoveNext
            
            Wend
            
            Unload Me   'Unload the Wizard
            
        End With
    
    Else    'Displaying an error message, asking the user to choose a record
    
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
    End If
    
End Sub
        
        


