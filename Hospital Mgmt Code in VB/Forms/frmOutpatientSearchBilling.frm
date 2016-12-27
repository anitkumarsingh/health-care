VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOutpatientSearchBilling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outpatient Search Wizard"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
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
      ItemData        =   "frmOutpatientSearchBilling.frx":0000
      Left            =   2160
      List            =   "frmOutpatientSearchBilling.frx":0019
      TabIndex        =   2
      Text            =   "-----------SELECT------------"
      Top             =   1320
      Width           =   2295
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgrdOutpatientsInfoTable 
      Height          =   4815
      Left            =   480
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
      Caption         =   "Outpatients Information Table"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   1335
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   960
      Top             =   1080
      Width           =   7455
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
      Left            =   4560
      TabIndex        =   16
      Top             =   1335
      Width           =   1215
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
      Left            =   6600
      TabIndex        =   15
      Top             =   3285
      Width           =   975
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
      Left            =   3840
      TabIndex        =   14
      Top             =   3285
      Width           =   975
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
      Left            =   4800
      TabIndex        =   13
      Top             =   2535
      Width           =   1335
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
      Left            =   960
      TabIndex        =   12
      Top             =   2535
      Width           =   1335
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
      Index           =   0
      Left            =   2760
      TabIndex        =   11
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   240
      X2              =   9000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   240
      X2              =   9000
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   2160
      Y2              =   7440
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   9000
      X2              =   9000
      Y1              =   2160
      Y2              =   7440
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Outpatients Search Wizard"
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
      TabIndex        =   10
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   -120
      Picture         =   "frmOutpatientSearchBilling.frx":0072
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   -120
      Picture         =   "frmOutpatientSearchBilling.frx":0114
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label lblWizardFooter 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   8520
      Width           =   3735
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
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   -120
      Picture         =   "frmOutpatientSearchBilling.frx":01B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Outpatients Search Wizard"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   3135
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1335
      Width           =   855
   End
   Begin VB.Shape shpSearchFrame 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   840
      Top             =   1080
      Width           =   7455
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1335
      Width           =   1215
   End
End
Attribute VB_Name = "frmOutpatientSearchBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean
Dim strCorporateID As String    'This variable will hold the Corporate ID of the sponsoring company
Dim strDoctorID As String   'This variable will hold the Assigned Doctor's ID
Dim totalWithoutDiscount As Double  'This variable will hold the addition of the Total and The VAT
Dim totalWithDiscount As Double 'This variable will hold the Total after subtracting the discount


Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
End Sub

Private Sub dgrdOutpatientsInfoTable_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call Outpatients_Maintenance    'Calling the Outpatients_Maintenance Procedure to interact with the recordset
    
    Set dgrdOutpatientsInfoTable.DataSource = rsOutpatientsMaintenance  'Setting the DataSource of the DataGrid
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsOutpatientsMaintenance
        
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
        
        Set dgrdOutpatientsInfoTable.DataSource = rsOutpatientsMaintenance  'Setting the DataSource of the DataGrid
    
    Else
    
        Form_Load   'Calling the Form_Load procedure
        
    End If
    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsOutpatientsMaintenance.RecordCount > 0 Then
    
        With rsOutpatientsMaintenance
        
            'Reset the textfields with the selected record
            frmOPDOverallBilling.txtPatientID.Text = .Fields(0).Value
            frmOPDOverallBilling.txtFirstName.Text = .Fields(1).Value
            frmOPDOverallBilling.txtSurname.Text = .Fields(2).Value
            frmOPDOverallBilling.txtAccountType.Text = .Fields(11).Value
            strCorporateID = .Fields(12).Value
            frmOPDOverallBilling.txtAssignedDoctorID.Text = .Fields(14).Value
        End With
    
    
        'Here, I am including the corporate discount if the patient is a corporate patient
        If frmOPDOverallBilling.txtAccountType.Text = "Corporate" Then
            Call Companies_Maintenance  'Calling the Companies_Maintenance recordset
            With rsCompaniesMaintenance
                .MoveFirst
                Do While .EOF = False
                    If .Fields(0).Value = strCorporateID Then
                        frmOPDOverallBilling.txtDiscount = .Fields(6).Value
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End With
        Else
            frmOPDOverallBilling.txtDiscount.Text = ""
        End If
    
        
        'Here, I am calculating the Doctor's Charges
        Call Doctors_Maintenance
        With rsDoctorsMaintenance
            .MoveFirst
            Do While .EOF = False
                If .Fields(0).Value = frmOPDOverallBilling.txtAssignedDoctorID.Text Then
                    frmOPDOverallBilling.txtDoctorsCharges.Text = .Fields(12).Value
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
        End With
    
    
        'Here, I am displaying the Hospital Charges
        frmOPDOverallBilling.txtHospitalCharges.Text = 1000

        
        'Here, I am calculating the total of all costs
        frmOPDOverallBilling.txtTotal.Text = Val(frmOPDOverallBilling.txtDoctorsCharges.Text) + Val(frmOPDOverallBilling.txtHospitalCharges.Text)

    
        'Here, I am calculating the VAT amount
        frmOPDOverallBilling.txtVAT.Text = Val(frmOPDOverallBilling.txtTotal.Text) * 0.15

        
        'Here, I am calculating the NETT TOTAL
        If frmOPDOverallBilling.txtDiscount.Text = "" Then
            frmOPDOverallBilling.txtNettTotal.Text = Val(frmOPDOverallBilling.txtTotal.Text) + Val(frmOPDOverallBilling.txtVAT.Text)
        Else
            totalWithoutDiscount = Val(frmOPDOverallBilling.txtTotal.Text) + Val(frmOPDOverallBilling.txtVAT.Text)
            totalWithDiscount = totalWithoutDiscount - ((Val(frmOPDOverallBilling.txtDiscount.Text) / 100) * totalWithoutDiscount)
            frmOPDOverallBilling.txtNettTotal.Text = totalWithDiscount
        End If
            
        Unload Me   'Unloading the form

    
    
    
    Else    'Displaying an error message, asking the user to choose a record
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
    End If
    
End Sub
        
        






