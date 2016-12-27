VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmWardsSearchWizardAdmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wards Search Wizard"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
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
      Left            =   5400
      TabIndex        =   1
      Top             =   1420
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
      ItemData        =   "frmWardsSearchWizardAdmit.frx":0000
      Left            =   1680
      List            =   "frmWardsSearchWizardAdmit.frx":000A
      TabIndex        =   0
      Text            =   "-----------SELECT------------"
      Top             =   1420
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdWardsInfoTable 
      Height          =   4815
      Left            =   240
      TabIndex        =   4
      Top             =   2280
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
      Caption         =   "Wards Information Table"
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
   Begin VB.Label lblSearchType 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria"
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
      TabIndex        =   17
      Top             =   1470
      Width           =   1695
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
      TabIndex        =   16
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
      Left            =   3960
      TabIndex        =   15
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
      Left            =   4920
      TabIndex        =   14
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   600
      Top             =   1200
      Width           =   7455
   End
   Begin VB.Label lblSearchFor 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   1470
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   11
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Label lblWizardHeader1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wards Search Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   4335
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   0
      Picture         =   "frmWardsSearchWizardAdmit.frx":0024
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmWardsSearchWizardAdmit.frx":00C6
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
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
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
      TabIndex        =   8
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmWardsSearchWizardAdmit.frx":0164
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Search Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2655
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
      Left            =   960
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
      Left            =   4560
      TabIndex        =   5
      Top             =   1335
      Width           =   1215
   End
End
Attribute VB_Name = "frmWardsSearchWizardAdmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call Relevant_Ward_Selection   'Calling the Relevant_Ward_Selection Procedure to interact with the recordset
    
    Set dgrdWardsInfoTable.DataSource = rsRelevantWardsSelection  'Setting the DataSource of the DataGrid
    
End Sub



Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
End Sub

Private Sub dgrdWardsInfoTable_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsRelevantWardsSelection
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[WardID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[WardNo] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
        
        Set dgrdWardsInfoTable.DataSource = rsRelevantWardsSelection  'Setting the DataSource of the DataGrid
            
    Else
        
        Form_Load   'Calling the Form_Load Procedure
        
    End If
    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsRelevantWardsSelection.RecordCount > 0 Then
        
        
        With rsRelevantWardsSelection
            
            'Reset the textfields with the selected record
            frmAdmitPatient.txtWardID.Text = .Fields(0).Value
            frmAdmitPatient.txtWardNo.Text = .Fields(1).Value
            
            
            Unload Me   'Unload the Wizard
            
        End With
    
    Else    'Displaying an error message, asking the user to choose a record
    
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
        
    End If
    
End Sub


