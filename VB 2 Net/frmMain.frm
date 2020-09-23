VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual baisc 6.0 to Visual Basic .NET"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   6855
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Continue"
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "300"
         Top             =   2280
         Width           =   6615
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "300"
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "My Form"
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Form1"
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Note* This will only convert the Form Code. In the Future I plan on Adding Support for the Rest of the Controls as well."
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "Form Width"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   6615
      End
      Begin VB.Label Label3 
         Caption         =   "Form Height"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Form Title"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label Label1 
         Caption         =   "Form Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   6855
      Begin VB.CommandButton cmdConvertNow 
         Caption         =   "Convert Now"
         Height          =   375
         Left            =   5640
         TabIndex        =   22
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste Code"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtVbCode 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtNetCode 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "Output Code"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Code"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Form Information"
            Object.ToolTipText     =   "Suppluy the Proper Form information."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Form Code"
            Object.ToolTipText     =   "Supply the Forms Code"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Converted .NET Code"
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4290
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
            Text            =   "Status: "
            TextSave        =   "Status: "
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
'// Check to see fi the User wishes to close the App and Quit?
Dim answer As Integer
answer = MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit?")

'// If answer = Yes
If answer = vbYes Then
'// End the Application
Unload Me
End
End If

'// If answer = No
If answer = vbNo Then
'// Act as if Nothing Happened
Exit Sub
End If

End Sub

Private Sub cmdClear_Click()
'// Clear te Textbox
txtVbCode.Text = ""
End Sub

Private Sub cmdConvert_Click()
'// Check to see if all fields are filled in.
If txtName.Text = "" Then
MsgBox "Please make sure all your Form's Data is Filled in", vbInformation, "Missing Data"
Exit Sub
Else:
TabStrip1.Enabled = True
TabStrip1.Tabs(2).Selected = True
Exit Sub
End If

If txtTitle.Text = "" Then
MsgBox "Please make sure all your Form's Data is Filled in", vbInformation, "Missing Data"
Exit Sub
Else:
TabStrip1.Enabled = True
TabStrip1.Tabs(2).Selected = True
Exit Sub
End If

If txtWidth.Text = "" Then
MsgBox "Please make sure all your Form's Data is Filled in", vbInformation, "Missing Data"
Exit Sub
Else:
TabStrip1.Enabled = True
TabStrip1.Tabs(2).Selected = True
Exit Sub
End If

If txtHeight.Text = "" Then
MsgBox "Please make sure all your Form's Data is Filled in", vbInformation, "Missing Data"
Exit Sub
Else:
TabStrip1.Enabled = True
TabStrip1.Tabs(2).Selected = True
Exit Sub
End If

End Sub

Private Sub cmdConvertNow_Click()
txtNetCode.Text = Convert(txtVbCode.Text)
TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub cmdCopy_Click()
'// Send the Text to the Clipboard
Clipboard.SetText txtNetCode.Text
End Sub

Private Sub cmdOutput_Click()
MsgBox "The Form will be saved to " & App.Path & "\" & frmMain.txtName.Text & ".vb", vbInformation, "File Information"
Call OutputCode(txtName.Text, txtTitle.Text, txtHeight.Text, txtWidth.Text)
End Sub

Private Sub cmdPaste_Click()
'// Paste Vb code to Text Box
txtVbCode = Clipboard.GetText
End Sub

Private Sub TabStrip1_Click()
'// Display the Frame from selected tab.
If TabStrip1.Tabs(1).Selected = True Then
Frame3.Visible = True
Frame2.Visible = False
Frame1.Visible = False
End If

If TabStrip1.Tabs(2).Selected = True Then
Frame3.Visible = False
Frame2.Visible = True
Frame1.Visible = False
End If

If TabStrip1.Tabs(3).Selected = True Then
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = True
End If

End Sub

Private Sub Form_Load()
stsBar.Panels(1).Text = "Status:"
End Sub
