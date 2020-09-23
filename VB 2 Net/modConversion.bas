Attribute VB_Name = "modConversion"
Function OutputCode(FormName As String, FormTitle As String, FormHeight As Integer, FormWidth As Integer)
'// Change Ststus bar text
frmMain.stsBar.Panels(1).Text = "Status: Outputing Code"

'// Save the VB.NET code to the VB.NET Form
Open App.Path & "\" & frmMain.txtName.Text & ".vb" For Append As #1
'// Print the Code
Print #1, "Public Class " & FormName
Print #1, "    Inherits System.Windows.Forms.Form" & vbCrLf
Print #1, "#Region ""Windows Form Designer generated code """ & vbCrLf
Print #1, "    Public Sub New()"
Print #1, "        MyBase.New()" & vbCrLf
Print #1, "        'This call is required by the Windows Form Designer."
Print #1, "        InitializeComponent()" & vbCrLf
Print #1, "        'Add any initialization after the InitializeComponent() call" & vbCrLf
Print #1, "    End Sub" & vbCrLf
Print #1, "    'Form overrides dispose to clean up the component list."
Print #1, "    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)"
Print #1, "        If disposing Then"
Print #1, "            If Not (components Is Nothing) Then"
Print #1, "                components.Dispose()"
Print #1, "            End If"
Print #1, "        End If"
Print #1, "        MyBase.Dispose(disposing)"
Print #1, "    End Sub" & vbCrLf
Print #1, "    'Required by the Windows Form Designer"
Print #1, "    Private components As System.ComponentModel.IContainer" & vbCrLf
Print #1, "    'NOTE: The following procedure is required by the Windows Form Designer"
Print #1, "    'It can be modified using the Windows Form Designer. "
Print #1, "    'Do not modify it using the code editor."
Print #1, "    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()"
Print #1, "        '"
Print #1, "        '" & FormName
Print #1, "        '"
Print #1, "        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)"
Print #1, "        Me.ClientSize = New System.Drawing.Size(" & FormHeight & "," & FormWidth & ")"
Print #1, "    Me.Name = " & FormName
Print #1, "    Me.Text = " & FormTitle & vbCrLf
Print #1, "    End Sub" & vbCrLf
Print #1, "#End Region" & vbCrLf
Print #1, frmMain.txtNetCode.Text & vbCrLf
Print #1, "End Class"
'// Close the File
Close #1

frmMain.stsBar.Panels(1).Text = "Status: Output Complete"

End Sub



Private Function ReplaceC(MainStr As String, OldStr As String, NewStr As String) As String
On Error GoTo error
ReplaceC = ""
Dim NewStrString As String
Dim i As Integer
For i = 1 To Len(MainStr)
  If Mid(MainStr, i, Len(OldStr)) = OldStr Then
  NewStrString = NewStrString & NewStr
  i = i + Len(OldStr) - 1
  Else
  NewStrString = NewStrString & Mid(MainStr, i, 1)
  End If
DoEvents
If DoCancel = True Then Exit Function
Next i
ReplaceC = NewStrString
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Convert(Text As String)

frmMain.stsBar.Panels(1).Text = "Status: Converting code"

On Error GoTo error

Dim s(0) As String
'Hold data for Controls
s(0) = "Private Sub " & frmMain.txtName.Text & "_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)"
's(1) = "Next Control will go here"


Text$ = ReplaceC(Text$, "Private Sub Form_Load()", s(0))
'Text$ = ReplaceC(Text$, "Nothing to change yet", s(1))

DoEvents
If DoCancel = True Then Exit Function
Convert = Text$

frmMain.stsBar.Panels(1).Text = "Status: Code Conversion Complete"

Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

