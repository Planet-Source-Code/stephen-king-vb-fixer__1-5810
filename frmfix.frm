VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmfix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB FIle Fixer"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmfix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdfix 
      Caption         =   "&Fix"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtfile 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbvbpfile 
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"frmfix.frx":0442
   End
   Begin VB.Label lblfile 
      BackStyle       =   0  'Transparent
      Caption         =   "Project File"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmfix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbrowse_Click()
 ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Project Files" & _
    "(*.vbp)|*.vbp"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
   txtfile = CommonDialog1.filename
   
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdfix_Click()
Dim found As Integer



    rtbvbpfile.filename = txtfile
    rtbvbpfile.LoadFile txtfile
    
    found = rtbvbpfile.Find("Retained", , , rtfWholeWord)
    If found <> -1 Then
        rtbvbpfile.SelStart = found
        rtbvbpfile.SelLength = 10
        rtbvbpfile.SelText = ""
        rtbvbpfile.SaveFile txtfile
        MsgBox "Finished fixing " & txtfile, vbInformation + vbApplicationModal + vbOKOnly, "VB Fixer"
        txtfile = ""
        rtbvbpfile = ""
    Else
        MsgBox "The Project file " & txtfile & " is fine", vbInformation + vbApplicationModal + vbOKOnly, "VB Fixer"
        txtfile = ""
        Exit Sub
    End If
    
        
End Sub

Private Sub txtfile_Change()
    If Len(txtfile) > 0 Then
        cmdfix.Enabled = True
    Else
        cmdfix.Enabled = False
    End If
    
End Sub
