VERSION 5.00
Begin VB.Form frmSukle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADDING MACHINE"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "frmSukle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Total/Change Monitor:"
      Height          =   3375
      Left            =   5160
      TabIndex        =   23
      Top             =   2880
      Width           =   4815
      Begin VB.Label lblOutput 
         Alignment       =   2  'Center
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Name:"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtItemname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantity:"
      Height          =   975
      Left            =   4560
      TabIndex        =   11
      Top             =   720
      Width           =   2775
      Begin VB.TextBox txtQuan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Item Price:"
      Height          =   975
      Left            =   7320
      TabIndex        =   10
      Top             =   720
      Width           =   2655
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   5040
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtMem 
      Height          =   285
      Left            =   5040
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Enter Cash:"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   4935
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4320
         Picture         =   "frmSukle.frx":0442
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtReport 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3135
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Total"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Price"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Item Name"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Quantity"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtMem256 
      Height          =   285
      Left            =   6000
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "SELECT:"
      Height          =   975
      Left            =   5040
      TabIndex        =   22
      Top             =   1800
      Width           =   4935
      Begin VB.CommandButton cmdAmount 
         Caption         =   "&Add Amount"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   1380
         TabIndex        =   29
         Top             =   430
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   2580
         TabIndex        =   28
         Top             =   430
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   3780
         TabIndex        =   27
         Top             =   430
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   430
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Jay-R Mega Mall"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   2160
      TabIndex        =   25
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmSukle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded By: Jaime B. Banaag Jr.
'IT-408 Project
'Project Titile "ADDING MACHINE

Option Explicit
Dim j, r, b, price, cash

Private Sub cmdAmount_Click()
On Error Resume Next
Dim Multi, total, msg

price = CDbl(txtPrice.Text)
   j = CDbl(txtMem.Text)
   r = CDbl(txtQuan.Text)
   b = CDbl(lblOutput.Caption)
If txtPrice.Text = "" Or txtQuan.Text = "" Then
msg = MsgBox("Check your inputs", vbInformation + vbOKOnly, "Please Check your work")
Else
If txtPrice.Text <= 0 Then
msg = MsgBox("Please name the price!", vbInformation + vbOKOnly, "Price please!")
Else
Beep
Multi = price * r
lblOutput.Caption = Multi
txtMem = Multi
total = b + Multi
txtTotal.Text = Multi
lblOutput.Caption = total
txtMem256.Text = total
txtPrice.Text = ""
txtQuan.Text = ""
txtItemname.SetFocus

txtReport.SelStart = Len(txtReport.Text)
txtReport.SelText = txtItemname.Text & vbTab & "    " & r & vbTab & "    " & price & vbTab & txtTotal.Text + vbCrLf
txtItemname.Text = ""
cmdReset.Enabled = True
cmdAmount.Enabled = False
 End If
End If
End Sub

Private Sub cmdChange_Click()
On Error GoTo Errors:
Dim msg
cash = CDbl(txtCash.Text)
b = CDbl(lblOutput.Caption)
lblOutput.Caption = cash - b
Beep
txtReport.SelStart = Len(txtReport.Text)
txtReport.SelText = vbTab & vbTab & vbTab & "_______" + vbCrLf & vbTab & vbTab & "Total: " & vbTab & txtMem256.Text + " P" & vbCrLf _
                    & "Cash : " & vbTab & txtCash.Text + " P" + vbCrLf _
                    & vbCrLf & "Your Change : " + lblOutput.Caption + " pesos" + vbCrLf + "THANK YOU!"
txtMem.Text = ""
txtMem256.Text = ""
cmdChange.Enabled = False
Exit Sub
Errors:
msg = MsgBox("You forgot to fill up something!", vbInformation + vbOKOnly, "Your cash please!")
End Sub

Private Sub cmdExit_Click()
Unload frmSukle
Set frmSukle = Nothing
End Sub

Private Sub cmdReset_Click()
On Error Resume Next
   j = CDbl(txtMem.Text)
   r = CDbl(txtQuan.Text)
   b = CDbl(lblOutput.Caption)
       j = 0
       r = 0
       b = 0
txtMem.Text = ""
txtTotal.Text = ""
txtMem256.Text = ""
txtQuan.Text = ""
txtItemname.Text = ""
txtPrice.Text = ""
txtCash.Text = ""
txtReport.Text = ""
lblOutput.Caption = "TOTAL"
txtItemname.SetFocus
End Sub

Private Sub Form_Load()
   cmdChange.Enabled = False
   cmdReset.Enabled = False
End Sub

Private Sub Label5_Click()
Dim msg
msg = MsgBox("Coded By: Jaime B. Banaag Jr.", vbInformation + vbOKOnly, "ADDING MACHINE")
End Sub

Private Sub txtCash_Change()
If txtCash.Text = "" Then
   cmdChange.Enabled = False
Else
   cmdChange.Enabled = True
End If
End Sub

Private Sub txtItemname_Change()
If txtItemname.Text = "" Then
   txtCash.Enabled = True
Else
   txtCash.Enabled = False
End If
End Sub

Private Sub txtPrice_Change()
If txtPrice.Text = "" Then
   txtCash.Enabled = True
Else
   txtCash.Enabled = False
   cmdAmount.Enabled = True
End If
End Sub

Private Sub txtQuan_Change()
If txtQuan.Text = "" Then
   txtCash.Enabled = True
Else
   txtCash.Enabled = False
End If
End Sub
