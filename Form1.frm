VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IMEI Generator v1.1"
   ClientHeight    =   2565
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   315
      Left            =   4451
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   315
      Left            =   3168
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2790
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Model:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Generated"
      Height          =   195
      Left            =   7080
      TabIndex        =   14
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "00000"
      Height          =   195
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "000000"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Type Allocation Code"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Reporting Body Identifier"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Full IMEI Presentation"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Luhn:"
      Height          =   195
      Left            =   4035
      TabIndex        =   5
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CheckSum"
      Height          =   195
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   465
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Copyright (C) 2012,2015  Nigel Todman (nigel.todman@gmail.com)
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
Dim x, y, z, tmp(), chk(), IMEI, TacTotal
Dim Build
Dim Valid As Boolean
Dim Tac(99999)
Dim TacTmp(99999)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub About_Click()
MsgBox ("IMEIGen v" & Build & " by Nigel Todman" & vbCrLf & "IMEI Generator v" & Build & vbCrLf & "E-Mail: nigel.todman@gmail.com" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "BitCoin: 127RB9VeUg1rVKDxyoB99SZ1D2FVHCVNRz" & vbCrLf & "Blog: nigelt.wordpress.com")
End Sub

Private Sub Command1_Click()
IMEI = Text1.Text

For x = 1 To Len(IMEI)
tmp(x) = Mid$(IMEI, x, 1)
Next x

For x = 1 To Len(IMEI)
    If x = 2 Or x = 4 Or x = 6 Or x = 8 Or x = 10 Or x = 12 Or x = 14 Then
        chk(x) = tmp(x) * 2
    Else
        chk(x) = tmp(x)
    End If
    If Len(chk(x)) > 1 Then
        chk(x) = Val(Mid$(chk(x), 1, 1)) + Val(Mid$(chk(x), 2, 1))
    End If
Next x


'00: Test and pre-series
'01: PCS Type Certification Review Board (PTCRB), North America
'10: DECT PP with GSM functionality
'30: Iridium
'31: ICO Satellite Management
'33: Directorate General of Posts and Telecommunications (DGTP), France
'35 + 44: British Approvals Board of Telecommunications (BABT), Great Britain
'45: National Telecom Agency (NTA), Denmark
'49: BZT / BAPT / Regulatory Authority / Federal Network Agency (FNA), Germany
'50: BZT ETS Certification GmbH, Germany
'51 + 52: CETECOM GmbH, Germany
'53: TÜV Product Service GmbH, Germany
'54: Phoenix Test-Lab GmbH, Germany
'86: Telecommunication Terminal Testing & Approval Forum (TAF), China
'91: Mobile Standard Alliance of India (MSAI), India
'99: Global Hexadecimal Administrator (GHA), the world


Label8.Caption = Mid$(IMEI, 1, 6) & "-" & Mid$(IMEI, 7, 2) & "-" & Mid$(IMEI, 9, 6) & "-" & Mid(IMEI, 15, 1)
If Mid$(IMEI, 1, 2) = "00" Then
Label9.Caption = "00 (Test and pre-series)"
ElseIf Mid$(IMEI, 1, 2) = "01" Then
Label9.Caption = "01 (PCS Type Certification Review Board (PTCRB), North America)"
ElseIf Mid$(IMEI, 1, 2) = "10" Then
Label9.Caption = "10 (DECT PP with GSM functionality)"
ElseIf Mid$(IMEI, 1, 2) = "30" Then
Label9.Caption = "30 (Iridium)"
ElseIf Mid$(IMEI, 1, 2) = "30" Then
Label9.Caption = "30 (Iridium)"
ElseIf Mid$(IMEI, 1, 2) = "31" Then
Label9.Caption = "31 (ICO Satellite Management)"
ElseIf Mid$(IMEI, 1, 2) = "33" Then
Label9.Caption = "33 (Directorate General of Posts and Telecommunications (DGTP), France)"
ElseIf Mid$(IMEI, 1, 2) = "33" Then
Label9.Caption = "33 (Directorate General of Posts and Telecommunications (DGTP), France)"
ElseIf Mid$(IMEI, 1, 2) = "35" Then
Label9.Caption = "35 (British Approvals Board of Telecommunications (BABT), Great Britain)"
ElseIf Mid$(IMEI, 1, 2) = "44" Then
Label9.Caption = "44 (British Approvals Board of Telecommunications (BABT), Great Britain)"
ElseIf Mid$(IMEI, 1, 2) = "45" Then
Label9.Caption = "45 (National Telecom Agency (NTA), Denmark)"
ElseIf Mid$(IMEI, 1, 2) = "49" Then
Label9.Caption = "49 (BZT / BAPT / Regulatory Authority / Federal Network Agency (FNA), Germany)"
ElseIf Mid$(IMEI, 1, 2) = "50" Then
Label9.Caption = "50 (BZT ETS Certification GmbH, Germany)"
ElseIf Mid$(IMEI, 1, 2) = "51" Then
Label9.Caption = "51 (CETECOM GmbH, Germany)"
ElseIf Mid$(IMEI, 1, 2) = "52" Then
Label9.Caption = "52 (CETECOM GmbH, Germany)"
ElseIf Mid$(IMEI, 1, 2) = "53" Then
Label9.Caption = "53 (TÜV Product Service GmbH, Germany)"
ElseIf Mid$(IMEI, 1, 2) = "54" Then
Label9.Caption = "54 (Phoenix Test-Lab GmbH, Germany)"
ElseIf Mid$(IMEI, 1, 2) = "86" Then
Label9.Caption = "86 (Telecommunication Terminal Testing & Approval Forum (TAF), China)"
ElseIf Mid$(IMEI, 1, 2) = "91" Then
Label9.Caption = "91 (Mobile Standard Alliance of India (MSAI), India)"
ElseIf Mid$(IMEI, 1, 2) = "98" Then
Label9.Caption = "98 (British Approvals Board of Telecommunications (BABT), Great Britain)"
ElseIf Mid$(IMEI, 1, 2) = "99" Then
Label9.Caption = "99 (Global Hexadecimal Administrator (GHA), the world)"
Else
Label9.Caption = Mid$(IMEI, 1, 2) & " (Unknown)"
End If
Label10.Caption = Mid$(IMEI, 1, 8)
Label11.Caption = Mid$(IMEI, 9, 6)
y = 0
Open VB.App.Path & "\tac1.csv" For Input As #1
Do While Not EOF(1)
y = y + 1
Line Input #1, Tac(y)
Loop
TacTotal = y
Close #1
'ReDim TacTmp(TacTotal)

For x = 2 To TacTotal
If Mid$(IMEI, 1, 8) = Mid$(Tac(x), 1, 8) Then
y = Split(Tac(x), ",")
Label13.Visible = True
Label14.Visible = True
Label13.Caption = "Manufacturer: " & y(1)
Label14.Caption = "Model: " & y(2)
End If
Next x

'490154203237518
'Full IMEI Presentation  490154-20-323751-8
'Reporting Body Identifier   49
'Type Approval Code  490154
'Final Assembly Code     20
'Serial Number   323751
'Check Digit     8
'358146031711756
'358391010609437
'353081035138133
'358428030277769
'354354030824489
'354354030969284
'354354030824331
'354354030831948
'353081039644417

Label1.Caption = "Check: " & chk(1) & chk(2) & chk(3) & chk(4) & chk(5) & chk(6) & chk(7) & chk(8) & chk(9) & chk(10) & chk(11) & chk(12) & chk(13) & chk(14)
z = chk(1) + chk(2) + chk(3) + chk(4) + chk(5) + chk(6) + chk(7) + chk(8) + chk(9) + chk(10) + chk(11) + chk(12) + chk(13) + chk(14)
Label2.Caption = "CheckSum: " & z
Label3.Caption = "Luhn: " & Val(Mid(IMEI, Len(IMEI), 1))
z = z + Val(Mid(IMEI, Len(IMEI), 1))
Label2.Caption = Label2.Caption & "(" & z & ")"


If z = 10 Then
Valid = True
ElseIf z = 20 Then
Valid = True
ElseIf z = 30 Then
Valid = True
ElseIf z = 40 Then
Valid = True
ElseIf z = 50 Then
Valid = True
ElseIf z = 60 Then
Valid = True
ElseIf z = 70 Then
Valid = True
ElseIf z = 80 Then
Valid = True
ElseIf z = 90 Then
Valid = True
ElseIf z = 100 Then
Valid = True
ElseIf z = 110 Then
Valid = True
ElseIf z = 120 Then
Valid = True
ElseIf z = 130 Then
Valid = True
ElseIf z = 140 Then
Valid = True
ElseIf z = 150 Then
Valid = True
ElseIf z = 160 Then
Valid = True
Else
Valid = False
MsgBox "IMEI is INVALID"
End If

If Valid = True Then
MsgBox "IMEI is VALID"
End If
End Sub

Private Sub Command2_Click()
Randomize Timer
Label12.Visible = True
Label13.Visible = False
Label14.Visible = False
y = 0
Do
y = y + 1
IMEI = Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1 & Int(Rnd * 9) + 1
Text1.Text = IMEI
Form1.Refresh
Text1.Refresh
Label12.Caption = "Generated: " & y
Sleep (50)
For x = 1 To Len(IMEI)
tmp(x) = Mid$(IMEI, x, 1)
Next x

For x = 1 To Len(IMEI)
    If x = 2 Or x = 4 Or x = 6 Or x = 8 Or x = 10 Or x = 12 Or x = 14 Then
        chk(x) = tmp(x) * 2
    Else
        chk(x) = tmp(x)
    End If
    If Len(chk(x)) > 1 Then
        chk(x) = Val(Mid$(chk(x), 1, 1)) + Val(Mid$(chk(x), 2, 1))
    End If
Next x

Label1.Caption = "Check: " & chk(1) & chk(2) & chk(3) & chk(4) & chk(5) & chk(6) & chk(7) & chk(8) & chk(9) & chk(10) & chk(11) & chk(12) & chk(13) & chk(14)
z = chk(1) + chk(2) + chk(3) + chk(4) + chk(5) + chk(6) + chk(7) + chk(8) + chk(9) + chk(10) + chk(11) + chk(12) + chk(13) + chk(14)
Label2.Caption = "CheckSum: " & z
Label3.Caption = "Luhn: " & Val(Mid(IMEI, Len(IMEI), 1))
z = z + Val(Mid(IMEI, Len(IMEI), 1))
Label2.Caption = Label2.Caption & "(" & z & ")"


If z = 10 Then
Valid = True
ElseIf z = 20 Then
Valid = True
ElseIf z = 30 Then
Valid = True
ElseIf z = 40 Then
Valid = True
ElseIf z = 50 Then
Valid = True
ElseIf z = 60 Then
Valid = True
ElseIf z = 70 Then
Valid = True
ElseIf z = 80 Then
Valid = True
ElseIf z = 90 Then
Valid = True
ElseIf z = 100 Then
Valid = True
ElseIf z = 110 Then
Valid = True
ElseIf z = 120 Then
Valid = True
ElseIf z = 130 Then
Valid = True
ElseIf z = 140 Then
Valid = True
ElseIf z = 150 Then
Valid = True
ElseIf z = 160 Then
Valid = True
Else
Valid = False
'MsgBox "IMEI is INVALID"
End If

If Valid = True Then
'MsgBox "IMEI is VALID"
End If
Loop Until Valid = True
End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Private Sub Form_Load()
ReDim tmp(16)
ReDim chk(16)
Build = "1.0"
Form1.Caption = "IMEI Generator v" & Build
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Text1.Text = "490154203237518"
End Sub

