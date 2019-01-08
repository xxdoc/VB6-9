VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registration"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cboCountry 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtEmailAgain 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ComboBox cboGender 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cboDay 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Country"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Verify Email"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Birthday"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRegister_Click()
    
    If txtName = "" Then
        MsgBox "Please enter your name ", vbExclamation, "Error"
        txtName.SetFocus
        Exit Sub      'this statement will exit the sub-routine
    End If

    If cboDay.ListIndex = -1 Then
        MsgBox "Please select a day for your birthday", _
        vbExclamation, "Error"
        cboDay.SetFocus
        Exit Sub
    End If

    If cboMonth.ListIndex = -1 Then
        MsgBox "Please select a month for your birthday", _
        vbExclamation, "Error"
        cboMonth.SetFocus
        Exit Sub
    End If

    If cboYear.ListIndex = -1 Then
        MsgBox "Please select a year for your birthday", _
        vbExclamation, "Error"
        cboYear.SetFocus
        Exit Sub
    End If

    If cboGender.ListIndex = -1 Then
        MsgBox "Please select gender", vbExclamation, "Error"
        cboGender.SetFocus
        Exit Sub
    End If


    If txtEmail.Text = "" Then
        MsgBox "Please enter your email address", _
        vbExclamation, "Error"
        txtEmail.SetFocus
        Exit Sub
    End If

    If txtEmailAgain.Text = "" Then
        MsgBox "Please re-enter your email address", _
        vbExclamation, "Error"
        txtEmailAgain.SetFocus
        Exit Sub
    End If

    If txtEmail.Text <> txtEmailAgain.Text Then
        MsgBox "email address mismatch !!!", vbExclamation, "Error"
        txtEmail.SetFocus
        Exit Sub
    End If

    If cboCountry.ListIndex = -1 Then
        MsgBox "Please select your country", vbExclamation, "Error"
        cboCountry.SetFocus
        Exit Sub
    End If
    
    strMessageString = "Name - " & txtName & Chr(13) & _
                       "Birthday - " & cboDay.Text & " " & cboMonth.Text & cboYear.Text & Chr(13) & _
                       "Gender - " & cboGender.Text & Chr(13) & _
                       "Email address - " & txtEmail.Text & Chr(13) & _
                       "Country -" & cboCountry.Text
   
    
    MsgBox strMessageString, vbInformation, "Registration Successful !!!"

    End
    
End Sub

Private Sub Form_Load()

    cboGender.AddItem "Male"
    cboGender.AddItem "Female"

    cboCountry.AddItem "England"
    cboCountry.AddItem "United States"
    cboCountry.AddItem "France"
    cboCountry.AddItem "India"
    cboCountry.AddItem "Germany"
    cboCountry.AddItem "China"
    cboCountry.AddItem "Japan"
    cboCountry.AddItem "Bangladesh"

    For d = 1 To 30
        cboDay.AddItem Str(d)
        d = Val(d)
    Next d

    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"

    For y = 1980 To 2012
        cboYear.AddItem Str(y)
        y = Val(y)
    Next y

End Sub

