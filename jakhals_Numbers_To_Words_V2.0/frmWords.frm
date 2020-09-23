VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWords 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(RK Jakhal's) Whole Number to words Canvertor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5160
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4680
      Top             =   3360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   3420
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3977
            MinWidth        =   2469
            Picture         =   "frmWords.frx":000C
            Text            =   "ABOUT"
            TextSave        =   "ABOUT"
            Key             =   "A"
            Object.ToolTipText     =   "Click to know about this Application."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2910
            MinWidth        =   2910
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   330
      Left            =   -600
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin MSComDlg.CommonDialog comDlgOpenFile 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Open Excel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1135
      Left            =   135
      TabIndex        =   16
      ToolTipText     =   "Write an Excel file with complete path in which you have numbers list."
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox cboSheetName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmWords.frx":045E
         Left            =   1440
         List            =   "frmWords.frx":0460
         TabIndex        =   2
         ToolTipText     =   "Select source worksheet from the list."
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Write an Excel file with complete path in which you have numbers list."
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Select an Excel file in which you have numbers list."
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblSheetName 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sheet Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Result Cell Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   2880
      TabIndex        =   13
      ToolTipText     =   "This is resulted Words list details here."
      Top             =   1200
      Width           =   2655
      Begin VB.ComboBox txtResultColumn 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Write column number of the Resulted words list converted from the numbers list."
         Top             =   840
         Width           =   1100
      End
      Begin VB.TextBox txtResultRow 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Write row number of the Resulted words list converted from the numbers list."
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Row Start:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Column:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame frmSource 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Source Cell Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "This is source Numbers list details here."
      Top             =   1200
      Width           =   2655
      Begin VB.ComboBox txtSourceColumn 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Write Column number of the source numbers list."
         Top             =   1440
         Width           =   1100
      End
      Begin VB.TextBox txtSourceRowEnd 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Write Ending row number of the source numbers list."
         Top             =   840
         Width           =   1100
      End
      Begin VB.TextBox txtSourceRowStrt 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Write Starting row number of the source numbers list."
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Row End:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Column:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Row Start:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdWords 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Write in Words"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click after filling all the fields to convert the numbers list to Words."
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Click to know about this Application."
      Top             =   3165
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "BOUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "Click to know about this Application."
      Top             =   3165
      Width           =   615
   End
   Begin VB.Shape shpAbout 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Height          =   225
      Left            =   90
      Top             =   3165
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Developer  :   Rajesh Jakhal
'Description:   This application can easily change numbers list entered
'               into Excel to number in words list.
'Input      :   1. Excel File with full address
'               2. Excel File's Sheet Name
'               3. Source starting row number
'               4. Source ending row number
'               5. Source starting column number
'               6. Result starting row number       FOR OUTPUT LOCATION
'               7. Result starting column number    FOR OUTPUT LOCATION
'Side Effect:   No side effect
'Future Plan:   Making or collecting utilities which will make easy to
'               operate for office workers. Helping friend can contact me.
'Contact No.:   (+91) 9896956660
'               rajesh_jakhal@rediffmail.com
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cboSheetName_GotFocus()
    cboSheetName.BackColor = &H80FFFF
End Sub

Private Sub cboSheetName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub cboSheetName_LostFocus()
    cboSheetName.BackColor = &HFFFFFF
End Sub

Private Sub cmdAbout_Click()
    lblAbout_Click
End Sub

Private Sub cmdAbout_GotFocus()
    shpAbout.Visible = True
End Sub

Private Sub cmdAbout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdAbout_LostFocus()
    shpAbout.Visible = False
End Sub

Private Sub cmdBrowse_Click()
    Dim str1 As String
    
    comDlgOpenFile.Filter = "Excel Files|*.xls"
    comDlgOpenFile.InitDir = App.Path
    comDlgOpenFile.ShowOpen
    str1 = comDlgOpenFile.FileName
    txtFile.Text = str1
    Call fillSheetNames(str1)
    
End Sub

Private Sub cmdBrowse_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub cmdWords_Click()
    Dim myCell As Double
    Dim i As Integer
    On Error GoTo r_Excel_Error
    Static xl As Excel.Application
    
    Me.MousePointer = vbHourglass
    
    If checkConstraints = False Then Exit Sub
    
    Set xl = New Excel.Application
    Dim wSheet As Excel.Worksheet
    
    With xl
        .Visible = False
        On Error GoTo r_File_Error
        .Workbooks.Open txtFile.Text        'App.Path & "\Work Order.xls"
        
        .ActiveWorkbook.Worksheets(cboSheetName.Text).Activate
        'Set wSheet = ActiveWorkbook.Worksheets("Work Order")
        Set wSheet = .ActiveSheet
        
        With wSheet                                             'in Numbers     = F8 : F305
            On Error GoTo r_Error
            For i = txtSourceRowStrt To txtSourceRowEnd         'in Words       = I8 : I305
                If Not Trim(.Range(txtSourceColumn & i).Value) = "" Then
                    If IsNumeric(.Range(txtSourceColumn & i).Value) Then
                        myCell = .Range(txtSourceColumn & i).Value
                        .Range(txtResultColumn & txtResultRow + i - txtSourceRowStrt).Value = mNumberToWords(Int(myCell))
                    Else
                        .Range(txtResultColumn & txtResultRow + i - txtSourceRowStrt).Value = "Error."
                    End If
                Else
                    .Range(txtResultColumn & txtResultRow + i - txtSourceRowStrt).Value = ""
                End If
            Next
        End With
        .Visible = True
    End With
    Set wSheet = Nothing
    Set xl = Nothing
    
    Me.MousePointer = vbNormal
    Exit Sub
r_Excel_Error:
    MsgBox "MS Excel does not exists in this system." & vbCr & vbCr & "System Requirements: MS-Office-97 or above should be installed on your system.", vbOKOnly + vbCritical, "Object Error..."
    Set wSheet = Nothing
    xl.Quit
    Set xl = Nothing
    Me.MousePointer = vbNormal
    Exit Sub
    
r_File_Error:
    MsgBox "This file does not exists.", vbOKOnly + vbCritical, "Open File Error..."
    Set wSheet = Nothing
    xl.Quit
    Set xl = Nothing
    Me.MousePointer = vbNormal
    Exit Sub
    
r_Error:
    MsgBox "There is some error in this file.", vbOKOnly + vbCritical, "Some Error..."
    Set wSheet = Nothing
    xl.Quit
    Set xl = Nothing
    
    Me.MousePointer = vbNormal
End Sub

Private Function checkConstraints() As Boolean
    If UCase(Right(txtFile.Text, 4)) <> ".XLS" Then
        checkConstraints = False
        MsgBox "Please select an excel file...", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    
    If txtSourceRowStrt.Text = "" Or Not IsNumeric(txtSourceRowStrt.Text) Then
        checkConstraints = False
        MsgBox "Please select Start Row for source.", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    If txtSourceRowEnd.Text = "" Or Not IsNumeric(txtSourceRowEnd.Text) Then
        checkConstraints = False
        MsgBox "Please select End Row for source.", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    
    If txtSourceColumn.Text = "" Or Not Len(Trim(txtSourceColumn.Text)) <= 2 Then
        checkConstraints = False
        MsgBox "Please select Column for source.", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    
    If txtResultRow.Text = "" Or Not IsNumeric(txtResultRow.Text) Then
        checkConstraints = False
        MsgBox "Please select Row for Result.", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    If txtResultColumn.Text = "" Or Not Len(Trim(txtResultColumn.Text)) <= 2 Then
        checkConstraints = False
        MsgBox "Please select Column for Result.", vbOKOnly + vbInformation, "Alert Message"
        Exit Function
    End If
    checkConstraints = True
End Function

Private Sub cmdWords_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set frmWords = Nothing
    Me.Icon = LoadResPicture(101, vbResIcon)
    Me.StatusBar1.Panels(1).Picture = LoadResPicture(101, vbResIcon)
    fillLists
    fillStatusBar
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Do you want to close?                 " & vbCr & "Yes/No", vbYesNo + vbInformation, "Confirmation...") = vbNo Then
        Cancel = True
    Else
        Me.Hide
        MsgBox "Thanks for using this utility." & vbCr & "This can reduce your burdon for writing a lot of words for numbers." & vbCr & vbCr & "You can freely use this code in your projects " & vbCr & "BUT GIVE YOUR VALUEABLE VOTE FOR THIS." & vbCr & vbCr & vbCr & "Thanking you again..."
    End If
    
End Sub

Private Sub lblAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 1 Then
        'MsgBox "Author: Rajesh Jakhal"
        frmAbout.Show vbModal
    End If
End Sub

Private Sub Timer1_Timer()
    fillStatusBar
End Sub

Private Sub Timer2_Timer()
    fillStatusBar
End Sub

Private Sub txtFile_GotFocus()
    txtFile.BackColor = &H80FFFF
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtFile_LostFocus()
    txtFile.BackColor = &HFFFFFF
End Sub

Private Sub txtResultColumn_GotFocus()
    txtResultColumn.BackColor = &H80FFFF
End Sub

Private Sub txtResultColumn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtResultColumn_LostFocus()
    txtResultColumn.BackColor = &HFFFFFF
End Sub

Private Sub txtResultRow_GotFocus()
    txtResultRow.BackColor = &H80FFFF
End Sub

Private Sub txtResultRow_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtResultRow_LostFocus()
    txtResultRow.BackColor = &HFFFFFF
End Sub

Private Sub txtSourceColumn_GotFocus()
    txtSourceColumn.BackColor = &H80FFFF
End Sub

Private Sub txtSourceColumn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSourceColumn_LostFocus()
    txtSourceColumn.BackColor = &HFFFFFF
End Sub

Private Sub txtSourceRowEnd_GotFocus()
    txtSourceRowEnd.BackColor = &H80FFFF
End Sub

Private Sub txtSourceRowEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSourceRowEnd_LostFocus()
    txtSourceRowEnd.BackColor = &HFFFFFF
End Sub

Private Sub txtSourceRowStrt_GotFocus()
    txtSourceRowStrt.BackColor = &H80FFFF
End Sub

Private Sub txtSourceRowStrt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    If IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub fillLists()
    Dim i As Integer, j As Integer
    
    'LIST txtSourceColumn +++++++++++++++++++++++++++++++++++++++++
    For i = 65 To 90
        txtSourceColumn.AddItem Chr(i)
    Next
    For i = 65 To 73            'from A to I
        If i < 73 Then
            For j = 65 To 90    'from A to Z
                txtSourceColumn.AddItem Chr(i) & Chr(j)
            Next
        Else
            For j = 65 To 86    'from A to V
                txtSourceColumn.AddItem Chr(i) & Chr(j)
            Next
        End If
    Next
    txtSourceColumn.ListIndex = 0
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'LIST txtResultColumn =========================================
    For i = 65 To 90
        txtResultColumn.AddItem Chr(i)
    Next
    For i = 65 To 73            'from A to I
        If i < 73 Then
            For j = 65 To 90    'from A to Z
                txtResultColumn.AddItem Chr(i) & Chr(j)
            Next
        Else
            For j = 65 To 86    'from A to V
                txtResultColumn.AddItem Chr(i) & Chr(j)
            Next
        End If
    Next
    txtResultColumn.ListIndex = 0
    '==============================================================
End Sub

Private Sub txtSourceRowStrt_LostFocus()
        txtSourceRowStrt.BackColor = &HFFFFFF
End Sub

Private Sub fillStatusBar()
    Dim Pm As String * 2
    Dim Hr As Integer
    
    StatusBar1.Panels(2).Text = IIf(Len(Day(Date)) < 2, "0" & Day(Date), Day(Date)) & "-" & MonthName(Month(Date)) & "-" & Year(Date)
    If Hour(Time) < 12 Then
        Pm = "AM"
        Hr = Hour(Time)
    Else
        Pm = "PM"
        Hr = Hour(Time) - 12
    End If
    StatusBar1.Panels(3).Text = "Time: " & IIf(Len(Trim(Hr)) < 2, "0" & Hr, Hr) & ":" & IIf(Len(Minute(Time)) < 2, "0" & Minute(Time), Minute(Time)) & ":" & IIf(Len(Second(Time)) < 2, "0" & Second(Time), Second(Time)) & " " & Pm
End Sub

Private Sub fillSheetNames(vPath As String)
    Dim i As Integer
    Static xl As Excel.Application
    Dim wSheet As Excel.Worksheet
    
    On Error GoTo r_Excel_Error

    Set xl = New Excel.Application
    Me.MousePointer = vbHourglass
    
    cboSheetName.Clear
    With xl
        .Visible = False
        .Workbooks.Open vPath
        
        For Each wSheet In .ActiveWorkbook.Worksheets
            cboSheetName.AddItem wSheet.Name
        Next wSheet
        Set wSheet = Nothing
    End With
    
    xl.Quit
    Set wSheet = Nothing
    Set xl = Nothing
    Me.MousePointer = vbNormal

    Exit Sub
r_Excel_Error:
    xl.Quit
    Set wSheet = Nothing
    Set xl = Nothing
    cboSheetName.AddItem "Error: MS Excel does not exists."
    
    Me.MousePointer = vbNormal
End Sub
