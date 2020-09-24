VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Import_CSV_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select the csv file to Import"
   ClientHeight    =   1560
   ClientLeft      =   3420
   ClientTop       =   1965
   ClientWidth     =   5940
   Icon            =   "Import_CSV_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5940
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton cmdExit4Small 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Browse 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Pick CSV File"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      Picture         =   "Import_CSV_Form.frx":263A
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Import_CSV_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Purpose : Converting CSV files to Text Files  ''''
''''          as per customisation required       ''''
''''                                              ''''
''''Date : 21st Nov '2003                         ''''
''''                                              ''''
''''Date Modified : 27th Dec '2003                ''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const JetDateFmt = "dd\/mm\/yyyy\"

Private Sub cmd_Browse_Click()
   ' CancelError is True.
     On Error GoTo errhandler
     Dim txtPath As String
     Dim EnableImport
     cmdImport.Enabled = True

    ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)|*.*|Comma Delimeted Files (*.CSV)|*.CSV|Test Files (*.txt)|*.txt"
    ' Specify default filter.
    CommonDialog1.FilterIndex = 2
    Me.CommonDialog1.InitDir = App.Path
    ' Display the Open dialog box.
    CommonDialog1.ShowOpen
    ' Call the open file procedure.
    'OpenFile (CommonDialog1.FileName)3
    txtPath = Me.CommonDialog1.FileName
    txtFileName.Text = txtPath
errhandler:
' User pressed Cancel button.
   Exit Sub
End Sub
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdExit4Small_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
cmdImport.Enabled = False
If txtFileName.Text = "" Then
    MsgBox "File Not Selected, Please Choose a File to Import"
Else
    Dim fso
    Dim act
    Dim total_imported_text
    Set fso = CreateObject("scripting.filesystemobject")
    'Set act = fso.OpenTextFile("C:\Documents and Settings\Administrator.BUGS\Desktop\Import CSV\cm11NOV2003bhav.csv")
    Set act = fso.OpenTextFile(Me.CommonDialog1.FileName)
    total_imported_text = act.ReadAll
    total_imported_text = Replace(total_imported_text, Chr(13), "*")
    total_imported_text = Replace(total_imported_text, Chr(10), "*")
    'Response.Write total_imported_text
    total_imported_text = Replace(total_imported_text, Chr(34), "")
    'Remove all the quotes (If your csv has quotes other than to seperate text
    'You may want to remove this modifier to the imported text
    total_split_text = Split(total_imported_text, "*")
    'Split the file up by comma
    total_num_imported = UBound(total_split_text)
    For i = 1 To total_num_imported - 1 '0 To total_num_imported '
        comma_split = Split(total_split_text(i), ",")
          On Error Resume Next
        If comma_split(0) <> "" Then
    
            Fileld2OfExcel = Trim(Mid(comma_split(0), 2))
            '****************Existing Condition*******************************
            'Check the column of the excel sheets if it is empty
            'if not then print then Row
            '****************As per Your Condition*******************************
            'A new text file will be created for each row with the text file name as the first column
            If Fileld2OfExcel <> "" Then
               '****************************************************
               'Debug.Print total_split_text(i)
               '****************************************************
               'Save Each Next Row that is Found
               If Dir(App.Path & "\Data\" & comma_split(0) & ".txt") = "" Then
                   Open App.Path & "\Data\" & comma_split(0) & ".txt" For Output As #1
                 '  Debug.Print comma_split(0)
                       Print #1, Format$(comma_split(10), JetDateFmt) & "," & comma_split(2) & "," & comma_split(3) & "," & comma_split(4) & "," & comma_split(5) & "," & comma_split(8) 'Mid(total_split_text(i), 2) & vbCrLf
                   Close #1
               
               Else
                    L_B_EqualFound = False
                    L_B_LessFound = False
                    L_B_GreaterFound = False
                    Dim myData As String
                    Dim dt1 As Date
                    Dim dt2 As Date
                    Dim dt3 As Date
                    myData = ""
                    mydata1 = ""
                    'Debug.Print comma_split(0)
                    Open App.Path & "\Data\" & comma_split(0) & ".txt" For Input As #1
                    Dim L_A_arr
                    ReDim L_A_arr(0)
                    Do While Not EOF(1)
                        Line Input #1, myData
                        If myData <> "" Then
                           ReDim Preserve L_A_arr(UBound(L_A_arr) + 1)
                           L_A_arr(UBound(L_A_arr)) = myData
                        End If
                    Loop
                    Close #1
                    For J = 1 To UBound(L_A_arr)
                           txt = Split(L_A_arr(J), ",")
                           dt1 = Format$(comma_split(10), JetDateFmt)
                           dt2 = txt(0)
                           If dt1 = dt2 Then
                              L_B_EqualFound = True
                              Exit For
                           End If
                    Next J
                    If L_B_EqualFound = False Then
                        Dim con As New ADODB.Connection
                        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Import.mdb;Persist Security Info=False"
                        con.Execute ("delete from TImport")
                        For J = 1 To UBound(L_A_arr)
                           sd = Split(L_A_arr(J), ",")
                           con.Execute ("insert into TImport values('" & sd(0) & "','" & sd(1) & "','" & sd(2) & "','" & sd(3) & "','" & sd(4) & "','" & sd(5) & "')")
                        Next J
                        con.Execute ("insert into TImport values('" & Format$(comma_split(10), "dd/mm/yyyy") & "','" & comma_split(2) & "','" & comma_split(3) & "','" & comma_split(4) & "','" & comma_split(5) & "','" & comma_split(8) & "')")
                        Dim rs As New ADODB.Recordset
                        'rs.Open "select * from Timport Order BY field1 asc", con
                        
                        rs.Open "SELECT TImport.*, TImport.field1 From TImport ORDER BY TImport.field1 dESC", con

                        
                        
                        If Not (rs.BOF Or rs.EOF) Then
                            rs.MoveFirst
                            Do While Not rs.EOF
                                If mydata1 = "" Then
                                'Debug.Print comma_split(0)
                                   mydata1 = rs(0) & "," & rs(1) & "," & rs(2) & "," & rs(3) & "," & rs(4) & "," & rs(5) & vbCrLf
                                Else
                                   mydata1 = mydata1 & rs(0) & "," & rs(1) & "," & rs(2) & "," & rs(3) & "," & rs(4) & "," & rs(5) & vbCrLf
                                End If
                                rs.MoveNext
                            Loop
                        End If

                        rs.Close
                        Set rs = Nothing
                        con.Execute ("delete from Timport")
                        Open App.Path & "\Data\" & comma_split(0) & ".txt" For Output As #1
                              Print #1, mydata1
                        Close #1

                    End If
                  End If
               End If
            End If
    Next i
    txtFileName.Text = ""
    MsgBox Me.CommonDialog1.FileTitle & " Has Been Imported"
End If
End Sub

Private Sub Form_Load()
     cmdImport.Enabled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Ans As Integer
Ans = MsgBox("Are you sure?", vbYesNo + vbExclamation, "CSV to Text")
If Ans = vbYes Then
' Close the Application
   End
   Unload Me
Else
 ' Cancel
   Cancel = True
End If
End Sub

