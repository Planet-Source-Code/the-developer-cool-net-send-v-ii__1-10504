VERSION 5.00
Begin VB.Form frmnetSend 
   Caption         =   "Net Send"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3570
   Icon            =   "frmnetSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtUser 
      Height          =   315
      Left            =   870
      TabIndex        =   4
      Top             =   210
      Width           =   2700
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   1965
      Width           =   1020
   End
   Begin VB.TextBox txtMSGID 
      Height          =   735
      Left            =   870
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   2685
   End
   Begin VB.Label lblResponse 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   870
      TabIndex        =   6
      Top             =   1545
      Width           =   2685
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Response:"
      Height          =   195
      Left            =   30
      TabIndex        =   5
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "NT Login/ IP adress/ Machine"
      Height          =   570
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   735
      Width           =   690
   End
End
Attribute VB_Name = "frmnetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Simple Nothing To It
'This will only work on NT systems.
'It shells the CMD Net Send command.
Private sfPath As String
Private slogPath As String
Private sExecFile As String


Private Sub cmdSend_Click()
  
    MousePointer = vbHourglass
    Enabled = False
       Call RunBatch
       Call SentList(txtUser.Text)
       Call LoadUserList
    Enabled = True
    ZOrder
    Me.SetFocus
    MousePointer = vbDefault

End Sub
Private Sub RunBatch()
    
    Dim t$
    Dim iFile As Integer
    
On Error GoTo RunBatch_ERROR
    
    lblResponse.Caption = ""
      
  'First write the Execution Prog
  'See you need to write it to a batch file cos VB won't send the Pipe to a file
  'To see if it has been successful!!
  iFile = FreeFile
  Open sExecFile For Output As iFile
    t$ = "net send " & txtUser.Text & " " & Chr$(34) & txtMSGID.Text & Chr$(34) & " >" & slogPath
    Print #iFile, t$
  Close iFile
  
    'Now Shell the Prog
    Shell sExecFile, vbMinimizedNoFocus
    
    Do While Len(Dir(slogPath)) = 0
         DoEvents ' This loop is so we can varify the message has been sent
    Loop
    
    Do While Len(lblResponse.Caption) = 0
        Call CheckSuccess
        DoEvents
    Loop
    


      Kill slogPath
Exit Sub
RunBatch_ERROR:

    Select Case Err.Number
        Case 70 ' Permission denied Probably doing somting!!
        
            DoEvents: DoEvents
            Resume
            
        Case 53 'File not found
            Resume Next
            
    Case Else
        MsgBox Err.Description
    End Select

End Sub
Private Sub CheckSuccess()
  Dim t$
  Dim iFile As Integer
  
        iFile = FreeFile
        Open slogPath For Input As iFile
            Do While Not EOF(iFile)
                DoEvents
                Line Input #iFile, t$
                If Len(Trim(t$)) <> 0 Then
                    If CBool(InStr(1, UCase(t$), UCase$("successfully sent to " & txtUser.Text))) Then
                        lblResponse.Caption = " Successfully Sent."
                    Else
                        lblResponse.Caption = " Not Successful."
                    End If
                End If
                DoEvents
            Loop
        Close iFile
        
End Sub
Private Sub SentList(sUser)

    Dim iFile As Integer
    Dim tUser As String
    Dim SaveUser As Boolean

    
    SaveUser = True
    If Len(Dir(sfPath)) <> 0 Then
    'This Checks the USers sent to log file to
    'See if they exist in the recent list.
        iFile = FreeFile
        Open sfPath For Input As iFile
            Do While Not EOF(iFile)
            
                Input #iFile, tUser
                If tUser = sUser Then
                  SaveUser = False
                  Exit Do
                End If
            Loop
        Close iFile
    End If
    
    If SaveUser Then
        iFile = FreeFile
        Open sfPath For Append As iFile
        
            Print #iFile, sUser
        
        Close iFile
    End If
    
End Sub
Private Sub Form_Load()

    sfPath = App.Path & "\sent.dat"
    slogPath = "c:\temp\fnnetmsg.log" ' be careful when changing this it may not work!!
    sExecFile = "c:\temp\send.bat"
    Call LoadUserList
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmnetSend = Nothing
    
End Sub

Private Sub LoadUserList()

    Dim iFile As Integer
    Dim tUser As String
    Dim stUser As String ' Current value
    Dim sfPath As String
    
    sfPath = App.Path & "\sent.dat"
    stUser = txtUser.Text
    txtUser.Clear
    
    If Len(Dir(sfPath)) <> 0 Then
        iFile = FreeFile

        
        Open App.Path & "\sent.dat" For Input As iFile
            Do While Not EOF(iFile)
            
                Input #iFile, tUser
                txtUser.AddItem tUser
                txtUser.Text = tUser  ' Will be the last person sent to!!
            Loop
        Close iFile
    End If
    
    If Len(stUser) <> 0 Then txtUser.Text = stUser
    
End Sub
