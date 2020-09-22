VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
Private Sub Form_Load()
Dim strtest As String

    MsgBox "apending the word ""teststring"" 20000 times using normal appending and stringbuffering." & vbCrLf & vbCrLf & "THIS CAN TAKE 20-30 SECONDS! BE PATIENT, THE RESULT IS ABSOLUTELY AMAZING!"

    
    '***************Start the first loop****************************
        time1 = Time
            For a = 0 To 20000
                  AppendToBuffer "teststring ", 1 ' "teststring" is the string you want to append, the argument 1 means stringbuffer one. i you can dim S2() and S3() as big as you want for more buffers
            Next a
        time1 = DateDiff("s", time1, Time)
        
    MsgBox "Using stringbuffering we spend " & time1 & " seconds" & vbCrLf & vbCrLf & "NOW BE PATIENT"
    '***************Clear some memory****************************
    ClearBuffer (1)
    '***************Starting second loop****************************
        time2 = Time
            For a = 0 To 20000
                   strtest = strtest & "teststring "
            Next a
        time2 = DateDiff("s", time2, Time)
      
     '***************Prompting results****************************
      
    MsgBox "Using normal appending we spend " & time2 & " seconds"
    
    MsgBox "Using stringbuffering: " & time1 & vbCrLf & "Normal appending: " & time2 & vbCrLf & vbCrLf & "WHAT A DIFFERENCE!!! Have fun using this code snippet, David"


    'MsgBox GetStringBuffer(1)
End Sub


