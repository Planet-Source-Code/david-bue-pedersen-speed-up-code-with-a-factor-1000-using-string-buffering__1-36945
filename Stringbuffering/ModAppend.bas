Attribute VB_Name = "ModAppend"
Dim S2(9) As String
Dim S3(9) As Long

Public Function GetStringBuffer(Index) As String
    GetStringBuffer = Left$(S2(Index), S3(Index))
End Function

Public Sub ClearBuffer(Index)
    S2(Index) = ""
    S3(Index) = 0
End Sub

Public Sub AppendToBuffer(ByVal AddString As String, Index As Long)
 
    Dim strTemp As String
    Dim lngLoop As Long
        
    ' Empty strings will cause a fatal error if not eliminated
    If Not Trim$(AddString) = "" Then
    
        '[1] DOES BUFFER NEED TO BE INCREASED?
        If S3(Index) + Len(AddString) > Len(S2(Index)) Then
            '  STORE S2
            strTemp = S2(Index)
            '  Increase memory storage bytes
            Do
                lngLoop = lngLoop + 1
                If (Len(S2(Index)) + (65536 * lngLoop)) >= (S3(Index) + Len(AddString)) Then
                    Exit Do
                End If
            Loop
            '  Resize buffer
            S2(Index) = String$(Len(S2(Index)) + (65536 * lngLoop), Chr(0))
            '  RESTORE S2
            Mid$(S2(Index), 1, S3(Index)) = strTemp
        End If
        
        '[2]  ADD STRING TO BUFFER
        Mid$(S2(Index), S3(Index) + 1, Len(AddString)) = AddString
        '[3]  SET STRING LENGTH IN BUFFER
        S3(Index) = S3(Index) + Len(AddString)
        
    End If
    
End Sub




