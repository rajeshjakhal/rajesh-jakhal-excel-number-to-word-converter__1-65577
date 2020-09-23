Attribute VB_Name = "modNumberToWords"
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

Public Function mNumberToWords(num As Double) As String
    If num = 0 Then mNumberToWords = "Zero "
    
    If num < 0 Then
        mNumberToWords = "(-ve) "
    End If
    mNumberToWords = mNumberToWords & mNumberToWordsIn(Abs(num)) & "Only."
    
End Function


Public Function mNumberToWordsIn(num As Double) As String
    Dim sFinal As String
    
    Select Case Len(CStr(num))
        Case 1
            mNumberToWordsIn = mNtoWSingle(num)
        Case 2
            mNumberToWordsIn = mNtoWDouble(num)
        Case 3
            If Val(Left(CStr(num), 1)) <> 0 Then
                mNumberToWordsIn = mNtoWSingle(Val(Left(CStr(num), 1))) & "Hundred " & mNtoWDouble(Val(Right(CStr(num), 2)))
            Else
                mNumberToWordsIn = mNtoWDouble(Val(Right(CStr(num), 2)))
            End If
        Case 4
            If Val(Left(CStr(num), 1)) <> 0 Then
                sFinal = mNtoWSingle(Val(Left(CStr(num), 1))) & "Thousand "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 3)))
            mNumberToWordsIn = sFinal
        Case 5
            If Val(Left(CStr(num), 2)) <> 0 Then
                sFinal = mNtoWDouble(Val(Left(CStr(num), 2))) & "Thousand "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 3)))
            mNumberToWordsIn = sFinal
        Case 6
            If Val(Left(CStr(num), 1)) <> 0 Then
                sFinal = mNtoWSingle(Val(Left(CStr(num), 1))) & "Lack "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 5)))
            mNumberToWordsIn = sFinal
        Case 7
            If Val(Left(CStr(num), 2)) <> 0 Then
                sFinal = mNtoWDouble(Val(Left(CStr(num), 2))) & "Lack "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 5)))
            mNumberToWordsIn = sFinal
        Case 8
            If Val(Left(CStr(num), 1)) <> 0 Then
                sFinal = mNtoWSingle(Val(Left(CStr(num), 1))) & "Crore "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 7)))
            mNumberToWordsIn = sFinal
        Case 9
            If Val(Left(CStr(num), 2)) <> 0 Then
                sFinal = mNtoWDouble(Val(Left(CStr(num), 2))) & "Crore "
            End If
            sFinal = sFinal & mNumberToWordsIn(Val(Right(CStr(num), 7)))
            mNumberToWordsIn = sFinal
    End Select
        
End Function

Private Function mNtoWSingle(n As Double) As String
    Select Case n
        Case 1
            mNtoWSingle = "One "
        Case 2
            mNtoWSingle = "Two "
        Case 3
            mNtoWSingle = "Three "
        Case 4
            mNtoWSingle = "Four "
        Case 5
            mNtoWSingle = "Five "
        Case 6
            mNtoWSingle = "Six "
        Case 7
            mNtoWSingle = "Seven "
        Case 8
            mNtoWSingle = "Eight "
        Case 9
            mNtoWSingle = "Nine "
    End Select
    
End Function

Private Function mNtoWDouble(n As Double) As String
    
    If n < 10 Then
        mNtoWDouble = mNtoWSingle(n)
    ElseIf n < 20 Then
        Select Case n
            Case 10
                mNtoWDouble = "Ten "
            Case 11
                mNtoWDouble = "Eleven "
            Case 12
                mNtoWDouble = "Tweleve "
            Case 13
                mNtoWDouble = "Thirteen "
            Case 14
                mNtoWDouble = "Forteen "
            Case 15
                mNtoWDouble = "Fifteen "
            Case 16
                mNtoWDouble = "Sixteen "
            Case 17
                mNtoWDouble = "Seventeen "
            Case 18
                mNtoWDouble = "Eighteen "
            Case 19
                mNtoWDouble = "Ninteen "
        End Select
    Else
        Select Case (n - (n Mod 10))
            Case 20
                mNtoWDouble = "Twenty " & mNtoWSingle(n Mod 10)
            Case 30
                mNtoWDouble = "Thirty " & mNtoWSingle(n Mod 10)
            Case 40
                mNtoWDouble = "Forty " & mNtoWSingle(n Mod 10)
            Case 50
                mNtoWDouble = "Fifty " & mNtoWSingle(n Mod 10)
            Case 60
                mNtoWDouble = "Sixty " & mNtoWSingle(n Mod 10)
            Case 70
                mNtoWDouble = "Seventy " & mNtoWSingle(n Mod 10)
            Case 80
                mNtoWDouble = "Eighty " & mNtoWSingle(n Mod 10)
            Case 90
                mNtoWDouble = "Ninty " & mNtoWSingle(n Mod 10)
        End Select
    End If
End Function
