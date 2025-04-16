'---------- 1. Data Cleaning Functions ------------------

' CleanCompanyName: ????? string ???? ?????????? ???? ? punctuation ????? ????????????? ????? ???? ???
Function CleanCompanyName(ByVal s As String) As String
    Dim tokens() As String, token As String, cleanTokens() As String
    Dim i As Long, count As Long
    Dim removeList As Variant
    
    s = LCase(s)
    s = Replace(s, ".", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, "-", " ")
    s = Application.WorksheetFunction.Trim(s)
    
    ' ?????????? ?????? ??????
    removeList = Array("inc", "ltd", "llc", "corp", "corporation", _
                       "co", "company", "&", "and", "the")
    
    tokens = Split(s, " ")
    ReDim cleanTokens(0 To UBound(tokens))
    count = 0
    
    For i = LBound(tokens) To UBound(tokens)
        token = Application.WorksheetFunction.Trim(tokens(i))
        ' ??? token ???? ?? ?? ??? ??? ??????? ?? ??
        If token <> "" And Len(token) > 1 Then
            If IsError(Application.Match(token, removeList, 0)) Then
                cleanTokens(count) = token
                count = count + 1
            End If
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve cleanTokens(0 To count - 1)
    Else
        CleanCompanyName = ""
        Exit Function
    End If
    
    ' ??? ??????????? ??????? (sort ?? ???, ???? ??? ????)
    CleanCompanyName = Join(cleanTokens, " ")
End Function

'---------- 2. Similarity Functions ------------------

' Levenshtein: ????? string-?? ????? Levenshtein distance ??? ???
Function Levenshtein(s1 As String, s2 As String) As Long
    Dim i As Long, j As Long, l1 As Long, l2 As Long
    Dim d() As Long, cost As Long

    s1 = Trim(s1)
    s2 = Trim(s2)
    
    l1 = Len(s1)
    l2 = Len(s2)
    ReDim d(0 To l1, 0 To l2)
    
    For i = 0 To l1
        d(i, 0) = i
    Next i
    For j = 0 To l2
        d(0, j) = j
    Next j

    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = Application.WorksheetFunction.Min( _
                        d(i - 1, j) + 1, _
                        d(i, j - 1) + 1, _
                        d(i - 1, j - 1) + cost)
        Next j
    Next i

    Levenshtein = d(l1, l2)
End Function

' SimilarityPercentage_Levenshtein: Levenshtein method ???????? similarity percentage ?????? ???
Function SimilarityPercentage_Levenshtein(s1 As String, s2 As String) As Double
    Dim levDist As Long, maxLen As Long
    
    s1 = CleanCompanyName(s1)
    s2 = CleanCompanyName(s2)
    
    maxLen = Application.WorksheetFunction.Max(Len(s1), Len(s2))
    
    If maxLen = 0 Then
        SimilarityPercentage_Levenshtein = 100
        Exit Function
    End If
    
    levDist = Levenshtein(s1, s2)
    SimilarityPercentage_Levenshtein = ((maxLen - levDist) / maxLen) * 100
End Function

' TokenOverlapSimilarity: ?????? common token ??? ??? ???????? similarity percentage ??? ???
Function TokenOverlapSimilarity(s1 As String, s2 As String) As Double
    Dim tokens1() As String, tokens2() As String
    Dim i As Long, j As Long, commonCount As Long
    Dim totalTokens As Long
    
    s1 = CleanCompanyName(s1)
    s2 = CleanCompanyName(s2)
    
    If s1 = "" Or s2 = "" Then
        TokenOverlapSimilarity = 0
        Exit Function
    End If
    
    tokens1 = Split(s1, " ")
    tokens2 = Split(s2, " ")
    
    commonCount = 0
    For i = LBound(tokens1) To UBound(tokens1)
        For j = LBound(tokens2) To UBound(tokens2)
            If tokens1(i) = tokens2(j) Then commonCount = commonCount + 1
        Next j
    Next i
    
    totalTokens = Application.WorksheetFunction.Max(UBound(tokens1) + 1, UBound(tokens2) + 1)
    TokenOverlapSimilarity = (commonCount / totalTokens) * 100
End Function

'---------- 3. Final Matching Function ------------------

' CheckCompanyMatch: ????? ??????? weighted average similarity (Levenshtein ? Token Overlap)
' ????? threshold ??????? "Match" ?? "Not Match" ??????? ????
Function CheckCompanyMatch(apolloName As String, linkedinName As String, Optional threshold As Double = 70) As String
    Dim simLev As Double, simToken As Double, overallSim As Double
    Dim weightLev As Double, weightToken As Double
    
    ' ????? ?????? weight ???, ???????? ???? ????? adjust ????
    weightLev = 0.5
    weightToken = 0.5
    
    simLev = SimilarityPercentage_Levenshtein(apolloName, linkedinName)
    simToken = TokenOverlapSimilarity(apolloName, linkedinName)
    
    overallSim = (simLev * weightLev) + (simToken * weightToken)
    
    If overallSim >= threshold Then
        CheckCompanyMatch = "Match (" & Format(overallSim, "0.00") & "%)"
    Else
        CheckCompanyMatch = "Not Match (" & Format(overallSim, "0.00") & "%)"
    End If
End Function


