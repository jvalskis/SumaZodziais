Attribute VB_Name = "SumaZodziais"
Option Explicit

Private Function ones() As Variant
    ones = Array("", "VIENAS", "DU", "TRYS", "KETURI", "PENKI", "ÐEÐI", "SEPTYNI", "AÐTUONI", "DEVYNI")
End Function
Private Function teens() As Variant
    teens = Array("DEÐIMT", "VIENUOLIKA", "DVYLIKA", "TRYLIKA", "KETURIOLIKA", "PENKIOLIKA", "ÐEÐIOLIKA", "SEPTYNIOLIKA", "AÐTUONIOLIKA", "DEVYNIOLIKA")
End Function
Private Function tens() As Variant
    tens = Array("", "DEÐIMT", "DVIDEÐIMT", "TRISDEÐIMT", "KETURIASDEÐIMT", "PENKIASDEÐIMT", "ÐEÐIASDEÐIMT", "SEPTYNIASDEÐIMT", "AÐTUONIASDEÐIMT", "DEVYNIASDEÐIMT")
End Function
Private Function hundreds() As Variant
    hundreds = Array("", "ÐIMTAS", "ÐIMTAI")
End Function
Private Function thousands() As Variant
    thousands = Array("TÛKSTANÈIØ", "TÛKSTANTIS", "TÛKSTANÈIAI")
End Function
Private Function millions() As Variant
    millions = Array("MILIJONØ", "MILIJONAS", "MILIJONAI")
End Function
Private Function currencies() As Variant
    currencies = Array("EURØ", "EURAS", "EURAI")
End Function

Function SkaiciusZodziais(NumberArg As Double) As String
    Dim parsedNumber As String
    Dim euros As String
    Dim result As String
     
    parsedNumber = Format(NumberArg, "000,000,000.00")
    If NumberArg < 1 Then
        result = "NULIS " + currencies()(0)
    Else
        result = Trim( _
            Resolve(Mid(parsedNumber, 1, 3), millions()) + " " + _
            Resolve(Mid(parsedNumber, 5, 3), thousands()) + " " + _
            Resolve(Mid(parsedNumber, 9, 3), currencies()) _
        )
    End If

    SkaiciusZodziais = UCase(Left(result, 1)) + LCase(Mid(result, 2)) + " ir " + Right(parsedNumber, 2) + " ct"
End Function

Private Function Resolve(numberPart As String, names As Variant)
    If numberPart <> "000" Then
        Resolve = NumberToText(numberPart) + " " + ResolvePluralForm(numberPart, names)
    Else
        Resolve = ""
    End If
End Function

Private Function ResolvePluralForm(numberPart, textValues As Variant)
    Dim tens
    Dim ones
    tens = Val(Mid(numberPart, 2, 1))
    ones = Val(Right(numberPart, 1))
    If tens = 1 Then
        ResolvePluralForm = textValues(0)
    Else
        Select Case ones
            Case 0
                ResolvePluralForm = textValues(0)
            Case 1
                ResolvePluralForm = textValues(1)
            Case Else
                ResolvePluralForm = textValues(2)
        End Select
    End If
End Function

Private Function NumberToText(number As String) As String
    Dim hundredsPart, tensPart, onesPart As Integer
    Dim hundredsText, tensText, onesText As String
    hundredsPart = Val(Left(number, 1))
    tensPart = Val(Mid(number, 2, 1))
    onesPart = Val(Right(number, 1))
    
    If hundredsPart > 0 Then
        hundredsText = ones()(IIf(hundredsPart = 1, 0, hundredsPart)) + " " + ResolvePluralForm(hundredsPart, hundreds)
    End If
    
    If tensPart = 1 Then
        tensText = teens()(onesPart)
    ElseIf tensPart > 1 Then
        tensText = tens()(tensPart)
    End If
    
    If tensPart <> 1 Then
        onesText = ones()(onesPart)
    End If

    NumberToText = Trim(hundredsText + " " + tensText + " " + onesText)
End Function






