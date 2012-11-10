'eRCode 0.2.0.17
'Copyright © 2000-2009 by Sergey N. Sitnic
'All rights reserved
Public Class eRCode
    Private Enum SymbolType As Integer
        LatinLowCase = 1 '97-122
        LatinUpperCase = 2 '65-90
        Digits = 3 '48-57
        LowLine = 4 '95
        CyrillicLowCase = 5 '224-255
        CyrillicUpperCase = 6 '192-223
    End Enum

    Public Enum AlgorithmTypes As Integer
        Standart = 0
        Advanced = 1
    End Enum

    Private UseLowLineValue As Boolean = False
    Private UseCyrillicValue As Boolean = False
    Private UseShuffleValue As Boolean = False
    Private ShuffleRoundsValue As Integer = 1
    Private LengthValue As Integer = 8
    Private LettersValue As Single = 0.7
    Private DigitsValue As Single = 0.3
    Private AlgorithmValue As AlgorithmTypes = AlgorithmTypes.Advanced


    Public Property Algorithm() As AlgorithmTypes
        Get
            Return AlgorithmValue
        End Get
        Set(ByVal value As AlgorithmTypes)
            AlgorithmValue = value
        End Set
    End Property

    Public Property UseLowline() As Boolean
        Get
            Return UseLowLineValue
        End Get
        Set(ByVal value As Boolean)
            UseLowLineValue = value
        End Set
    End Property

    Public Property UseCyrillic() As Boolean
        Get
            Return UseCyrillicValue
        End Get
        Set(ByVal value As Boolean)
            UseCyrillicValue = value
        End Set
    End Property

    Public Property UseShuffle() As Boolean
        Get
            Return UseShuffleValue
        End Get
        Set(ByVal value As Boolean)
            UseShuffleValue = value
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Return LengthValue
        End Get
        Set(ByVal value As Integer)
            LengthValue = value
            If LengthValue < 6 Then
                LengthValue = 6
            ElseIf LengthValue > 99 Then
                LengthValue = 99
            End If
        End Set
    End Property

    Public Property LettersPercent() As Single
        Get
            Return LettersValue
        End Get
        Set(ByVal value As Single)
            LettersValue = value
            If LettersValue < 0 Then
                LettersValue = 0
            ElseIf LettersValue > 1 Then
                LettersValue = 1
            End If
        End Set
    End Property

    Public Property DigitsPercent() As Single
        Get
            Return DigitsValue
        End Get
        Set(ByVal value As Single)
            DigitsValue = value
            If DigitsValue < 0 Then
                DigitsValue = 0
            ElseIf DigitsValue > 1 Then
                DigitsValue = 1
            End If
        End Set
    End Property

    Public Property ShuffleRounds() As Integer
        Get
            Return ShuffleRoundsValue
        End Get
        Set(ByVal value As Integer)
            ShuffleRoundsValue = value
        End Set
    End Property

    Public ReadOnly Property Result() As String
        Get
            If AlgorithmValue = AlgorithmTypes.Standart Then
                Return Standart()
            Else
                Return Advanced()
            End If
        End Get
    End Property

    Private Function Advanced() As String
        Dim Rnd As New Random
        Dim Count As Integer = LengthValue
        Dim Chars(Count - 1) As Char
        Dim Symbol As Integer
        Dim ILetter As Integer, IDigit As Integer, ILowline As Integer
        Dim [Char] As Integer
        Dim LettersCount As Integer = CInt(Math.Floor(LettersValue * LengthValue))
        Dim DigitsCount As Integer = CInt(Math.Floor(DigitsValue * LengthValue))
        Dim Sum As Integer = LettersCount + DigitsCount
        If Sum = LengthValue Then
            LettersCount -= Math.Abs(CInt(UseLowLineValue))
        ElseIf Sum > LengthValue Then
            LettersCount -= (Sum - LengthValue) + Math.Abs(CInt(UseLowLineValue))
        ElseIf Sum < LengthValue Then
            LettersCount += (LengthValue - Sum) - Math.Abs(CInt(UseLowLineValue))
        End If
        Dim Type As SymbolType
        Do Until Symbol = Count
            Type = CType(Rnd.Next(1, 7), SymbolType)
            Select Case Type
                Case SymbolType.LatinLowCase
                    If ILetter <> LettersCount Then
                        [Char] = Rnd.Next(97, 123)
                        ILetter += 1
                    End If
                Case SymbolType.LatinUpperCase
                    If ILetter <> LettersCount Then
                        [Char] = Rnd.Next(65, 91)
                        ILetter += 1
                    End If
                Case SymbolType.Digits
                    If IDigit <> DigitsCount Then
                        [Char] = Rnd.Next(48, 58)
                        IDigit += 1
                    End If
                Case SymbolType.LowLine
                    If UseLowLineValue Then
                        If ILowline <> 1 Then
                            [Char] = 95
                            ILowline += 1
                        End If
                    Else
                        If ILetter <> LettersCount Then
                            [Char] = Rnd.Next(97, 123)
                            ILetter += 1
                        End If
                    End If
                Case SymbolType.CyrillicLowCase
                    If UseCyrillicValue Then
                        If ILetter <> LettersCount Then
                            [Char] = Rnd.Next(224, 256)
                            ILetter += 1
                        End If
                    Else
                        If ILetter <> LettersCount Then
                            [Char] = Rnd.Next(97, 123)
                            ILetter += 1
                        End If
                    End If
                Case SymbolType.CyrillicUpperCase
                    If UseCyrillicValue Then
                        If ILetter <> LettersCount Then
                            [Char] = Rnd.Next(192, 224)
                            ILetter += 1
                        End If
                    Else
                        If ILetter <> LettersCount Then
                            [Char] = Rnd.Next(97, 123)
                            ILetter += 1
                        End If
                    End If
            End Select
            If [Char] <> 0 Then
                Chars(Symbol) = Chr([Char])
                Symbol += 1
                [Char] = 0
            End If
        Loop
        If UseShuffleValue Then
            For I As Integer = 1 To ShuffleRoundsValue
                Chars = Shuffle(Chars)
            Next
        End If
        Return Chars
    End Function

    Private Function Standart() As String
        Dim Rnd As New Random
        Dim Count As Integer = LengthValue
        Dim Chars(Count - 1) As Char
        Dim Symbol As Integer
        Dim [Char] As Integer, _flag As Integer
        Dim Type As SymbolType
        For Symbol = 0 To Count - 1
            Type = CType(Rnd.Next(1, 7), SymbolType)
            Select Case Type
                Case SymbolType.LatinLowCase
                    [Char] = Rnd.Next(97, 123)
                Case SymbolType.LatinUpperCase
                    [Char] = Rnd.Next(65, 91)
                Case SymbolType.Digits
                    [Char] = Rnd.Next(48, 58)
                Case SymbolType.LowLine
                    If UseLowLineValue Then
                        _flag = Rnd.Next(20, 50)
                        If _flag = 30 Then
                            [Char] = 95
                        Else
                            [Char] = Rnd.Next(97, 123)
                        End If
                    Else
                        [Char] = Rnd.Next(97, 123)
                    End If
                Case SymbolType.CyrillicLowCase
                    If UseCyrillicValue Then
                        [Char] = Rnd.Next(224, 256)
                    Else
                        [Char] = Rnd.Next(97, 123)
                    End If
                Case SymbolType.CyrillicUpperCase
                    If UseCyrillicValue Then
                        [Char] = Rnd.Next(192, 224)
                    Else
                        [Char] = Rnd.Next(97, 123)
                    End If
            End Select
            Chars(Symbol) = Chr([Char])
        Next
        If UseShuffleValue Then
            For I As Integer = 1 To ShuffleRoundsValue
                Chars = Shuffle(Chars)
            Next
        End If
        Return Chars
    End Function

    Private Function Shuffle(ByVal Chars As Char()) As Char()
        Dim Tmp As Char
        Dim RandomIndex As Integer
        Dim Rnd As New Random
        For I As Integer = 0 To Chars.Length - 1
            Tmp = Chars(I)
            RandomIndex = Rnd.Next(0, Chars.Length)
            Chars(I) = Chars(RandomIndex)
            Chars(RandomIndex) = Tmp
        Next
        Return Chars
    End Function
End Class
