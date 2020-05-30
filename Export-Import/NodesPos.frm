VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NodesPos 
   Caption         =   "NodesPosition v2"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   OleObjectBlob   =   "NodesPos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NodesPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CmbBottom_Click()
    Dim s As Shape
    Dim n As Node
    Dim x As Double, y As Double
    Dim tmp As Double
    Dim first As Boolean
    
    ActiveDocument.BeginCommandGroup "NodePosition"
    
    first = True
    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.GetPosition x, y
            If first Then
                tmp = y
                first = False
            Else
                If tmp > y Then tmp = y
            End If
        Next n
    Next s

    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.PositionY = tmp
        Next n
    Next s
    
    ActiveDocument.EndCommandGroup
    End
End Sub

Private Sub CmbLeft_Click()
    Dim s As Shape
    Dim n As Node
    Dim x As Double, y As Double
    Dim tmp As Double
    Dim first As Boolean
        
    ActiveDocument.BeginCommandGroup "NodePosition"
    
    first = True
    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.GetPosition x, y
            If first Then
                tmp = x
                first = False
            Else
                If tmp > x Then tmp = x
            End If
        Next n
    Next s

    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.PositionX = tmp
        Next n
    Next s
    
    ActiveDocument.EndCommandGroup
    End
End Sub

Private Sub CmbRight_Click()
    Dim s As Shape
    Dim n As Node
    Dim x As Double, y As Double
    Dim tmp As Double
    Dim first As Boolean
    
    ActiveDocument.BeginCommandGroup "NodePosition"
    
    first = True
    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.GetPosition x, y
            If first Then
                tmp = x
                first = False
            Else
                If tmp < x Then tmp = x
            End If
        Next n
    Next s

    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.PositionX = tmp
        Next n
    Next s
    
    ActiveDocument.EndCommandGroup
    End
End Sub

Private Sub CmbTop_Click()
    Dim s As Shape
    Dim n As Node
    Dim x As Double, y As Double
    Dim tmp As Double
    Dim first As Boolean
    
    ActiveDocument.BeginCommandGroup "NodePosition"
    
    first = True
    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.GetPosition x, y
            If first Then
                tmp = y
                first = False
            Else
                If tmp < y Then tmp = y
            End If
        Next n
    Next s

    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.PositionY = tmp
        Next n
    Next s
    
    ActiveDocument.EndCommandGroup
    End
End Sub

Private Sub CommandButton1_Click()
    Dim n As Node
    Dim x As Double, y As Double
    Dim chX As Double, chY As Double
    Dim stri As String
    Dim s As Shape
    Dim stateX As String, stateY As String
    ' "+-/*" - прибавляем к каждой координате
    ' "n" - полученное число заменяет текущую коор.
    ' "s" - ничего не меняем, пропускаем
    ' "e" - ничего не делаем, ошибка
    
    ActiveDocument.Unit = cdrMillimeter
    
    ActiveDocument.BeginCommandGroup "NodePosition"
    
    'сначала определяем что написал юзер
    stri = Trim(nPositionX.Text)
    If stri <> "" Then
        stri = comma2dot(stri)
        str2number stri, chX, stateX
    Else
        stateX = "s"
    End If
    
    stri = Trim(nPositionY.Text)
    If stri <> "" Then
        stri = comma2dot(stri)
        str2number stri, chY, stateY
    Else
        stateY = "s"
    End If
    
    For Each s In ActiveSelection.Shapes
        For Each n In s.Curve.Selection
            n.GetPosition x, y
            Select Case stateX
                Case "+", "-", "*", "/"
                    n.PositionX = firstsign2number(stateX, x, chX)
                Case "n"
                    n.PositionX = chX
                Case "s"
                    'n.PositionX = x
                Case "e"
                    nPositionX = ""
                    Exit Sub
            End Select

            Select Case stateY
                Case "+", "-", "*", "/"
                    n.PositionY = firstsign2number(stateY, Round(y, 3), chY)
                Case "n"
                    n.PositionY = chY
                Case "s"
                    'n.PositionY = y
                Case "e"
                    nPositionY = ""
                    Exit Sub
            End Select
        Next n
    Next s
    
    ActiveDocument.EndCommandGroup
    End
End Sub

Private Sub str2number(ByVal stri$, ByRef ch As Double, ByRef st$)
                          'crch - current node position X|Y
    Dim i As Integer
    Dim s As String
    
    Dim dgchar As Byte
    Dim mathchar As Byte
    Dim dotchar As Byte
    Dim spacechar As Byte
    Dim elsechar As Byte
    
    Dim frs As String 'first letter in the stri
    
    'Анализ
    For i = 1 To Len(stri)
        s = Mid(stri, i, 1)
        Select Case s
            Case 0 To 9
                dgchar = dgchar + 1
            Case "+", "-", "*", "/"
                mathchar = mathchar + 1
            Case "."
                dotchar = dotchar + 1
            Case " "
                spacechar = spacechar + 1
            Case Else
                elsechar = elsechar + 1
                Exit For
        
        End Select
    Next
    
    If elsechar > 0 Or dotchar > 2 Or mathchar > 1 Then
        st = "e" 'error
        Exit Sub
    End If
    
    If spacechar > 0 And mathchar = 0 Then
        st = "e" 'error, не может быть 345 546
        Exit Sub
    End If
    
    frs = Left(stri, 1)
   
     Select Case frs
        Case "+", "-", "*", "/"
            If dotchar > 1 Then 'не может быть две точки, если знак вначале +1.5
                st = "e" 'error
                Exit Sub
            End If
            st = frs
            ch = Val(Mid(stri, 2))
        Case 0 To 9
            If mathchar > 0 Then 'мат-оператор в середине (если в начале, то проверено выше)
                ch = concat(stri)
                st = "n"
            Else
                ch = Val(stri)
                st = "n"
            End If
        Case "."
            If mathchar > 0 Then 'всё то же самое, что в предыдущем кейсе
                ch = concat(stri)
                st = "n"
            Else
                ch = Val(stri)
                st = "n"
            End If
        Case Else
            st = "e"
    End Select
    
    
End Sub

Private Function concat(stri$) As Double
    Dim opt(3) As String
    Dim ch As Double
    Dim i As Integer, j As Integer
    
    opt(0) = "+": opt(1) = "-"
    opt(2) = "*": opt(3) = "/"
    
    For i = 0 To 3
        j = InStr(stri, opt(i))
        If j <> 0 Then
            Select Case opt(i)
                Case "+"
                    ch = Val(Mid(stri, 1, j - 1)) + Val(Mid(stri, j + 1))
                Case "-"
                    ch = Val(Mid(stri, 1, j - 1)) - Val(Mid(stri, j + 1))
                Case "*"
                    ch = Val(Mid(stri, 1, j - 1)) * Val(Mid(stri, j + 1))
                Case "/"
                    ch = Val(Mid(stri, 1, j - 1)) / Val(Mid(stri, j + 1))
            End Select
            Exit For
        Else
        ch = Val(stri)
        End If
    Next i
    Erase opt
    concat = ch
End Function

Private Function comma2dot(s$) As String
    s = Replace(s, ",", ".")
    comma2dot = s
End Function

Private Function firstsign2number(sn$, crch#, ch#) As Double
    Dim temp As Double
    Select Case sn
        Case "+"
            temp = crch + ch
        Case "-"
            temp = crch - ch
        Case "*"
            temp = crch * ch
        Case "/"
            temp = crch / ch
    End Select
    firstsign2number = temp
End Function

Private Sub CommandButton2_Click()
    End
End Sub
