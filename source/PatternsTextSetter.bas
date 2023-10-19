Attribute VB_Name = "PatternsTextSetter"
'===============================================================================
'   Макрос          : PatternsTextSetter
'   Версия          : 2023.10.19
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "PatternsTextSetter"

'===============================================================================

Private Type NodeData
    Node As Node
    Angle As Double
End Type

Private Const TEXT_OFFSET As Double = 1

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    ActiveDocument.Unit = cdrMillimeter
    
    Dim Text As String
    Text = VBA.InputBox("Введите номер", APP_NAME)
    If Text = vbNullString Then Exit Sub
    
    BoostStart APP_NAME, RELEASE
    
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    
    SetTextOnShapes Shapes, Text
    
    Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Sub SetTextOnShapes(ByVal Shapes As ShapeRange, ByVal Text As String)
    Dim Curve As Curve
    Dim Shape As Shape
    For Each Shape In Shapes
        If Shape.Type = cdrGroupShape Then
            Set Curve = GetCombinedCurve(Shape.Shapes.All)
        ElseIf Shape.Type = cdrCurveShape Then
            Set Curve = Shape.Curve
        End If
        SetTextOnCurve Curve, Text
    Next Shape
End Sub

Private Sub SetTextOnCurve(ByVal Curve As Curve, ByVal Text As String)
    Dim Node As NodeData
    Node = FindCornerNode(Curve)
    With ActiveLayer.CreateArtisticText(0, 0, Text)
        .SetSize , 4
        .LeftX = Node.Node.PositionX + TEXT_OFFSET
        .BottomY = Node.Node.PositionY + TEXT_OFFSET
        .RotationCenterX = .LeftX
        .RotationCenterY = .BottomY
        If Node.Angle < 90 Or Node.Angle > 270 Then
            .Rotate Node.Angle
        Else
            .Rotate Node.Angle - 180
        End If
    End With
End Sub

Function FindCornerNode(ByVal Curve As Curve) As NodeData
    Dim Segment As Segment
    Dim Offset As Double
    With Curve
        Set Segment = _
            .FindClosestSegment(.BoundingBox.Left, .BoundingBox.Bottom, Offset)
        Select Case Offset
            Case Is < 0.5
                Set FindCornerNode.Node = Segment.StartNode
                FindCornerNode.Angle = Segment.StartingControlPointAngle
            Case Is > 0.5
                Set FindCornerNode.Node = Segment.Next.StartNode
                FindCornerNode.Angle = Segment.Next.StartingControlPointAngle
        End Select
    End With
End Function


'===============================================================================
' # тесты

Private Sub testSomething()
'
End Sub
