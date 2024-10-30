Sub Box_Scale()
    ' SEOBLI :)
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    ' Set reference point to center
    ActiveDocument.ReferencePoint = cdrCenter
    
    ' Get the current dimensions
    Dim currentWidth As Double
    Dim currentHeight As Double
    Dim compWidth As Double
    Dim compHeight As Double
    Dim aspectRatio As Double
    
    currentWidth = OrigSelection.SizeWidth
    currentHeight = OrigSelection.SizeHeight
    
    ' Calculate aspect ratio
    aspectRatio = currentWidth / currentHeight
    
    ' Check conditions and resize while maintaining the aspect ratio
    If currentWidth > currentHeight Then
        compWidth = 7.480315 * aspectRatio
        If compWidth < 3.543307 Then
            OrigSelection.SetSize 3.543307, 3.543307 / aspectRatio
        Else
            OrigSelection.SetSize 7.480315 * aspectRatio, 7.480315
        End If
    ElseIf currentWidth < currentHeight Then
        compHeight = 3.543307 / aspectRatio
         If compHeight < 7.480315 Then
            OrigSelection.SetSize 7.480315 * aspectRatio, 7.480315
        Else
            OrigSelection.SetSize 3.543307, 3.543307 / aspectRatio
        End If
    End If
End Sub
Sub Volumetric_Box_Scale()
    ' SEOBLI :)
    Dim doc As Document
    Dim pg As Page
    Dim bfo As Shape
    Dim currentWidth As Double
    Dim currentHeight As Double
    Dim compWidth As Double
    Dim compHeight As Double
    Dim aspectRatio As Double
    
    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    doc.ReferencePoint = cdrCenter
    
    For Each bfo In pg.Shapes
        If bfo.Type = cdrBitmapShape Then
            currentWidth = bfo.SizeWidth
            currentHeight = bfo.SizeHeight
            aspectRatio = currentWidth / currentHeight
            If currentWidth > currentHeight Then
                compWidth = 7.480315 * aspectRatio
                If compWidth < 3.543307 Then
                    bfo.SetSize 3.543307, 3.543307 / aspectRatio
                Else
                    bfo.SetSize 7.480315 * aspectRatio, 7.480315
                End If
            ElseIf currentWidth < currentHeight Then
                compHeight = 3.543307 / aspectRatio
                 If compHeight < 7.480315 Then
                    bfo.SetSize 7.480315 * aspectRatio, 7.480315
                Else
                    bfo.SetSize 3.543307, 3.543307 / aspectRatio
                End If
            End If
        End If
    Next bfo
End Sub

Sub Seobli_Adaptive_Clip()
    ' SEOBLI :)
    Dim doc As Document
    Dim pg As Page
    Dim bfo As Shape
    Dim rect As Shape
    Dim currentWidth As Double
    Dim currentHeight As Double
    Dim compWidth As Double
    Dim compHeight As Double
    Dim aspectRatio As Double
    Dim rectQueue As Collection
    Dim accumulationID As Integer
    Dim pclipContainer As Shape
    
    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    doc.ReferencePoint = cdrCenter
    Set rectQueue = New Collection
    
    For Each rect In pg.Shapes
        If rect.Type = cdrRectangleShape Then
            rectQueue.Add rect
        End If
    Next rect
    
    accumulationID = 1
    
    For Each bfo In pg.Shapes
        If bfo.Type = cdrBitmapShape Then
            currentWidth = bfo.SizeWidth
            currentHeight = bfo.SizeHeight
            aspectRatio = currentWidth / currentHeight
            
            If currentWidth > currentHeight Then
                compWidth = 7.480315 * aspectRatio
                If compWidth < 3.543307 Then
                    bfo.SetSize 3.543307, 3.543307 / aspectRatio
                Else
                    bfo.SetSize 7.480315 * aspectRatio, 7.480315
                End If
            Else
                compHeight = 3.543307 / aspectRatio
                If compHeight < 7.480315 Then
                    bfo.SetSize 7.480315 * aspectRatio, 7.480315
                Else
                    bfo.SetSize 3.543307, 3.543307 / aspectRatio
                End If
            End If
            
            If accumulationID <= rectQueue.Count Then
                Set rect = rectQueue(accumulationID)
                
                If Not rect.PowerClip Is Nothing Then
                    bfo.AddToPowerClip rect, True
                Else
                    MsgBox "Selected Container is not a valid PowerClip Container."
                End If
                
                accumulationID = accumulationID + 1
            End If
        End If
    Next bfo
End Sub
