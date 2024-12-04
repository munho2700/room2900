Sub InsertImagesAndToolboxesWithFormulas()
    Dim ws As Worksheet, wsLocation As Worksheet
    Dim shp As Shape, row As Range
    Dim imgFolder As String, imagePath As String
    Dim i As Long, startRow As Long, endRow As Long
    Dim searchValues() As Variant, toolboxValues As Variant
    Dim offsetY As Double, groupIndex As Long
    Dim groupShapes As Collection, groupShape As Shape
    Dim grp As ShapeRange
    Dim offsetX As Double, offsetYGroup As Double, customOffsetX As Double, customOffsetY As Double
    Dim shapeIndex As Long

    Set ws = ThisWorkbook.Sheets("도면")
    Set wsLocation = ThisWorkbook.Sheets("LOCATION")
    imgFolder = "C:\Users\user\OneDrive\사진\" ' 이미지 폴더 경로

    ' 검색 값 가져오기 ('도면' 시트의 Y4부터 빈 셀 전까지 값 가져오기)
    searchValues = ws.Range("Y4", ws.Cells(ws.Rows.Count, "Y").End(xlUp)).Value
    ' 도구상자 값 가져오기 ('도면' 시트의 AH4부터 빈 셀 전까지 값 가져오기)
    toolboxValues = ws.Range("AH4:AL" & ws.Cells(ws.Rows.Count, "AH").End(xlUp).row).Value

    ' 기존 도형 삭제 (폼 컨트롤 및 OLE 객체 제외)
    For Each shp In ws.Shapes
        If Not (shp.Type = msoFormControl Or shp.Type = msoOLEControlObject) Then
            shp.Delete
        End If
    Next shp

    ' 사용자 지정 X, Y 오프셋 설정
    customOffsetX = -60
    customOffsetY = -60

    ' 여러 그룹의 검색 값을 처리
    For groupIndex = LBound(searchValues) To UBound(searchValues)
        searchValue = searchValues(groupIndex, 1)
        offsetYGroup = customOffsetY + (groupIndex - 1) * 300 ' 각 그룹의 Y 좌표를 300 포인트 아래로 이동
        offsetX = customOffsetX + (groupIndex - 1) * 100 ' 각 그룹의 X 좌표를 100 포인트 오른쪽으로 이동

        ' LOCATION 시트에서 A 열에서 검색 값과 일치하는 시작 행 찾기
        startRow = 0
        For i = 2 To wsLocation.Cells(wsLocation.Rows.Count, "A").End(xlUp).row
            If wsLocation.Cells(i, "A").Value = searchValue Then
                startRow = i
                Exit For
            End If
        Next i

        ' 종료 행 찾기 ("Rounded Rectangle"이 끝나는 지점)
        If startRow > 0 Then
            For i = startRow + 1 To wsLocation.Cells(wsLocation.Rows.Count, "A").End(xlUp).row
                If Not InStr(wsLocation.Cells(i, "A").Value, "Rounded Rectangle") > 0 Then
                    endRow = i - 1
                    Exit For
                End If
            Next i
        End If

        ' 만약 끝나는 지점을 찾지 못하면 마지막 행까지 설정
        If endRow = 0 Then endRow = wsLocation.Cells(wsLocation.Rows.Count, "A").End(xlUp).row

        ' 그룹 도형들을 저장할 컬렉션 초기화
        Set groupShapes = New Collection

        shapeIndex = 1 ' 각 그룹 내의 이미지/도형에 대한 인덱스를 초기화

        ' LOCATION 시트에서 위치 데이터와 수식 읽기
        For Each row In wsLocation.Range("A" & startRow & ":F" & endRow)
            ' 이미지 파일 경로 설정 (이미지 도형에 대한 처리)
            imagePath = imgFolder & row.Cells(1, 1).Text & ".png"
            If Dir(imagePath) <> "" Then
                ' 각 개별 이미지에 대해 위치를 조정하는 추가 오프셋 설정
                Dim individualOffsetX As Double, individualOffsetY As Double
                individualOffsetX = 0
                individualOffsetY = 0

                ' 이미지 삽입
                Set shp = ws.Shapes.AddPicture(imagePath, False, True, _
                                               row.Cells(1, 2).Value + offsetX + individualOffsetX, _
                                               row.Cells(1, 3).Value + offsetYGroup + individualOffsetY, _
                                               row.Cells(1, 4).Value, row.Cells(1, 5).Value)
                shp.Name = row.Cells(1, 1).Text ' 도형 이름 설정
                groupShapes.Add shp ' 그룹에 추가
            End If

            ' 도구상자 삽입
            If groupIndex <= UBound(toolboxValues) Then
                Dim columnIndex As Long
                For columnIndex = 1 To 5 ' AH부터 AL까지의 값만 사용
                    ' 둥근 모서리 사각형 도형 생성
                    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                                 row.Cells(1, 2).Value + offsetX + (columnIndex - 1) * 120, _
                                                 row.Cells(1, 3).Value + offsetYGroup + 100, _
                                                 100, 50)
                    With shp
                        .Name = row.Cells(1, 1).Text & "_Toolbox_" & columnIndex
                        .Fill.Visible = msoFalse
                        .Line.Visible = msoTrue
                        .Line.ForeColor.RGB = RGB(0, 0, 0)
                        .Line.Weight = 1

                        ' 도구상자 텍스트 설정 (AH~AL 열에 직접 값을 넣기)
                        .TextFrame2.TextRange.Text = toolboxValues(groupIndex, columnIndex)

                        ' 텍스트 서식 설정
                        With .TextFrame2.TextRange.Font
                            .Name = "Arial"
                            .Size = 11
                            .Bold = msoTrue
                            .Fill.ForeColor.RGB = RGB(0, 0, 0)
                        End With
                        .TextFrame2.HorizontalAnchor = msoAnchorCenter
                        .TextFrame2.VerticalAnchor = msoAnchorMiddle

                        ' 도구상자 크기 고정 설정
                        shp.LockAspectRatio = msoFalse
                        shp.Width = 100
                        shp.Height = 50
                    End With
                    groupShapes.Add shp ' 그룹에 추가
                Next columnIndex
            End If

            shapeIndex = shapeIndex + 1 ' 다음 도형/이미지를 위해 인덱스 증가
        Next row

        ' 그룹화하여 레이블 지정 및 위치 조정
        If groupShapes.Count > 0 Then
            Dim shapeArray() As Variant
            ReDim shapeArray(1 To groupShapes.Count)
            For i = 1 To groupShapes.Count
                shapeArray(i) = groupShapes(i).Name
            Next i
            On Error Resume Next
            Set grp = ws.Shapes.Range(shapeArray).Group
            On Error GoTo 0
            If Not grp Is Nothing Then
                grp.Name = "Group" & groupIndex ' 그룹에 이름 부여

                ' 그룹의 위치 조정 (2번째 그룹은 추가적으로 X, Y 방향 이동)
                If groupIndex = 2 Then
                    grp.IncrementLeft 100
                    grp.IncrementTop 100
                End If
            End If
        End If
    Next groupIndex

    MsgBox "모든 이미지와 도구상자가 지정된 위치에 삽입되었습니다."
End Sub


