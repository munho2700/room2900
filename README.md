# room2900

Sub InsertImagesAndToolboxesWithFormulas()
    Dim ws As Worksheet, wsLocation As Worksheet
    Dim shp As Shape, row As Range
    Dim imgFolder As String, imagePath As String
    Dim i As Long, startRow As Long, endRow As Long
    Dim searchValue As String

    ' 특정 시트("도면")와 이미지 폴더 경로 설정
    Set ws = ThisWorkbook.Sheets("도면") ' "도면" 시트를 명시적으로 설정
    Set wsLocation = ThisWorkbook.Sheets("LOCATION")
    imgFolder = "C:\Users\user\OneDrive\사진\" ' 이미지 폴더 경로

    ' 검색 값 가져오기 ('도면' 시트의 ao4 셀 값)
    searchValue = ThisWorkbook.Sheets("도면").Range("ao4").Value

    ' 기존 도형 삭제 (폼 컨트롤 및 OLE 객체 제외)
    For Each shp In ws.Shapes
        If Not (shp.Type = msoFormControl Or shp.Type = msoOLEControlObject) Then
            shp.Delete
        End If
    Next shp

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

    ' LOCATION 시트에서 위치 데이터와 수식 읽기
    For Each row In wsLocation.Range("A" & startRow & ":F" & endRow)
        ' 이미지 파일 경로 설정 (이미지 도형에 대한 처리)
        imagePath = imgFolder & row.Cells(1, 1).Text & ".png"
        If Dir(imagePath) <> "" Then
            ' 이미지 삽입
            Set shp = ws.Shapes.AddPicture(imagePath, False, True, _
                                           row.Cells(1, 2).Value, row.Cells(1, 3).Value, _
                                           row.Cells(1, 4).Value, row.Cells(1, 5).Value)
            shp.Name = row.Cells(1, 1).Text ' 도형 이름 설정
        End If

        ' 도형이 Rounded Rectangle일 경우 생성 및 수식 적용
        If InStr(row.Cells(1, 1).Text, "Rounded Rectangle") > 0 Then
            ' 둥근 모서리 사각형 도형 생성
            Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                         row.Cells(1, 2).Value, row.Cells(1, 3).Value, _
                                         row.Cells(1, 4).Value, row.Cells(1, 5).Value)
            With shp
                .Name = row.Cells(1, 1).Text
                .Fill.Visible = msoFalse
                .Line.Visible = msoTrue ' 테두리 없음
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Weight = 1

                ' 도형에 수식 또는 텍스트 적용 (F열에 저장된 수식/텍스트)
                '.TextFrame2.TextRange.Text = row.Cells(6).Text
                .TextFrame2.TextRange.Text = row.Cells(1, 6).Text

                ' 텍스트 서식 설정
                With .TextFrame2.TextRange.Font
                    .Name = "Arial"
                    .Size = 11
                    .Bold = msoTrue
                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
                End With
                .TextFrame2.HorizontalAnchor = msoAnchorCenter
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
            End With
        End If
    Next row

    MsgBox "모든 이미지와 도구상자가 지정된 위치에 삽입되었습니다."
End Sub






