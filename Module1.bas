Attribute VB_Name = "Module1"

Option Explicit

'Public TerrainSTOP As Boolean
Private Type FHbmp_type 'bmp File Header structure
    BF As String * 2 'field 1 is always "BF"
    filesize As Long 'file size in bytes
    h0 As Long 'reserved fields, always zero
    OffSet As Long 'offset to start of color index array.
End Type
Private Type IHbmp_type 'bmp Information header structure
    hHdrSize As Long 'size of header, always h28 for our purposes
    BitWidth As Long 'bitmap width
    BitHeight As Long 'bitmap height
    Planes As Integer 'planes, always 1 for .bmp
    BitCnt As Integer 'bit count, lenght of RGBQuad address, set to h18
    comp As Long 'compression, set to 0
    ImageSize As Long 'image size
    XpM As Long 'X pixels per meter, set to 0
    YpM As Long 'Y pixels per meter, set to 0
    ClrUsed As Long 'Colors used, 0 for this case
    ClrImp As Long 'colors important
End Type
'a three-byte structure for storing RGB info in the Bitmap
Private Type ColorIndex_Type
    iBlue As Byte
    iGreen As Byte
    iRed As Byte
End Type

Private Const BI_RGB As Integer = 0
Private Const BI_BITFIELDS As Integer = 3
Public ColorObject As New CScale

Public Sub MakeExposureBMP(iPixH As Long, iPixW As Long, _
                           filespec As String, iexp() As Double, _
                           ExpMin As Long, ExpMax As Long)
    
    Dim ColorIndex() As ColorIndex_Type  'bitmap array for .bmp files
    Dim FHbmp As FHbmp_type 'File header
    Dim IHbmp As IHbmp_type 'Info header
    Dim i As Integer, j As Integer, k As Integer
    Dim iFile As Integer
    Dim Pad_Zeros As Integer
    Dim SrcW As Long, SrcH As Long
    Dim SrcWL As Long, SrcWH As Long
    Dim SrcHL As Long, SrcHH As Long
    
    Dim rpart As Double, gpart As Double
    
    'ColorObject.Init ExpMin, ExpMax, RGB(0, 0, 64)
    
    'first the headers:
    With IHbmp
        .BitHeight = iPixH
        .BitWidth = iPixW
        .comp = BI_RGB
        .XpM = 0
        .YpM = 0
        .ClrImp = 0
        .Planes = 1
        .BitCnt = 24
        .ClrUsed = 0
        .hHdrSize = 40
        'the image size is the bitHeight*Width*Count/8 PLUS
        'the padding to make each scan line end with a complete DWord.
        .ImageSize = (.BitHeight * .BitWidth * .BitCnt / 8) + ((.BitWidth Mod 4) * .BitHeight)
    End With
    With FHbmp
        .BF = "BM"
        .filesize = IHbmp.ImageSize + 54
        .h0 = 0
        .OffSet = 54
    End With
    'location of the Source and Width/Height of the "+"
    SrcW = iPixW / 2
    SrcH = iPixH / 2
    SrcWL = SrcW - (0.03 * iPixW)
    SrcWH = SrcW + (0.03 * iPixW)
    SrcHL = SrcH - (0.03 * iPixH)
    SrcHH = SrcH + (0.03 * iPixH)
    
    Screen.MousePointer = vbHourglass
        'create an array of color byte-triplets, ColorIndex().
        'each byte in the triplet represents the intensity
        'of Red, Green or Blue for the current pixel.
        'Center the array on SrcW, SrcH.
        'The array should be iPixH points (pixels) high by iPixW points wide.
        'plus a padding of up to 3-bytes of zeros on the end of each line to
        'make each line-length a multiple of 4 bytes.
        ReDim ColorIndex(1 To iPixW, 1 To iPixH)
        For i = 1 To iPixH
            For j = 1 To iPixW
                If i = SrcH And j > SrcWL And j < SrcWH Or _
                   j = SrcW And i > SrcHL And i < SrcHH Then
                   'Center Cross Hair
                    ColorIndex(j, i).iRed = 0
                    ColorIndex(j, i).iGreen = 0
                    ColorIndex(j, i).iBlue = 0
'                ElseIf i Mod 90 = 0 Or _
'                       j Mod 90 = 0 Then
'                       'Grid Lines
'                    ColorIndex(j, i).iRed = 192
'                    ColorIndex(j, i).iGreen = 192
'                    ColorIndex(j, i).iBlue = 192
                Else 'get a color for this exposure level
                
                
                'five fixed color groups
                    Select Case iexp(j, i)
                    Case Is < 0.5
                        ColorIndex(j, i).iRed = 255
                        ColorIndex(j, i).iGreen = 255
                        ColorIndex(j, i).iBlue = 255
                    Case 0.5 To 1#
                        ColorIndex(j, i).iRed = 0
                        ColorIndex(j, i).iGreen = 225
                        ColorIndex(j, i).iBlue = 0
                    Case 1# To 1.5
                        ColorIndex(j, i).iRed = 255
                        ColorIndex(j, i).iGreen = 255
                        ColorIndex(j, i).iBlue = 0
                    Case 1.5 To 2#
                        ColorIndex(j, i).iRed = 0
                        ColorIndex(j, i).iGreen = 255
                        ColorIndex(j, i).iBlue = 255
                    Case 2# To 2.5
                        ColorIndex(j, i).iRed = 175
                        ColorIndex(j, i).iGreen = 0
                        ColorIndex(j, i).iBlue = 200
                    Case Is > 2.5
                        ColorIndex(j, i).iRed = 255
                        ColorIndex(j, i).iGreen = 0
                        ColorIndex(j, i).iBlue = 0
                    End Select
                    
                    
'                'green-yellow-red gradiant
'                rpart = iexp(j, i) * 25.5
'                If rpart > 255 Then rpart = 255
'                gpart = 255 - rpart
'
'                ColorIndex(j, i).iRed = rpart
'                ColorIndex(j, i).iGreen = gpart
'                ColorIndex(j, i).iBlue = 0
                    
'                'green-yellow-red gradiant
'                rpart = iexp(j, i) * 25.5 / 8
'                If rpart > 255 Then rpart = 255
'                gpart = 255 - rpart
'
'                ColorIndex(j, i).iRed = rpart
'                ColorIndex(j, i).iGreen = gpart
'                ColorIndex(j, i).iBlue = 0
                    
                    
'                'green-yellow gradiant.  Red over 95%
'                ColorIndex(j, i).iBlue = 32
'
'                Select Case iexp(j, i)
'                    Case Is < 55
'                        rpart = iexp(j, i) * 25.5 / 4
'                        If rpart > 255 Then rpart = 255
'                        ColorIndex(j, i).iRed = rpart
'                        ColorIndex(j, i).iGreen = 255
'
'                    Case Else
'                        ColorIndex(j, i).iRed = 255
'
'                        gpart = 255 - ((iexp(j, i) - 50) * 8.5)
'                        If gpart < 0 Then gpart = 0
'                        ColorIndex(j, i).iGreen = gpart
'
'                End Select
                    
'                    ColorIndex(j, i).iRed = ColorObject.getRed(iexp(j, i))
'                    ColorIndex(j, i).iGreen = ColorObject.getGreen(iexp(j, i))
'                    ColorIndex(j, i).iBlue = ColorObject.getBlue(iexp(j, i))
                End If
            Next j
        Next i
        
        'write the data to a file
        iFile = FreeFile
    
        Open filespec For Binary Access Write As iFile
            Put iFile, , FHbmp
            Put iFile, , IHbmp
            
            Pad_Zeros = iPixW Mod 4
            'Dim Buffer() As ColorIndex_Type
            Dim Buffer() As Byte
            'ReDim Buffer(1 To iPixW) As ColorIndex_Type
            ReDim Buffer(1 To iPixW * 3 + Pad_Zeros) As Byte
            For i = 1 To iPixH
                For j = 1 To iPixW
                    'Buffer(j) = ColorIndex(j, i)
                    Buffer((j - 1) * 3 + 1) = ColorIndex(j, i).iBlue
                    Buffer((j - 1) * 3 + 2) = ColorIndex(j, i).iGreen
                    Buffer((j - 1) * 3 + 3) = ColorIndex(j, i).iRed
                    'Put iFile, , ColorIndex(j, i)
                Next j
                'pad the end of each scan line with zero's to fill out the last DWord
                'For k = 1 To Pad_Zeros
                '    Buffer(iPixW * 3 + k) = CByte(0)
                '    'Put iFile, , CByte(0)
                'Next k
                Put iFile, , Buffer
            Next i
        Close iFile
    
    Screen.MousePointer = vbNormal



End Sub

Public Sub MakeBMPKey(filespec As String, CRange As Integer)
    
    Dim ColorIndex() As ColorIndex_Type  'bitmap array for .bmp files
    Dim FHbmp As FHbmp_type 'File header
    Dim IHbmp As IHbmp_type 'Info header
    Dim i As Integer, j As Integer, k As Integer
    Dim iFile As Integer
    Dim Pad_Zeros As Integer
    Dim SrcW As Long, SrcH As Long
    Dim SrcWL As Long, SrcWH As Long
    Dim SrcHL As Long, SrcHH As Long
    Dim rpart As Double, gpart As Double, bpart As Double
    
    'ColorObject.Init ExpMin, ExpMax, RGB(0, 0, 64)
    
    'first the headers:
    With IHbmp
        .BitHeight = 720
        .BitWidth = 100
        .comp = BI_RGB
        .XpM = 0
        .YpM = 0
        .ClrImp = 0
        .Planes = 1
        .BitCnt = 24
        .ClrUsed = 0
        .hHdrSize = 40
        'the image size is the bitHeight*Width*Count/8 PLUS
        'the padding to make each scan line end with a complete DWord.
        .ImageSize = (.BitHeight * .BitWidth * .BitCnt / 8) + ((.BitWidth Mod 4) * .BitHeight)
    End With
    With FHbmp
        .BF = "BM"
        .filesize = IHbmp.ImageSize + 54
        .h0 = 0
        .OffSet = 54
    End With
    
    Screen.MousePointer = vbHourglass
        'create an array of color byte-triplets, ColorIndex().
        'each byte in the triplet represents the intensity
        'of Red, Green or Blue for the current pixel.
        'Center the array on SrcW, SrcH.
        'The array should be iPixH points (pixels) high by iPixW points wide.
        'plus a padding of up to 3-bytes of zeros on the end of each line to
        'make each line-length a multiple of 4 bytes.
        ReDim ColorIndex(1 To 100, 1 To 720)
        
        Dim vpct As Double
        
        For i = 1 To 720
            vpct = i / 240
            
            If i Mod 60 = 0 Then
                       'Grid Lines
                    rpart = 0
                    gpart = 0
                    bpart = 0
            Else
                'five fixed color groups
                    Select Case vpct
                    Case Is < 0.5
                        rpart = 255
                        gpart = 255
                        bpart = 255
                    Case 0.5 To 1
                        rpart = 0
                        gpart = 225
                        bpart = 0
                    Case 1 To 1.5
                        rpart = 255
                        gpart = 255
                        bpart = 0
                    Case 1.5 To 2
                        rpart = 0
                        gpart = 255
                        bpart = 255
                    Case 2 To 2.5
                        rpart = 175
                        gpart = 0
                        bpart = 200
                    Case Is > 2.5
                        rpart = 255
                        gpart = 0
                        bpart = 0
                    End Select
'                bpart = 32
'
'                Select Case vpct
'                    Case Is < 25
'                        rpart = 175
'                        gpart = 175
'                        bpart = 175
'                    Case Is < 40
'                        rpart = 0
'                        gpart = 255
'                    Case Is < 90
'                        rpart = (vpct - 40) * 7
'                        If rpart > 255 Then rpart = 255
'                        gpart = 255
'                    Case Else
'                        rpart = 255
'                        gpart = 255 - ((vpct - 90) * 20)
'                        If gpart < 0 Then gpart = 0
'                End Select
            End If
            
            For j = 1 To 100
                ColorIndex(j, i).iBlue = bpart
                ColorIndex(j, i).iRed = rpart
                ColorIndex(j, i).iGreen = gpart
            Next j
        Next i
        
        'write the data to a file
        iFile = FreeFile
    
        Open filespec For Binary Access Write As iFile
            Put iFile, , FHbmp
            Put iFile, , IHbmp
            
            Pad_Zeros = 100 Mod 4
            'Dim Buffer() As ColorIndex_Type
            Dim Buffer() As Byte
            'ReDim Buffer(1 To iPixW) As ColorIndex_Type
            ReDim Buffer(1 To 100 * 3 + Pad_Zeros) As Byte
            For i = 1 To 720
                For j = 1 To 100
                    'Buffer(j) = ColorIndex(j, i)
                    Buffer((j - 1) * 3 + 1) = ColorIndex(j, i).iBlue
                    Buffer((j - 1) * 3 + 2) = ColorIndex(j, i).iGreen
                    Buffer((j - 1) * 3 + 3) = ColorIndex(j, i).iRed
                    'Put iFile, , ColorIndex(j, i)
                Next j
                'pad the end of each scan line with zero's to fill out the last DWord
                'For k = 1 To Pad_Zeros
                '    Buffer(iPixW * 3 + k) = CByte(0)
                '    'Put iFile, , CByte(0)
                'Next k
                Put iFile, , Buffer
            Next i
        Close iFile
    Screen.MousePointer = vbNormal
End Sub

Public Sub ReplaceColorInBMP(SourceBMP As String, DestBMP As String, OldColor As Long, NewColor As Long)
    'replace OldColor with NewColor in SourceBMP and save it to DestBMP
    
    Dim FHbmp As FHbmp_type 'File header
    Dim IHbmp As IHbmp_type 'Info header
    Dim i As Long, j As Long
    Dim iFile As Integer, jFile As Integer
    Dim BinsHigh As Long, BinsWide As Long
    Dim Buffer() As Byte 'ColorIndex_Type
    Dim OC As ColorIndex_Type
    Dim NC As ColorIndex_Type
    Dim Bufx As ColorIndex_Type
    Dim Pad_Zeros As Integer
    Dim LineWidth As Long
    
    If gFSO.FileExists(SourceBMP) Then
        Screen.MousePointer = vbHourglass
        GetRGB OldColor, OC.iRed, OC.iGreen, OC.iBlue
        GetRGB NewColor, NC.iRed, NC.iGreen, NC.iBlue
        iFile = FreeFile
        Open SourceBMP For Binary Access Read As iFile
            'get the headers
            Get iFile, , FHbmp
            Get iFile, , IHbmp
            'bmp dims
            BinsHigh = IHbmp.BitHeight
            BinsWide = IHbmp.BitWidth
            Pad_Zeros = BinsWide Mod 4
            'open the destination file
            jFile = FreeFile
            Open DestBMP For Binary Access Write As jFile
                'put the headers
                Put jFile, , FHbmp
                Put jFile, , IHbmp
                'dim Buffer() for one line of data
                LineWidth = BinsWide * 3 + Pad_Zeros
                ReDim Buffer(1 To LineWidth)
                'get data, swap old color for new, one line at a time
                For i = 1 To BinsHigh
                    'get a line of data
                    Get iFile, , Buffer
                    'check each color triplet for a match to Old Color
                    For j = 1 To LineWidth - 2 Step 3
                        If (Buffer(j) = OC.iBlue) And _
                        (Buffer(j + 1) = OC.iGreen) And _
                        (Buffer(j + 2) = OC.iRed) Then
                           'if a match, replace Old Color with New Color
                            Buffer(j) = NC.iBlue
                            Buffer(j + 1) = NC.iGreen
                            Buffer(j + 2) = NC.iRed
                        End If
                        'If j > LineWidth - 10 Then Stop
                    Next j
                    'write data to the new file
                    Put jFile, , Buffer
                Next i 'next line of data
            Close jFile
        Close iFile
        Screen.MousePointer = vbNormal
    End If

End Sub

Public Sub GetRGB(Composite As Long, ColorR As Byte, ColorG As Byte, ColorB As Byte)
    
'given a long containing a color code, return the rgb byte values
    ColorR = Composite And &H10000FE
    ColorG = (Composite And &H100FE00) / &H100
    ColorB = (Composite And &H1FE0000) / &H10000

End Sub

Public Sub MakeMultiExposureBMP(iPixH As Long, iPixW As Long, _
                           filespec As String, iexp() As Double, CRange As Integer)
    
    Dim ColorIndex() As ColorIndex_Type  'bitmap array for .bmp files
    Dim FHbmp As FHbmp_type 'File header
    Dim IHbmp As IHbmp_type 'Info header
    Dim i As Integer, j As Integer, k As Integer
    Dim iFile As Integer
    Dim Pad_Zeros As Integer
    Dim SrcW As Long, SrcH As Long
    Dim SrcWL As Long, SrcWH As Long
    Dim SrcHL As Long, SrcHH As Long
    Dim rpart As Double, gpart As Double, bpart As Double, vpct As Double
    Dim tReport As String, TS As TextStream
    Dim MaxH As Long, MaxW As Long, writeline As Boolean
    Dim MaxVal As Double, TxtLine As String
    
    With gFSO
        tReport = .BuildPath(.GetParentFolderName(filespec), .GetBaseName(filespec) & ".csv")
    End With
    Set TS = gFSO.CreateTextFile(tReport, True)
    
    TxtLine = " ,"
    For j = 1 To iPixW Step 10
        TxtLine = TxtLine & CStr(j / 10) & ", "
    Next j
    TS.writeline TxtLine
    
    'first the headers:
    With IHbmp
        .BitHeight = iPixH
        .BitWidth = iPixW
        .comp = BI_RGB
        .XpM = 0
        .YpM = 0
        .ClrImp = 0
        .Planes = 1
        .BitCnt = 24
        .ClrUsed = 0
        .hHdrSize = 40
        'the image size is the bitHeight*Width*Count/8 PLUS
        'the padding to make each scan line end with a complete DWord.
        .ImageSize = (.BitHeight * .BitWidth * .BitCnt / 8) + ((.BitWidth Mod 4) * .BitHeight)
    End With
    With FHbmp
        .BF = "BM"
        .filesize = IHbmp.ImageSize + 54
        .h0 = 0
        .OffSet = 54
    End With
    
    Screen.MousePointer = vbHourglass
        'create an array of color byte-triplets, ColorIndex().
        'each byte in the triplet represents the intensity
        'of Red, Green or Blue for the current pixel.
        'Center the array on SrcW, SrcH.
        'The array should be iPixH points (pixels) high by iPixW points wide.
        'plus a padding of up to 3-bytes of zeros on the end of each line to
        'make each line-length a multiple of 4 bytes.
        ReDim ColorIndex(1 To iPixW, 1 To iPixH)
        MaxVal = -999
        For i = 1 To iPixH
            writeline = i Mod 10 = 0
            If writeline Then TxtLine = i / 10 & ", "
            For j = 1 To iPixW
                vpct = iexp(j, i)
                If vpct > MaxVal Then
                    'new peak value
                    MaxVal = vpct
                    MaxH = i
                    MaxW = j
                End If
                If writeline Then
                    If j Mod 10 = 0 Then
                        TxtLine = TxtLine & vpct & ", "
                    End If
                End If
                If vpct = -999 Then
                   'source center
                    rpart = 0
                    gpart = 0
                    bpart = 0
                ElseIf i Mod 100 = 0 Or _
                       j Mod 100 = 0 Then
                       'Grid Lines
                    rpart = 5
                    gpart = 5
                    bpart = 5
                Else 'get a color for this exposure level
                    'green-yellow gradiant.  Red over 95%
                    'bpart = 32
                    
                    
                    'five fixed color groups
                    Select Case CRange
                    
                    Case 0 '1 to 5%
                        Select Case iexp(j, i)
                        Case Is < 1
                            rpart = 255
                            gpart = 255
                            bpart = 255
                        Case 1 To 2
                            rpart = 0
                            gpart = 225
                            bpart = 0
                        Case 2 To 3
                            rpart = 255
                            gpart = 255
                            bpart = 0
                        Case 3 To 4
                            rpart = 0
                            gpart = 255
                            bpart = 255
                        Case 4 To 5
                            rpart = 175
                            gpart = 0
                            bpart = 200
                        Case Is > 5
                            rpart = 255
                            gpart = 0
                            bpart = 0
                        End Select
                    Case 1 '2 to 12%
                        'five fixed color groups
                        Select Case iexp(j, i)
                        Case Is < 2.5
                            rpart = 255
                            gpart = 255
                            bpart = 255
                        Case 2.5 To 5
                            rpart = 0
                            gpart = 225
                            bpart = 0
                        Case 5 To 7.5
                            rpart = 255
                            gpart = 255
                            bpart = 0
                        Case 7.5 To 10
                            rpart = 0
                            gpart = 255
                            bpart = 255
                        Case 10 To 12.5
                            rpart = 175
                            gpart = 0
                            bpart = 200
                        Case Is > 12.5
                            rpart = 255
                            gpart = 0
                            bpart = 0
                        End Select
                    Case 2 '5 to 25%
                        'five fixed color groups
                        Select Case iexp(j, i)
                        Case Is < 5
                            rpart = 255
                            gpart = 255
                            bpart = 255
                        Case 5 To 10
                            rpart = 0
                            gpart = 225
                            bpart = 0
                        Case 10 To 15
                            rpart = 255
                            gpart = 255
                            bpart = 0
                        Case 15 To 20
                            rpart = 0
                            gpart = 255
                            bpart = 255
                        Case 20 To 25
                            rpart = 175
                            gpart = 0
                            bpart = 200
                        Case Is > 25
                            rpart = 255
                            gpart = 0
                            bpart = 0
                        End Select
                    Case 3 '10 to 50%
                        'five fixed color groups
                        Select Case iexp(j, i)
                        Case Is < 10
                            rpart = 255
                            gpart = 255
                            bpart = 255
                        Case 10 To 20
                            rpart = 0
                            gpart = 225
                            bpart = 0
                        Case 20 To 30
                            rpart = 255
                            gpart = 255
                            bpart = 0
                        Case 30 To 40
                            rpart = 0
                            gpart = 255
                            bpart = 255
                        Case 40 To 50
                            rpart = 175
                            gpart = 0
                            bpart = 200
                        Case Is > 50
                            rpart = 255
                            gpart = 0
                            bpart = 0
                        End Select
                    Case 4 '20 to 100%
                        'five fixed color groups
                        Select Case iexp(j, i)
                        Case Is < 20
                            rpart = 255
                            gpart = 255
                            bpart = 255
                        Case 20 To 40
                            rpart = 0
                            gpart = 225
                            bpart = 0
                        Case 40 To 60
                            rpart = 255
                            gpart = 255
                            bpart = 0
                        Case 60 To 80
                            rpart = 0
                            gpart = 255
                            bpart = 255
                        Case 80 To 100
                            rpart = 175
                            gpart = 0
                            bpart = 200
                        Case Is > 100
                            rpart = 255
                            gpart = 0
                            bpart = 0
                        End Select
                    End Select
                    
'                    Select Case vpct
'                        Case Is < 50
'                            rpart = 175
'                            gpart = 175
'                            bpart = 175
'                        Case Is < 75
'                            rpart = 0
'                            gpart = 255
'                        Case Is < 99
'                            rpart = (vpct - 75) * 7
'                            If rpart > 255 Then rpart = 255
'                            gpart = 255
'                        Case Else
'                            rpart = 255
'                            gpart = 0 '255 - ((vpct - 90) * 20)
'                            'If gpart < 0 Then gpart = 0
'                    End Select
                    
    '                Select Case vpct
    '                    Case Is < 25
    '                        rpart = 175
    '                        gpart = 175
    '                        bpart = 175
    '                    Case Is < 40
    '                        rpart = 0
    '                        gpart = 255
    '                    Case Is < 90
    '                        rpart = (vpct - 40) * 7
    '                        If rpart > 255 Then rpart = 255
    '                        gpart = 255
    '                    Case Else
    '                        rpart = 255
    '                        gpart = 255 - ((vpct - 90) * 20)
    '                        If gpart < 0 Then gpart = 0
    '                End Select
                
                End If
                ColorIndex(j, i).iRed = rpart
                ColorIndex(j, i).iGreen = gpart
                ColorIndex(j, i).iBlue = bpart
            Next j
            If writeline Then TS.writeline TxtLine
        Next i
        
        TS.writeline
        TS.Write "Max Value = " & MaxVal & ".  Location = " & MaxW & " East and " & MaxH & " North."
        
        TS.Close
        'write the data to a file
        iFile = FreeFile
    
        Open filespec For Binary Access Write As iFile
            Put iFile, , FHbmp
            Put iFile, , IHbmp
            
            Pad_Zeros = iPixW Mod 4
            'Dim Buffer() As ColorIndex_Type
            Dim Buffer() As Byte
            'ReDim Buffer(1 To iPixW) As ColorIndex_Type
            ReDim Buffer(1 To iPixW * 3 + Pad_Zeros) As Byte
            For i = 1 To iPixH
                For j = 1 To iPixW
                    'Buffer(j) = ColorIndex(j, i)
                    Buffer((j - 1) * 3 + 1) = ColorIndex(j, i).iBlue
                    Buffer((j - 1) * 3 + 2) = ColorIndex(j, i).iGreen
                    Buffer((j - 1) * 3 + 3) = ColorIndex(j, i).iRed
                    'Put iFile, , ColorIndex(j, i)
                Next j
                'pad the end of each scan line with zero's to fill out the last DWord
                'For k = 1 To Pad_Zeros
                '    Buffer(iPixW * 3 + k) = CByte(0)
                '    'Put iFile, , CByte(0)
                'Next k
                Put iFile, , Buffer
            Next i
        Close iFile
    
    Screen.MousePointer = vbNormal
End Sub


