Attribute VB_Name = "modTwoPoint"
Option Explicit

Public Const BinFileHead As String = "rfEngineers .bin file  Copyright 2004, all rights reserved."
Public HGain(0 To 359) As Single
Public VGain(-90 To 90) As Single

Enum vbFormClose
    vbFormControlMenu = 0 'The user chose the Close command from the Control menu on the form.
    vbFormCode = 1 'The Unload statement is invoked from code.
    vbAppWindows = 2 'The current Microsoft Windows operating environment session is ending.
    vbAppTaskManager = 3 'The Microsoft Windows Task Manager is closing the application.
    vbFormMDIForm = 4 'An MDI child form is closing because the MDI form is closing.
    vbFormOwner = 5 'A form is closing because its owner is closing.
End Enum


Public Type AntennaFile_Type
'format of data stored in an antenna file
    Name As String
    Format As String
    MaxGain As Double
    Horizontal As String
    Vertical As String
End Type

Public Type MPEType
    ID As String
    SN As Single
    SE As Single
    SH As Single
    TN As Single
    TE As Single
    TH As Single
    Freq As Double
    VGain As Double
    AGain As Double
    MERP As Double
    AERP As Double
    VAng As Double
    AAng As Double
    Dist As Double
    Pden As Double
    Pocc As Double
    Pgp As Double
    Locc As Double
    Lgp As Double
End Type

Public MPE As MPEType

Public outfile As String
Public PiOvr180 As Double
Public gX1 As Double
Public gY1 As Double
Public gZ1 As Double
Public gX2 As Double
Public gY2 As Double
Public gZ2 As Double
Public gAlpha As Double
Public gBeta As Double
Public gGamma As Double
Public gMag As Double
Public gFSO As New FileSystemObject
Public gxoFileTool As New rfSTool01.FileTool

Public Sub Main()

PiOvr180 = 4 * Atn(1) / 180
frmSplash.Show

End Sub

Public Sub Export()
    Dim TS As TextStream
    
    On Error GoTo ETrap
        exportHeader
        If gFSO.FileExists(outfile) Then
            Set TS = gFSO.OpenTextFile(outfile, ForAppending, False)
                TS.Write MPE.ID & Chr(9)
                TS.Write MPE.SE & Chr(9)
                TS.Write MPE.SN & Chr(9)
                TS.Write MPE.SH & Chr(9)
                TS.Write MPE.TE & Chr(9)
                TS.Write MPE.TN & Chr(9)
                TS.Write MPE.TH & Chr(9)
                TS.Write MPE.Freq & Chr(9)
                TS.Write MPE.VGain & Chr(9)
                TS.Write MPE.AGain & Chr(9)
                TS.Write MPE.MERP & Chr(9)
                TS.Write MPE.VAng & Chr(9)
                TS.Write MPE.AAng & Chr(9)
                TS.Write MPE.Dist & Chr(9)
                TS.Write MPE.AERP & Chr(9)
                TS.Write MPE.Pden & Chr(9)
                TS.Write MPE.Pocc & Chr(9)
                TS.writeline MPE.Pgp
            TS.Close
        End If
    On Error GoTo 0
    Exit Sub
ETrap:
    MsgBox "Error Number " & Err.Number & vbCrLf & _
            Err.Description, vbCritical
End Sub

Public Sub exportHeader()
    Dim TS As TextStream
    
    outfile = GetTxtFileForSave
    If Len(outfile) > 0 Then
        If Not gFSO.FileExists(outfile) Then
            Set TS = gFSO.OpenTextFile(outfile, ForWriting, True)
                TS.Write "Source ID" & Chr(9)
                TS.Write "Source-X" & Chr(9)
                TS.Write "Source-Y" & Chr(9)
                TS.Write "Source-Z" & Chr(9)
                TS.Write "Target-X" & Chr(9)
                TS.Write "Target-Y" & Chr(9)
                TS.Write "Target-Z" & Chr(9)
                TS.Write "Frequency" & Chr(9)
                TS.Write "V-Gain" & Chr(9)
                TS.Write "Az-Gain" & Chr(9)
                TS.Write "Max ERP" & Chr(9)
                TS.Write "Elevation" & Chr(9)
                TS.Write "Azimuth" & Chr(9)
                TS.Write "Distance" & Chr(9)
                TS.Write "Adj ERP" & Chr(9)
                TS.Write "Power Density" & Chr(9)
                TS.Write "% Occupational Exposure" & Chr(9)
                TS.writeline "% Public Exposure"
            TS.Close
        End If
    End If
End Sub


''Public Function GetAntennaPattern() As AntennaFile_Type
'''Antenna Pattern Input File Format:
'''
'''The first line contains a string descriptor of the antenna.
'''
'''The Second line contains one word:
'''Power, Field or, Log depending on what format the info is in.
'''Log means the data is in deciBel's
'''Field means the data is in rms Field units (10^[dB/20])
'''Power means the data is in numeric power gain units (10^[dB/10])
'''
'''The next line contains the Maximum Antenna Gain in dB referenced to an isotropic radiator.
'''This is the pattern gain at the max-gain horizontal/vertical angle.
'''
'''The next line should contain the word "Horizontal" or the letter "H".
'''
'''The next 360 lines contain the horizontal antenna pattern
'''values beginning with 0-degrees and ending at 359-degrees.
'''The line may contain a single numeric value representing
'''the NORMALIZED gain of the antenna.
'''
'''The next line should contain the word "Vertical" or the letter "V".
'''
'''The next 181 lines contain the NORMALIZED Vertical antenna pattern values beginning
'''with 0-degrees (straight down) and increasing to 180-degrees (straight up).
'''
''
''    Dim FileSpec As String
''    Dim AntDat As AntennaFile_Type
''    Dim ts As TextStream
''    Dim i As Integer
''    Dim MaxG As Double
''
''    FileSpec = GetTxtFileForOpen
''    If FileSpec <> "" Then
''        On Error GoTo EFunction
''            Set ts = gFSO.OpenTextFile(FileSpec)
''                With AntDat
''                    .Name = ts.ReadLine
''                    .Format = Trim(UCase(ts.ReadLine))
''                    If .Format = "LOG" Then
''                        MaxG = 0
''                    Else
''                        MaxG = 1
''                    End If
''                    .MaxGain = CDbl(ts.ReadLine)
''                    .Horizontal = ts.ReadLine
''                    For i = 0 To 359
''                        .HGain(i) = CDbl(ts.ReadLine)
''                        If .HGain(i) > MaxG Then
''                            Error.Raise vbObjectError + 10002, "RFR Calculator", "Horizontal antenna gain exceeds unity."
''                        End If
''                    Next i
''                    .Vertical = ts.ReadLine
''                    For i = 0 To 180
''                        .VGain(i) = CDbl(ts.ReadLine)
''                        If .VGain(i) > MaxG Then
''                            Error.Raise vbObjectError + 10001, "RFR Calculator", "Vertical antenna gain exceeds unity."
''                        End If
''                    Next i
''                End With
''            ts.Close
''        On Error GoTo 0
''    End If
''EFunction:
''If Err.Number = 0 Then
''    GetAntennaPattern = AntDat
''End If
''End Function

Public Function GetTxtFileForOpen() As String
    Static LastFile As String

    With frmSplash!CommonDialog1
        .DialogTitle = "Read From Tab Separated Value File"
        .DefaultExt = ".tsv"
        '.InitDir = gJobDIR
        .FileName = LastFile
        .Flags = cdlOFNFileMustExist
'        .Filter = "Text File (*.tsv)|*.tsv"
        .Filter = "Tab delimited File (*.tsv)|*.tsv|ERI File (*.txt)|*.txt|All Files (*.*)|*.*"
        .CancelError = True
        On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                GetTxtFileForOpen = ""
            Else
                GetTxtFileForOpen = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function

Public Function GetVPatFileForOpen() As String
    Static LastFile As String

    With frmSplash!CommonDialog1
        .DialogTitle = "Read From Vertical Antenna Pattern File"
        .DefaultExt = ".tsv"
        '.InitDir = gJobDIR
        .FileName = LastFile
        .Flags = cdlOFNFileMustExist
        .Filter = "Tab delimited File (*.tsv)|*.tsv|ERI File (*.txt)|*.txt|Generic Vert File (*.gv)|*.gv"
        .CancelError = True
        On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                GetVPatFileForOpen = ""
            Else
                GetVPatFileForOpen = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function

Public Function GetHPatFileForOpen() As String
    Static LastFile As String

    With frmSplash!CommonDialog1
        .DialogTitle = "Read From Horizontal Antenna Pattern File"
        .DefaultExt = ".gh"
        '.InitDir = gJobDIR
        .FileName = LastFile
        .Flags = cdlOFNFileMustExist
        .Filter = "Tab delimited File (*.tsv)|*.tsv|Generic Az File (*.gh)|*.gh"
        .CancelError = True
        On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                GetHPatFileForOpen = ""
            Else
                GetHPatFileForOpen = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function


Public Function GetTxtFileForSave() As String
    Static LastFile As String
    
    With frmSplash!CommonDialog1
        .DialogTitle = "Save to Tab Separated Value File"
        .DefaultExt = ".tsv"
        '.InitDir = gJobDIR
        .Flags = cdlOFNPathMustExist 'cdlOFNOverwritePrompt
        .FileName = LastFile
        .Filter = "Tab Separated Value File (*.tsv)|*.tsv"
        .CancelError = True
        On Error Resume Next
            .ShowSave
            If Err <> 0 Then
                GetTxtFileForSave = ""
            Else
                GetTxtFileForSave = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function
Public Function GetTxtFileForRasterSave() As String
    Static LastFile As String
    
    With frmSplash!CommonDialog1
        .DialogTitle = "Save to Tab Separated Value File"
        .DefaultExt = ".tsv"
        '.InitDir = gJobDIR
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .FileName = LastFile
        .Filter = "Tab Separated Value File (*.tsv)|*.tsv"
        .CancelError = True
        On Error Resume Next
            .ShowSave
            If Err <> 0 Then
                GetTxtFileForRasterSave = ""
            Else
                GetTxtFileForRasterSave = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function

Public Function GetbmpFileForSave() As String
    Static LastFile As String
    
    With frmSplash!CommonDialog1
        .DialogTitle = "Save to Bitmap File"
        .DefaultExt = ".bmp"
        '.InitDir = gJobDIR
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .FileName = LastFile
        .Filter = "Bitmap File (*.bmp)|*.bmp"
        .CancelError = True
        On Error Resume Next
            .ShowSave
            If Err <> 0 Then
                GetbmpFileForSave = ""
            Else
                GetbmpFileForSave = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function


Public Function GetBinFileForSave() As String
    Static LastFile As String
    
    With frmSplash!CommonDialog1
        .DialogTitle = "Save to Data File"
        .DefaultExt = ".bin"
        '.InitDir = gJobDIR
        .Flags = cdlOFNPathMustExist 'cdlOFNOverwritePrompt
        .FileName = LastFile
        .Filter = "Binary Data File (*.bin)|*.bin"
        .CancelError = True
        On Error Resume Next
            .ShowSave
            If Err <> 0 Then
                GetBinFileForSave = ""
            Else
                GetBinFileForSave = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function
Public Function GetBinFileForOpen() As String
    Static LastFile As String

    With frmSplash!CommonDialog1
        .DialogTitle = "Read From Data File"
        .DefaultExt = ".bin"
        '.InitDir = gJobDIR
        .FileName = LastFile
        .Flags = cdlOFNFileMustExist
        .Filter = "Binary Data File (*.bin)|*.bin"
        .CancelError = True
        On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                GetBinFileForOpen = ""
            Else
                GetBinFileForOpen = .FileName
                LastFile = .FileName
            End If
        On Error GoTo 0
    End With
End Function

Public Sub LoadMPE()
    Dim lFile As String
    Dim iFile As Integer
    Dim Txt As String * 59
    
    lFile = GetBinFileForOpen
    If gFSO.FileExists(lFile) Then
        iFile = FreeFile
        Open lFile For Binary Access Read As #iFile
            Get #iFile, , Txt
            If Txt = BinFileHead Then
                Get #iFile, , MPE
            End If
        Close #iFile
    End If

End Sub

Public Sub saveMPE()
    Dim lFile As String
    Dim iFile As Integer
    Dim Txt As String * 59
    
    lFile = GetBinFileForSave
    If Len(lFile) > 0 Then
        iFile = FreeFile
        Open lFile For Binary Access Write As #iFile
            Put #iFile, , BinFileHead
            Put #iFile, , MPE
        Close #iFile
    End If
End Sub


Public Static Function Log10(X As Double) As Double
    If X = 0 Then
        Log10 = 0
    Else
       Log10 = Log(X) / Log(10#)
    End If
End Function


