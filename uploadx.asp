<%Dim FormData, FormSize, Divider, bCrLf
FormSize = Request.TotalBytes
FormData = Request.BinaryRead(FormSize)
bCrLf = ChrB(13) & ChrB(10)
Divider = LeftB(FormData, InStrB(FormData, bCrLf) - 1)

Function SaveFile(FormFileField, Path, MaxSize, SavType)
Dim StreamObj,StreamObj1
Set StreamObj = Server.CreateObject("ADODB.Stream")
Set StreamObj1 = Server.CreateObject("ADODB.Stream")
StreamObj.Mode = 3
StreamObj1.Mode = 3
StreamObj.Type = 1
StreamObj1.Type = 1
SaveFile = ""
StartPos = LenB(Divider) + 2
FormFileField = Chr(34) & FormFileField & Chr(34)
If Right(Path,1) <> "\" Then
Path = Path & "\"
End If
Do While StartPos > 0
strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
SearchStr = MidB(FormData, StartPos, strlen)
If InStr(bin2str(SearchStr), FormFileField) > 0 Then
FileName = bin2str(GetFileName(SearchStr,path,SavType))
If FileName <> "" Then
FileStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
FileLen = InStrB(StartPos, FormData, Divider) - 2 - FileStart
If FileLen <= MaxSize*1024 Then
FileContent = MidB(FormData, FileStart, FileLen)
StreamObj.Open
StreamObj1.Open
StreamObj.Write FormData
StreamObj.Position=FileStart-1
StreamObj.CopyTo StreamObj1,FileLen
If SavType =0 Then
SavType = 1
End If 
StreamObj1.SaveToFile Path & FileName, SavType
StreamObj.Close
StreamObj1.Close
If SaveFile <> "" Then
SaveFile = SaveFile & "," & FileName
Else
SaveFile = FileName
End If
Else
If SaveFile <> "" Then
SaveFile = SaveFile & ",*TooBig*"
Else
SaveFile = "*TooBig*"
End If
End If
End If
End If
If InStrB(StartPos, FormData, Divider) < 1 Then
Exit Do
End If
StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
Loop
End Function

Function GetFormVal(FormName)
GetFormVal = ""
StartPos = LenB(Divider) + 2
FormName = Chr(34) & FormName & Chr(34)
Do While StartPos > 0
strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
SearchStr = MidB(FormData, StartPos, strlen)
If InStr(bin2str(SearchStr), FormName) > 0 Then
ValStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
ValLen = InStrB(StartPos, FormData, Divider) - 2 - ValStart
ValContent = MidB(FormData, ValStart, ValLen)
If GetFormVal <> "" Then
GetFormVal = GetFormVal & "," & bin2str(ValContent)
Else
GetFormVal = bin2str(ValContent)
End If
End If
If InStrB(StartPos, FormData, Divider) < 1 Then
Exit Do
End If
StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
Loop
End Function

Function bin2str(binstr)
Dim varlen, clow, ccc, skipflag
skipflag = 0
ccc = ""
varlen = LenB(binstr)
For i = 1 To varlen
If skipflag = 0 Then
clow = MidB(binstr, i, 1)
If AscB(clow) > 127 Then
ccc = ccc & Chr(AscW(MidB(binstr, i + 1, 1) & clow))
skipflag = 1
Else
ccc = ccc & Chr(AscB(clow))
End If
Else
skipflag = 0
End If
Next
bin2str = ccc
End Function

Function str2bin(str)
For i = 1 To Len(str)
str2bin = str2bin & ChrB(Asc(Mid(str, i, 1)))
Next
End Function

Function GetFileName(str,path,savtype)
Set fs = Server.CreateObject("Scripting.FileSystemObject")
str = RightB(str,LenB(str)-InstrB(str,str2bin("filename="))-9)
GetFileName = ""
FileName = ""
For i = LenB(str) To 1 Step -1
If MidB(str, i, 1) = ChrB(Asc("\")) Then
FileName = MidB(str, i + 1, LenB(str) - i - 1)
Exit For
End If
Next
If savtype = 0 and fs.FileExists(path & bin2str(FileName)) = True Then
hFileName = FileName
rFileName = ""
For i = LenB(FileName) To 1 Step -1
If MidB(FileName, i, 1) = ChrB(Asc(".")) Then
hFileName = LeftB(FileName, i-1)
rFileName = RightB(FileName, LenB(FileName)-i+1)
Exit For
End If
Next
For i = 0 to 9999 
hFileName = hFileName & str2bin(i)
If fs.FileExists(path & bin2str(hFileName) & i & bin2str(rFileName)) = False Then
Set fs = Server.CreateObject("Scripting.FileSystemObject")
File = Server.MapPath("upload/"&bin2str(hFileName)& bin2str(rFileName))

On Error Resume Next ' �Y�����~�̵M�V�U����
fs.DeleteFile File, True
If Err.Number = 53 Then
Response.Write "�ɦW�M�{��������!<BR>"
Response.Write "<a href=upload0.asp?mode=del&del_file="&File&">�Ч���ɦW��A���s�W��</a>"
Response.End
ElseIf Err.Number = 70 Then
Response.Write File & " �w�Q��w!"
Response.End
ElseIf Err.Number <> 0 Then 
Response.Write "��L���������~, ���~�s��=" & Err.Number
Response.End
End If 
FileName = hFileName & str2bin & rFileName 'str2bin ���ɦW
Exit For
End If
Next
End If
Set fs = Nothing
GetFileName = FileName
End Function%>