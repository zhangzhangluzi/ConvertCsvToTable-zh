"""Reproduce the Chinese localization from the upstream source checkout."""
from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
BUILD.mkdir(exist_ok=True)


def upstream(path):
    raw = subprocess.check_output(["git", "show", f"origin/master:{path}"], cwd=ROOT)
    return raw.decode("cp932").replace("\r\r\n", "\n").replace("\r\n", "\n")


def zh(text):
    return 'ZhText("' + " ".join(f"{ord(c):04X}" for c in text) + '")'


def save_source(path, content):
    content = "\n".join(line.rstrip() for line in content.splitlines()) + "\n"
    (ROOT / path).write_text(content, encoding="utf-8-sig", newline="\n")
    (BUILD / Path(path).name).write_text(content, encoding="ascii", errors="replace", newline="\r\n")
    # VBA source imports depend on the host code page. Runtime labels use ChrW,
    # so the compiled source can safely stay ASCII on either Chinese or Japanese Office.
    code = re.sub(r"^Attribute .*\n", "", content, flags=re.M)
    code = re.sub(r"\AVERSION 1\.0 CLASS\nBEGIN\n.*?\nEND\n", "", code, flags=re.S)
    code = "\n".join(line for line in code.splitlines() if not line.lstrip().startswith("'"))
    (BUILD / (Path(path).stem + ".code.txt")).write_text(code, encoding="ascii", errors="replace", newline="\r\n")


control_path = "src/Modules/StdCsvOpenerControl.bas"
control = upstream(control_path)
labels = {
    "ApplicationName": "CSV 转换",
    "ControlLabel": "自动转换",
    "ToolsLabel": "工具",
    "StartAutomaticConversionLabel": "开启自动转换",
    "StopAutomaticConversionLabel": "停止自动转换",
    "ManualConversionLabel": "手动转换",
    "TextToNumberConversionLabel": "文本转数字",
    "TextToLinkConversionLabel": "文本转超链接",
}
for callback, label in labels.items():
    pattern = rf"(Sub RibbonControl_{callback}\([^\n]+\)\n).*?\nEnd Sub"
    control, count = re.subn(pattern, lambda m: m[1] + "    ReturnedVal = " + zh(label) + "\nEnd Sub", control, flags=re.S)
    assert count == 1, callback
control += '''
' 中文标签通过 Unicode 码位生成，避免 VBA 工程代码页影响汉字。
Public Function ZhText(ByVal HexCodes As String) As String
    Dim CodePoint As Variant
    For Each CodePoint In Split(HexCodes, " ")
        ZhText = ZhText & ChrW(CLng("&H" & CodePoint))
    Next CodePoint
End Function
'''
save_source(control_path, control)

reader_path = "src/Class Modules/ClsCsvOpener.cls"
reader = upstream(reader_path)
for english, chinese in {
    "Start: Converting CSV file to table": "开始将 CSV 转换为表格…",
    "Reading CSV file ...": "正在读取 CSV 文件…",
    "Analyzing CSV file ...": "正在分析 CSV 文件…",
    "Converting CSV file to UTF-8 ...": "正在转换为 UTF-8 编码…",
    "Converting CSV file to tables ...": "正在生成表格…",
}.items():
    assert f'"{english}"' in reader
    reader = reader.replace(f'"{english}"', zh(chinese))
reader, count = re.subn(r'MsgBox IIf\( _\n\s*Application\.International\(xlCountryCode\) = 81, _\n\s*"CSVファイルではありません", _\n\s*"It\'s not a CSV file !" _\n\s*\)', "MsgBox " + zh("当前文件不是 CSV 文件，请先打开 CSV 文件。") + ", vbInformation, " + zh("CSV 转换"), reader)
assert count == 1
reader, count = re.subn(r'MsgBox IIf\( _\n\s*Application\.International\(xlCountryCode\) = 81, _\n\s*"既に変換済みのようです", _\n\s*"It failed to convert or has already converted !" _\n\s*\)', "MsgBox " + zh("文件可能已经转换，或本次转换未能完成。") + ", vbInformation, " + zh("CSV 转换"), reader)
assert count == 1
# Prefer strict UTF-8 detection over the original locale-dependent ADODB guess.
# Legacy Chinese files fall back to GBK; BOM-marked UTF-16 stays explicit.
reader = reader.replace("    ReDim ByteBuffer(0 To ByteLength - 1)", "    If ByteLength = 0 Then\n        Close #FileNumber\n        Exit Function\n    End If\n    ReDim ByteBuffer(0 To ByteLength - 1)")
reader, count = re.subn(r'    If FileCharset = "AUTO" And 3 <= ByteLength Then\n.*?\n    End If', '''    If FileCharset = "AUTO" Then
        If IsValidUtf8(ByteBuffer) Then
            FileCharset = "utf-8"
        ElseIf ByteLength >= 2 Then
            If ByteBuffer(0) = &HFF And ByteBuffer(1) = &HFE Then
                FileCharset = "unicode"
            ElseIf ByteBuffer(0) = &HFE And ByteBuffer(1) = &HFF Then
                FileCharset = "unicodeFFFE"
            Else
                FileCharset = "gb2312"
            End If
        Else
            FileCharset = "gb2312"
        End If
    End If''', reader, count=1, flags=re.S)
assert count == 1
utf8_validator = '''
' 严格校验 UTF-8：拒绝截断、过长编码、代理区和超出 Unicode 范围的字节。
Private Function IsValidUtf8(ByRef Data() As Byte) As Boolean
    Dim Index As Long, Last As Long, Lead As Long, Needed As Long, Offset As Long
    Last = UBound(Data)
    Index = LBound(Data)
    Do While Index <= Last
        Lead = CLng(Data(Index))
        If Lead < &H80 Then
            Needed = 0
        ElseIf Lead >= &HC2 And Lead <= &HDF Then
            Needed = 1
        ElseIf Lead >= &HE0 And Lead <= &HEF Then
            Needed = 2
        ElseIf Lead >= &HF0 And Lead <= &HF4 Then
            Needed = 3
        Else
            Exit Function
        End If
        If Index + Needed > Last Then Exit Function
        For Offset = 1 To Needed
            If Data(Index + Offset) < &H80 Or Data(Index + Offset) > &HBF Then Exit Function
        Next Offset
        If Needed > 0 Then
            If Lead = &HE0 And Data(Index + 1) < &HA0 Then Exit Function
            If Lead = &HED And Data(Index + 1) > &H9F Then Exit Function
            If Lead = &HF0 And Data(Index + 1) < &H90 Then Exit Function
            If Lead = &HF4 And Data(Index + 1) > &H8F Then Exit Function
        End If
        Index = Index + Needed + 1
    Loop
    IsValidUtf8 = True
End Function

'''
reader = reader.replace("Public Sub AnalyzeCsvText(", utf8_validator + "Public Sub AnalyzeCsvText(")
save_source(reader_path, reader)

# The upstream source uses Set to return a Boolean. Recompiling the project
# exposes that dormant type error, so return the value with normal assignment.
timer_path = "src/Class Modules/ClsTimeMeasurement.cls"
timer = upstream(timer_path)
timer = timer.replace("Private PrivateTimeMeasurementIsValid\n", "Private PrivateTimeMeasurementIsValid As Boolean\n")
timer = timer.replace("Set TimeMeasurementIsValid = PrivateTimeMeasurementIsValid", "TimeMeasurementIsValid = PrivateTimeMeasurementIsValid")
save_source(timer_path, timer)

script_text = {
    'IIf(IsJA, "CSV変換", "Convert CSV To Table")': '"CSV 转换（中文版）"',
    'IIf(IsJA, "Excel を全て閉じてください！", "Please close all Excel applications !")': '"请先保存工作并关闭所有 Excel 窗口。"',
    'IIf(IsJA, "アドインをインストールしますか？", "Do you want to install this add-in ?")': '"是否安装 CSV 转换中文版？"',
    'IIf(IsJA, "アドインのインストールが完了しました", "Installation is now complete !")': '"安装完成。打开 Excel 后，可在“CSV 转换”选项卡中开启自动转换。"',
    'IIf(IsJA, "アドインをアンインストールしますか？", "Do you want to uinstall this add-in ?")': '"是否卸载 CSV 转换中文版？"',
    'IIF(IsJA, "アドインのアンインストールが完了しました", "Uninstallation is now complete !")': '"卸载完成。"',
    'IIf(IsJA, "エラーが発生しました: " & CStr(Err.Number) & vbCrLf & "実行環境を確認してください", "An error has occurred." & vbCrLf & "Please check your environment.")': '"操作失败，错误代码：" & CStr(Err.Number) & vbCrLf & "请检查 Excel 是否已经关闭。"',
}
for name in ["Install.vbs", "Uninstall.vbs"]:
    content = upstream("addin/" + name)
    for original, replacement in script_text.items():
        content = content.replace(original, replacement)
    content = content.replace('addInFileName = "ConvertCsvToTable.xlam"', 'addInFileName = "ConvertCsvToTableZh.xlam"')
    if name == "Install.vbs":
        content = content.replace('objFileSys.CopyFile addInFileName , installPath, True', 'objFileSys.CopyFile objFileSys.BuildPath(objFileSys.GetParentFolderName(WScript.ScriptFullName), addInFileName), installPath, True')
    (ROOT / "addin" / name).write_text(content, encoding="utf-16", newline="\r\n")

print("Localized source, Chinese installer/uninstaller, and ASCII VBA build inputs generated.")
