Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Public Class Form1
    <DllImport("Everything64.dll", EntryPoint:="Everything_SetSearchW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_SetSearchW(search As String) As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_SetRequestFlags", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_SetRequestFlags(dwRequestFlags As UInt32) As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_QueryW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_QueryW(bWait As Integer) As Integer
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_GetNumResults", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_GetNumResults() As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_GetResultFullPathNameW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_GetResultFullPathNameW(index As UInt32, buf As StringBuilder, size As UInt32) As UInt32
    End Function
    Public Const UppLim = 1000000
    Public Const EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME = &H4
    Public FolderPath1 As String
    Public FolderPath2 As String
    Public RawData1() As String = New String(UppLim) {}
    Public RawData2() As String = New String(UppLim) {}
    Public RawCount1 As Long
    Public RawCount2 As Long
    Public DupCheck_1() As Long = New Long(UppLim) {}
    Public DupCheck_2() As Long = New Long(UppLim) {}
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FolderPath1 = ""
        FolderPath2 = ""
        RawCount1 = 0
        RawCount2 = 0
        ListBox4.Items.Add(Now + "请选择双文件夹路径")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FolderBrowserDialog1 As New FolderBrowserDialog
        Dim AppPath As String, Comp_1 As String, Comp_2 As String
        Dim i As Long
        Dim FolderCheck As Boolean
        AppPath = Directory.GetCurrentDirectory()
        If Mid(AppPath, Len(AppPath), 1) <> "\" Then
            AppPath += "\"
        End If
        If (FolderBrowserDialog1.ShowDialog = DialogResult.OK) Then
            FolderPath1 = FolderBrowserDialog1.SelectedPath
            If Mid(FolderPath1, Len(FolderPath1), 1) <> "\" Then
                FolderPath1 += "\"
            End If
            ListBox4.Items.Add(Now + "已选择A文件夹路径:" + FolderPath1)
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Else
            ListBox4.Items.Add(Now + "已取消A文件夹选择")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
        FolderCheck = True
        If Len(FolderPath1) = Len(AppPath) Then
            Comp_1 = ""
            Comp_2 = ""
            For i = 1 To Len(FolderPath1)
                If "a" <= Mid(FolderPath1, i, 1) And Mid(FolderPath1, i, 1) <= "z" Then
                    Comp_1 += Chr(Asc(Mid(FolderPath1, i, 1)) - 32)
                Else
                    Comp_1 += Mid(FolderPath1, i, 1)
                End If
                If "a" <= Mid(AppPath, i, 1) And Mid(AppPath, i, 1) <= "z" Then
                    Comp_2 += Chr(Asc(Mid(AppPath, i, 1)) - 32)
                Else
                    Comp_2 += Mid(AppPath, i, 1)
                End If
            Next i
            If StrComp(Comp_1, Comp_2) = 0 Then FolderCheck = False
        End If
        If FolderCheck = False Then
            ListBox4.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:文件夹路径不能为程序所在文件夹", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            FolderPath2 = FolderBrowserDialog1.SelectedPath
            If Mid(FolderPath2, Len(FolderPath2), 1) <> "\" Then
                FolderPath2 += "\"
            End If
            If StrComp(FolderPath1, FolderPath2) <> 0 Then
                ListBox4.Items.Add(Now + "已选择B文件夹路径:" + FolderPath2)
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                Button2.Enabled = True
                Button3.Enabled = False
            Else
                ListBox4.Items.Add(Now + "文件夹路径错误,程序已终止")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("错误:不能选择两个相同的文件夹路径", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
                Button2.Enabled = False
                Button3.Enabled = False
                Exit Sub
            End If
        Else
            ListBox4.Items.Add(Now + "已取消B文件夹选择")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
        FolderCheck = True
        If Len(FolderPath2) = Len(AppPath) Then
            Comp_1 = ""
            Comp_2 = ""
            For i = 1 To Len(FolderPath2)
                If "a" <= Mid(FolderPath2, i, 1) And Mid(FolderPath2, i, 1) <= "z" Then
                    Comp_1 += Chr(Asc(Mid(FolderPath2, i, 1)) - 32)
                Else
                    Comp_1 += Mid(FolderPath2, i, 1)
                End If
                If "a" <= Mid(AppPath, i, 1) And Mid(AppPath, i, 1) <= "z" Then
                    Comp_2 += Chr(Asc(Mid(AppPath, i, 1)) - 32)
                Else
                    Comp_2 += Mid(AppPath, i, 1)
                End If
            Next i
            If StrComp(Comp_1, Comp_2) = 0 Then FolderCheck = False
        End If
        If FolderCheck = False Then
            ListBox4.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:文件夹路径不能为程序所在文件夹", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Filename As New StringBuilder(260)
        Dim FlagStr As String
        Dim AllLength1 As Long
        Dim AllLength2 As Long
        Dim NumResults As Long
        Dim i As Long
        Button3.Enabled = False
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        If Mid(FolderPath1, Len(FolderPath1), 1) <> "\" Then FolderPath1 += "\"
        If Mid(FolderPath2, Len(FolderPath2), 1) <> "\" Then FolderPath2 += "\"
        ListBox4.Items.Add(Now + "正在确认A文件夹路径是否有误")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If Boo_DirExist(FolderPath1) = 0 Then
            ListBox4.Items.Add(Now + "A文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:未找到A文件夹路径,请重新指定", 0 + vbCritical + vbSystemModal, "A文件夹路径错误")
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
        ListBox4.Items.Add(Now + "A文件夹路径无误,正在确认B文件夹路径是否有误")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If Boo_DirExist(FolderPath2) = 0 Then
            ListBox4.Items.Add(Now + "B文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:未找到B文件夹路径,请重新指定", 0 + vbCritical + vbSystemModal, "B文件夹路径错误")
            Button2.Enabled = False
            Button3.Enabled = False
            Exit Sub
        End If
        ListBox4.Items.Add(Now + "B文件夹路径无误,开始枚举双文件夹下的全部文件")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Everything_SetSearchW(FolderPath1)
        Everything_SetRequestFlags(EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME)
        Everything_QueryW(1)
        NumResults = Everything_GetNumResults()
        RawCount1 = 0
        For i = 0 To NumResults - 1
            Everything_GetResultFullPathNameW(i, Filename, Filename.Capacity)
            If File.Exists(Filename.ToString) = True Then
                RawData1(RawCount1) = Filename.ToString
                If StrComp(FolderPath1, Mid(RawData1(RawCount1), 1, Len(FolderPath1))) = 0 Then
                    RawCount1 += 1
                End If
                If RawCount1 = UppLim Then
                    ListBox4.Items.Add(Now + "A文件夹文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    MsgBox("A文件夹文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件", 0 + vbExclamation + vbSystemModal, "警告")
                    Exit For
                End If
            End If
        Next i
        For i = 0 To RawCount1 - 1
            RawData1(i) = Mid(RawData1(i), Len(FolderPath1) + 1)
        Next i
        ListBox4.Items.Add(Now + "A文件夹枚举完成,开始输出文件列表")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        ListBox1.Items.Add(Now + "|A文件夹文件列表如下:")
        For i = 0 To RawCount1 - 1
            ListBox1.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + RawData1(i))
        Next i
        ListBox4.Items.Add(Now + "A文件夹文件列表输出完成,导入文件总个数为:" + CStr(RawCount1))
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If RawCount1 = 0 Then
            ListBox1.Items.Add("A文件夹下没有文件")
            MsgBox("A文件夹下没有文件", 0 + vbInformation + vbSystemModal, "文件夹对比")
        End If
        Everything_SetSearchW(FolderPath2)
        Everything_SetRequestFlags(EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME)
        Everything_QueryW(1)
        NumResults = Everything_GetNumResults()
        RawCount2 = 0
        For i = 0 To NumResults - 1
            Everything_GetResultFullPathNameW(i, Filename, Filename.Capacity)
            If File.Exists(Filename.ToString) = True Then
                RawData2(RawCount2) = Filename.ToString
                If StrComp(FolderPath2, Mid(RawData2(RawCount2), 1, Len(FolderPath2))) = 0 Then
                    RawCount2 += 1
                End If
                If RawCount2 = UppLim Then
                    ListBox4.Items.Add(Now + "B文件夹文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    MsgBox("B文件夹文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件", 0 + vbExclamation + vbSystemModal, "警告")
                    Exit For
                End If
            End If
        Next i
        For i = 0 To RawCount2 - 1
            RawData2(i) = Mid(RawData2(i), Len(FolderPath2) + 1)
        Next i
        ListBox4.Items.Add(Now + "B文件夹枚举完成,开始输出文件列表")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        ListBox2.Items.Add(Now + "|B文件夹文件列表如下:")
        For i = 0 To RawCount2 - 1
            ListBox2.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + RawData2(i))
        Next i
        ListBox4.Items.Add(Now + "B文件夹文件列表输出完成,导入文件总个数为:" + CStr(RawCount2))
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If RawCount2 = 0 Then
            ListBox2.Items.Add("B文件夹下没有文件")
            MsgBox("B文件夹下没有文件", 0 + vbInformation + vbSystemModal, "文件夹对比")
        End If
        ListBox4.Items.Add(Now + "开始计算A文件夹大小")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        AllLength1 = 0
        For i = 0 To RawCount1 - 1
            Dim finfo As New FileInfo(FolderPath1 + RawData1(i))
            If finfo.Exists = True Then
                DupCheck_1(i) = finfo.Length
                AllLength1 += finfo.Length
            Else
                DupCheck_1(i) = 0
                ListBox4.Items.Add(Now + "错误:A文件夹文件" + RawData1(i) + "文件大小读取失败,将该文件大小设置为0")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            End If
        Next i
        ListBox4.Items.Add(Now + "A文件夹大小计算完成,共" + CStr(AllLength1) + "字节")
        ListBox4.Items.Add(Now + "开始计算B文件夹大小")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        AllLength2 = 0
        For i = 0 To RawCount2 - 1
            Dim finfo As New FileInfo(FolderPath2 + RawData2(i))
            If finfo.Exists = True Then
                DupCheck_2(i) = finfo.Length
                AllLength2 += finfo.Length
            Else
                DupCheck_2(i) = 0
                ListBox4.Items.Add(Now + "错误:B文件夹文件" + RawData2(i) + "文件大小读取失败,将该文件大小设置为0")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            End If
        Next i
        ListBox4.Items.Add(Now + "B文件夹大小计算完成,共" + CStr(AllLength2) + "字节")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If AllLength1 = AllLength2 Then
            FlagStr = "双文件夹大小相同"
        Else
            FlagStr = "双文件夹大小不同"
        End If
        MsgBox("双文件夹大小计算完成" + vbCrLf + "A文件夹大小为:" + CStr(AllLength1) + "字节" + vbCrLf + "B文件夹大小为:" + CStr(AllLength2) + "字节" + vbCrLf + FlagStr, 0 + vbInformation + vbSystemModal, "文件夹对比")
        If RawCount1 <> 0 Or RawCount2 <> 0 Then
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i_Same As Long, i_A_Diff As Long, i_B_Diff As Long
        Dim i As Long, j As Long
        Dim Flag1() As Long = New Long(UppLim) {}
        Dim Flag2() As Long = New Long(UppLim) {}
        Dim Same() As String = New String(UppLim) {}
        Dim Same_Size() As Long = New Long(UppLim) {}
        Dim A_Diff() As String = New String(UppLim) {}
        Dim A_Serial() As Long = New Long(UppLim) {}
        Dim A_Size() As Long = New Long(UppLim) {}
        Dim B_Diff() As String = New String(UppLim) {}
        Dim B_Serial() As Long = New Long(UppLim) {}
        Dim B_Size() As Long = New Long(UppLim) {}
        ListBox3.Items.Clear()
        ListBox4.Items.Add(Now + "开始文件夹对比")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        i_Same = -1
        For i = 0 To RawCount1 - 1
            For j = 0 To RawCount2 - 1
                If StrComp(RawData1(i), RawData2(j)) = 0 And DupCheck_1(i) = DupCheck_2(j) Then
                    i_Same += 1
                    Same(i_Same) = RawData1(i)
                    Same_Size(i_Same) = DupCheck_1(i)
                    Flag1(i) = 1
                    Flag2(j) = 1
                    Exit For
                End If
            Next j
        Next i
        i_A_Diff = -1
        For i = 0 To RawCount1 - 1
            If Flag1(i) = 0 Then
                i_A_Diff += 1
                A_Diff(i_A_Diff) = RawData1(i)
                A_Serial(i_A_Diff) = i
                A_Size(i_A_Diff) = DupCheck_1(i)
            End If
        Next i
        i_B_Diff = -1
        For i = 0 To RawCount2 - 1
            If Flag2(i) = 0 Then
                i_B_Diff += 1
                B_Diff(i_B_Diff) = RawData2(i)
                B_Serial(i_B_Diff) = i
                B_Size(i_B_Diff) = DupCheck_2(i)
            End If
        Next i
        ListBox4.Items.Add(Now + "文件夹对比完成,开始输出对比结果")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        ListBox3.Items.Add(Now + "|文件夹对比结果如下:")
        ListBox3.Items.Add("双文件夹存在以下相同文件:")
        For i = 0 To i_Same
            ListBox3.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + Same(i))
        Next i
        If i_Same = -1 Then ListBox3.Items.Add("双文件夹下无相同文件")
        ListBox3.Items.Add("A文件夹存在以下不同文件:")
        For i = 0 To i_A_Diff
            ListBox3.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + NumLength(Trim(Str(A_Serial(i) + 1))) + "|" + A_Diff(i) + "|" + CStr(A_Size(i)) + "字节")
        Next i
        If i_A_Diff = -1 Then ListBox3.Items.Add("A文件夹下无不同文件")
        ListBox3.Items.Add("B文件夹存在以下不同文件:")
        For i = 0 To i_B_Diff
            ListBox3.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + NumLength(Trim(Str(B_Serial(i) + 1))) + "|" + B_Diff(i) + "|" + CStr(B_Size(i)) + "字节")
        Next i
        If i_B_Diff = -1 Then ListBox3.Items.Add("B文件夹下无不同文件")
        ListBox4.Items.Add(Now + "对比结果输出完成,双文件夹相同文件个数:" + CStr(i_Same + 1) + "|A文件夹不同文件个数:" + CStr(i_A_Diff + 1) + "|B文件夹不同文件个数:" + CStr(i_B_Diff + 1))
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        MsgBox("对比结果输出完成" + vbCrLf + "双文件夹相同文件个数:" + CStr(i_Same + 1) + vbCrLf + "A文件夹不同文件个数:" + CStr(i_A_Diff + 1) + vbCrLf + "B文件夹不同文件个数:" + CStr(i_B_Diff + 1), 0 + vbInformation + vbSystemModal, "文件夹对比")
    End Sub
    Private Shared Function Boo_DirExist(StrPath As String) As Boolean
        Boo_DirExist = Directory.Exists(StrPath)
    End Function
    Public Shared Function NumLength(Num As String) As String
        If Len(Num) = 1 Then
            NumLength = "     " + Num
        ElseIf Len(Num) = 2 Then
            NumLength = "    " + Num
        ElseIf Len(Num) = 3 Then
            NumLength = "   " + Num
        ElseIf Len(Num) = 4 Then
            NumLength = "  " + Num
        ElseIf Len(Num) = 5 Then
            NumLength = " " + Num
        Else
            NumLength = Num
        End If
    End Function
End Class