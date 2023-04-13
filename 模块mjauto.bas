Attribute VB_Name = "模块mjauto"



Sub chaxundingdan(x1)
Exit Sub
evt = Application.EnableEvents
 Application.EnableEvents = False

 Dim rr, c As Range
 Dim tuhao, xstr As String
 Dim col, quexiaocol As Long
 
 tuhao = Cells(x1, "A").Value
 With Worksheets("订单计划")
    Set rr = Range(.Cells(1, "A"), .Cells(3000, "CZ"))
     
 End With
  
  Set c = rr.Columns("A").find(tuhao, LookIn:=xlValues, LookAt:=xlWhole)
  
    xstr = ""
    If Not c Is Nothing Then
       For col = Columns("C").Column To Worksheets("订单计划").Range("缺少订单").Column - 1 Step 3
             If rr(c.row, col).Value <> 0 And rr(c.row, col).Value <> "" Then
                xstr = xstr & rr(2, col + 1).Value & ") " & rr(c.row, col).Value & "/ " & rr(c.row, col + 2).Value & "/ " & rr(c.row, col + 1).Value & ", "
             End If
        Next
     If xstr <> "" Then
        MsgBox xstr & "                 ", vbOKOnly, "订单情况- 公司/欠单/出库/订单  "
    End If
   End If
  Application.EnableEvents = evt

End Sub


Sub chaxunchurutongji(a)
 evt = Application.EnableEvents
  Application.EnableEvents = False

 Dim rr, c As Range
 Dim tuhao, xstr As String
 Dim col, quexiaocol As Long
 Dim x1, x2, x3 As Single
 '‘On Error GoTo errorexit
 unprotectsub
 If a.Value <> "" Then
 
 Select Case a.Column
    
     Case 1
      
      If a.row <= 4 Then
          ActiveSheet.AutoFilterMode = False
       If Cells(2, "A").Value = "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
        
         If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/")
         If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<=" & "20" & datecode(Cells(3, "A").Value, "/")

          If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<=" & "20" & datecode(Cells(3, "A").Value, "/")
     Else
        
            Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(a.Value, "/"), Criteria2:="<=" & "20" & datecode(a.Value, "/")
     End If
      Case 2
       If Cells(2, "B").Value <> "<-开始日期" Then
              ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=2, Criteria1:="=" & a.Value & "*", Operator:=xlOr, Criteria2:=Cells(2, "B").Value
         Else
             ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=2, Criteria1:="=" & a.Value & "*"
         End If
     
         Cells(3, "B").Value = a.Value
         
       Case 3
          Cells(3, "C").Value = a.Value
          'If Cells(2, "C").Value <> "" Then
             ' ActiveSheet.Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:="=" & a.Value, Operator:=xlOr, Criteria2:=Cells(2, "C").Value
        ' Else
           
                   ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:="=" & a.Value
              If Len(Cells(3, "D").Value) >= 5 Then
                   chazhaojianyonggx
              End If
             
         'End If
       
        Case 4
          'If Cells(2, "D") <> "" Then
       '  ActiveSheet.Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter Field:=4, Criteria1:="=" & a.Value, Operator:=xlOr, Criteria2:=Cells(2, "D").Value
            '  If Len(Cells(3, "D").Value) >= 5 Then
              ' chazhaojianyonggx
           ' End If
       '  Else
              ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=4, Criteria1:="=" & a.Value
             ' chazhaojianyonggx
        ' End If
         Cells(3, "D").Value = a.Value
        Case 12
            ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=12, Criteria1:="=" & datecode(a.Value, "/")
        Case 13
          ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=13, Criteria1:=a.Value & "*"
        Case 17
          If Cells(2, "Q").Value <> "" Then
              ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=17, Criteria1:="=" & a.Value & "*", Operator:=xlOr, Criteria2:=Cells(2, "Q").Value & "*"
         Else
             ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=17, Criteria1:="=" & a.Value & "*"
         End If
          
            Cells(3, "Q").Value = a.Value
        Case Columns("P").Column
            ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("P").Column, Criteria1:=a.Value & "*"
        Case Columns("S").Column
            ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("S").Column, Criteria1:=a.Value & "*"
        Case Else
            ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=a.Column, Criteria1:=a.Value & "*"
         
End Select

  
  ' For crow = 5 To Range("A1000000").End(xlUp).Row
  '   x1 = x1 + Cells(crow, "F").Value
  '   x2 = x2 + Cells(crow, "G").Value
  '   x3 = x3 + Cells(crow, "H").Value
    
'Next
     
        
           Application.GoTo Cells(5, "A"), Scroll:=True
      
End If
errorexit:
protectsub
' ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
 Application.EnableEvents = evt
End Sub
Sub sheetjizhun(Optional sh1 As String)

evt = Application.EnableEvents
 wbcheck
 
 Application.EnableEvents = False

Dim x1 As String
Dim a As Range
  
      
      
      
    If ActiveWorkbook.Name = wb.Name Then
       If ows Is Nothing Then Set ows = ActiveSheet
 
              Set ows = ActiveSheet
              If sh1 = "" Then
                 MsgBox ActiveSheet.Name & " 基准 "
              Else
                 Debug.Print ActiveSheet.Name & " 基准 "
              End If
   End If
 
  Application.EnableEvents = True
End Sub
Sub cancelfilter()

evt = Application.EnableEvents

wbcheck
 
 Application.EnableEvents = False
If Val(activerow) < 5 Then activerow = 5
Dim x1 As String
Dim a As Range
  
  On Error GoTo Endp
  On Error GoTo 0
   unprotectsub
      
     If ActiveSheet.Name <> "生产计划" Then
                 Application.OnKey "^c"
                 Application.OnKey "^v"
                 Application.OnKey "^x"
                 Application.OnKey "^C"
                 Application.OnKey "^V"
                 Application.OnKey "^X"
                 Application.OnKey "{del}"
     End If
      If ows Is Nothing Then Set ows = ActiveSheet
      If wb.Worksheets("Sheet1").Cells(1, "E") = "Usrform" Then
          Set ows = wb.Worksheets("Sheet1")
          wb.Worksheets("Sheet1").Cells(1, "E") = ""
      End If
    Select Case ActiveSheet.Name
 
    Case "出入统计total", "机床性能巡检记录表", "日抽检表", "生产计划数据", "出入库单"
   
      ActiveSheet.AutoFilterMode = False
        If Cells(2, "A").Value = "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
        
         If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & Format(Cells(2, "A").Value, "YYYY/MM/DD")

         If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<" & Format(Cells(3, "A").Value, "YYYY/MM/DD")

          If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<" & "20" & datecode(Cells(3, "A").Value, "/")
          Application.GoTo Cells(5, "A"), Scroll:=True
          If ActiveSheet.Name <> "机床性能巡检记录表" Then
             If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
           End If
             
            
         
           
         Case "采购订单"
  
        
          ActiveSheet.AutoFilterMode = False
          If Cells(2, "A").Value = "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
        
         If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/")
         
         If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<" & "20" & datecode(Cells(3, "A").Value, "/")

          If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<" & "20" & datecode(Cells(3, "A").Value, "/")
      
           Application.GoTo Cells(5, "A"), Scroll:=True
   
   
    Case "外协订单"
          ActiveSheet.AutoFilterMode = False
          Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("W").Column, Criteria1:=""
          Application.GoTo Cells(5, "A"), Scroll:=True
          If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow - 20, "A"), Scroll:=True
    Case wsname
          ActiveSheet.AutoFilterMode = False
          Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
          ActiveSheet.Calculate
          Application.GoTo Cells(5, "A"), Scroll:=True
          If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
       
    Case "工序号码"
          ActiveSheet.AutoFilterMode = False
          Range("D4:Z" & Range("E1000000").End(xlUp).row).AutoFilter
          Application.GoTo Cells(5, "A"), Scroll:=True
          If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
        
     Case "送货情况"
          ActiveSheet.AutoFilterMode = False
          Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
          Application.GoTo Cells(5, "A"), Scroll:=True
         If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
    Case "加工情况"
         'If OWS.Name = "Err_list" Then
             'OWS.Activate
        ' Else
             ActiveSheet.AutoFilterMode = False
             Range("C4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<> "
             'Set a = Range("C5:C" & Range("C1000000").End(xlUp).row).find(Cells(1, "D").Value, LookIn:=xlValues, LookAt:=xlWhole)
             'If Not a Is Nothing Then
                'Application.Goto Cells(a.row, "A"), Scroll:=True
            'End If
            Application.GoTo Cells(5, "A"), Scroll:=True
           If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
        
    Case "仓库物资出入库明细"
          ActiveSheet.AutoFilterMode = False
          If Cells(2, "A").Value = "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
        
          If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/")
         If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<" & "20" & datecode(Cells(3, "A").Value, "/")

          If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<" & "20" & datecode(Cells(3, "A").Value, "/")
      Case "订单目录"
     
          
           ActiveSheet.AutoFilterMode = False
           ActiveSheet.Range("A4:AW" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("W").Column, Criteria1:="="
           Application.GoTo Cells(5, "A"), Scroll:=True
           If activerow <= Range("A1000000").End(xlUp).row And activerow > 20 Then Application.GoTo Cells(activerow, "A"), Scroll:=True
          protectsub
        
       Case "查询"
           zhikanqiandan
        Case "外协情况"
            ActiveSheet.AutoFilterMode = False
            ActiveSheet.Range("A3:Z3").AutoFilter Field:=Columns("A").Column, Criteria1:="<>"
            Application.GoTo Cells(activerow, "A"), Scroll:=True
       Case "Sheet1"
          On Error Resume Next
           ows.Activate
       Case "Err_list"
          
           ActiveSheet.Range("A4:L" & Range("A1000000").End(xlUp).row).AutoFilter
         Case "欠单加工情况"
           
          
            'ActiveSheet.Range("A4:Z4" & Range("A1000000").End(xlUp).row - 1).AutoFilter Field:=Columns("J").Column, Criteria1:="<>"
        
       Case Else
         If at("出入库总计", ActiveWorkbook.Name) > 0 Then
                      ActiveSheet.AutoFilterMode = False
                      ActiveSheet.Range("C4:Z" & Range("C1000000").End(xlUp).row).AutoFilter
                      GoTo Endp
         End If
           
                            If at("周计划表", ActiveWorkbook.Name) > 0 Then
                      ActiveSheet.AutoFilterMode = False
                      ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("H").Column, Criteria1:="="
                      GoTo Endp
           End If
           If at("原材料库存表", ActiveWorkbook.Name) > 0 Then
                      ActiveSheet.AutoFilterMode = False
                      ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("F").Column, Criteria1:="<>"
                      GoTo Endp
           End If
            If ActiveWorkbook.Name = "青塔产品成品检测报告.xlsm" Then
                          Worksheets("项目").Activate
             Else
               If ActiveWorkbook.Name = "图纸和作业标准和检验报告.xlsx" Then
                  ActiveSheet.Range("A3:Z3").AutoFilter
               Else
                   If ActiveSheet.ProtectScenarios = True Then
                      ActiveSheet.Unprotect Password:="jyc0908"
                    
                          ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
                          protectsub
                    Else
                       ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
                       
                    End If
               End If
             End If
             
        End Select
     
       
Endp:
 'Application.EnableEvents = True
  On Error GoTo 0
 '
  protectsub
  Application.EnableEvents = True
  数据jiagonggengxin_status = False
End Sub
Sub unprotectsub(Optional wsname = "", Optional init As Boolean = True)
Dim protect As Boolean
Dim pwd As String
 evt = Application.EnableEvents
  wbcheck
 Application.EnableEvents = False
  If init = True Then protectsub_initialize

 protect = False
 pwd = ""
  If wsname = "" Then wsname = ActiveSheet.Name
 
  Select Case wsname
     
       Case "出入统计total", "日抽检表", "送货情况", "订单目录", "原材料情况", "外协订单", "加工情况", "工序号码登记簿", "原材料顺序", "客户产品号码", "客户价格表", "工序号码"
            pwd = "jyc0908"
            protect = wb.Worksheets(wsname).ProtectScenarios
      Case "采购订单"
           pwd = "12789"
          protect = wb.Worksheets(wsname).ProtectScenarios
      Case Else
       On Error Resume Next
           pwd = "jyc0908"
            protect = wb.Worksheets(wsname).ProtectScenarios
     End Select
    
     
   
     If protect And pwd <> "" Then wb.Worksheets(wsname).Unprotect Password:=pwd
Endp:
 Application.EnableEvents = evt
End Sub

Sub protectsub_initialize()
'ReDim protect_array(30)
Dim ws As Worksheet
wbcheck
  i = 0
   
   For Each ws In wb.Worksheets
     If ws.ProtectScenarios = True Then
       protect_array(i) = ws.Name
       i = i + 1
     End If
   Next
   protect_array(i) = ""
   GoTo xx1:
   For crow = i To 30
       protect_array(crow) = ""
  Next
xx1:
   On Error Resume Next
    Debug.Print
    
Endp:
End Sub
Function check_protect(wsname)
Dim crow As Long
   check_protect = False
   On Error GoTo Endp:
   For crow = 0 To 30
     If protect_array(crow) = "" Then Exit For
     If protect_array(crow) = wsname Then
        check_protect = True
        Exit For
      End If
    Next

Endp:
End Function
Sub protectsub(Optional wsname = "", Optional init As Boolean)
On Error Resume Next
 evt = Application.EnableEvents
 Application.EnableEvents = False
 
  If wsname = "" Then wsname = ActiveSheet.Name
  If Application.UserName = "王鑫" Or Application.UserName = "周丽丽" Or Application.UserName = "jyc" Then GoTo Endp
If check_protect(wsname) = True Or init = True Then
 Select Case wsname
   Case "出入统计total"
         pwd = "jyc0908"
         wb.Worksheets(wsname).Unprotect Password:=pwd
         'wb.Worksheets(wsname).Range("A4:CZ" & Range("A1000000").End(xlUp).row).Locked = True
         'wb.Worksheets(wsname).Columns("E").Locked = False
        ' wb.Worksheets(wsname).Columns("I:M").Locked = False
        ' wb.Worksheets(wsname).Columns("T:BY").Locked = False
        ' wb.Worksheets(wsname).Cells(1, "L") = "不修改"
         If wb.Worksheets(wsname).ProtectScenarios = True Then GoTo Endp

   Case "订单目录"
       pwd = "jyc0908"
     'Range("A4:CZ" & Range("A100000").End(xlUp).Row).Locked = True
      If wb.Worksheets(wsname).ProtectScenarios = False Then wb.Worksheets(wsname).Columns("I").Locked = False
     
       'Cells(1, "H").Value = "不修改"
                If wb.Worksheets(wsname).ProtectScenarios = True Then GoTo Endp

     Case "日抽检表"
         pwd = "jyc0908"
        If wb.Worksheets(wsname).ProtectScenarios = False Then Rows("1:5").Locked = False
                 If wb.Worksheets(wsname).ProtectScenarios = True Then GoTo Endp

   Case "采购订单"
      pwd = "12789"
       If Worksheets(wsname).ProtectScenarios = False Then Rows("1:4").Locked = False
                If wb.Worksheets(wsname).ProtectScenarios = True Then GoTo Endp

   Case "出入统计total", "日抽检表", "送货情况", "订单目录", "原材料情况", "外协订单", "加工情况", "工序号码登记簿", "原材料顺序", "客户产品号码", "客户价格表", "工序号码"
           pwd = "jyc0908"
         If wb.Worksheets(wsname).ProtectScenarios = True Then GoTo Endp

   Case Else
      
  
   End Select
   
   If pwd <> "" Then wb.Worksheets(wsname).protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
   If init = True Then protectsub_initialize
 End If
Endp: Application.EnableEvents = evt
End Sub

Sub datechangefilter()
evt = Application.EnableEvents
 Application.EnableEvents = False

Dim x1 As String
unprotectsub
   If ActiveSheet.Name = "出入统计total" Then
     ' ActiveSheet.unprotect Password:="jyc0908"

            ActiveSheet.AutoFilterMode = False
       If Cells(2, "A").Value = "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter
        
         If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/")
         If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="<" & "20" & datecode(Cells(3, "A").Value, "/")

          If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<" & "20" & datecode(Cells(3, "A").Value, "/")
        ' ‘   If Cells(3, "B").Value <> "" Then ActiveSheet.Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter field:=2, Criteria1:=Cells(3, "Ｂ").Value
      ' ‘ ‘ If Cells(3, "C").Value <> "" Then ActiveSheet.Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter field:=3, Criteria1:=Cells(3, "C").Value
           
        
       ' If Cells(3, "D").Value <> "" Then ActiveSheet.Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter field:=4, Criteria1:=Cells(3, "D").Value
      
        
 End If
 
         
           'Cells(1, "K").Value = WorksheetFunction.Max(Range("K5:J20000"))
           Application.GoTo Cells(5, "A"), Scroll:=True
         '  ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
        protectsub
          Application.EnableEvents = evt
End Sub
Sub datechangefilter1(a)
 evt = Application.EnableEvents

 Application.EnableEvents = False

Dim x1 As String
unprotectsub
   If ActiveSheet.Name = "出入统计total" Or ActiveSheet.Name = "仓库物资出入库明细" Or ActiveSheet.Name = "机床性能巡检记录表" Or ActiveSheet.Name = "日抽检表" Then
    ' ActiveSheet.unprotect Password:="jyc0908"
            ActiveSheet.AutoFilterMode = False
         Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:="=" & datecode(Cells(1, "A").Value, "/")
         Application.GoTo Cells(5, "A"), Scroll:=True
 End If
 
    
           Application.GoTo Cells(5, "A"), Scroll:=True
         '  ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
          Application.EnableEvents = evt
protectsub
End Sub
Sub ctrlepress()
Application.EnableEvents = False
   rowmax = Range("A1000000").End(xlUp).row
   If ActiveSheet.Name = "加工情况" Then rowmax = Range("C1000000").End(xlUp).row
   If ActiveSheet.Name = "工序号码" Then rowmax = Range("E1000000").End(xlUp).row
    Application.GoTo Cells(rowmax, "A")
Application.EnableEvents = True
End Sub
Sub ctrlzpress()
Dim ws As Worksheet
''Dim wb As Workbook
  On Error GoTo Endp
  Application.EnableEvents = False

  wbcheck
 
   Select Case ActiveSheet.Name
  
     Case "出入统计total"
         crow = 200
         If wb.Worksheets("change").Cells(crow, "A") <> "" And wb.Worksheets("change").Cells(crow, "DA") <> "" Then
                'c.Row = wb.Worksheets("change").Cells(crow, "A")
                Set c = Columns("CZ").find(wb.Worksheets("change").Cells(crow, "DA"), LookIn:=xlValues, LookAt:=xlWhole)
                If Not c Is Nothing Then
                    Range("A" & c.row & ":" & "AA" & c.row).Value = wb.Worksheets("change").Range("B" & 2000 & ":AB" & 200).Value
                    Cells(c.row, "A").Select
                End If
                If Cells(c.row, "AA") = "" And Cells(c.row, "CZ") <> "" Then color Cells(c.row, "A"), 0
               
         End If
          wb.Worksheets("change").Range("A100:DA200").Copy Destination:=wb.Worksheets("change").Range("A101")
     Case Else
         
   End Select
Endp:
   On Error GoTo 0
  Application.EnableEvents = True
End Sub
Sub ctrlcpress()
       If ActiveSheet.Name <> "生产计划" Then
          Application.OnKey "^C"
           GoTo Endp
    End If
     Set selectedrange = Range(Selection.Address)
    ' Set selectedrange = currentselectedrange
     Erase delscjharray
Endp:
End Sub
Sub ctrlvpress()
Dim i As Long
Dim crow, col As Long
Dim scjhsh_rowmax As Long
Dim riqi As Date
Dim chongxindisp机床  As Boolean
On Error GoTo Endp
On Error GoTo 0
    wbcheck
    evt = Application.EnableEvents
     If ActiveSheet.Name <> "生产计划" Then
      Application.OnKey "^v"
      GoTo Endp
    End If
    Application.EnableEvents = False
            wb.Worksheets("生产计划数据").AutoFilterMode = False
   
    targetrow = Range(Selection.Address).row
    targetcolumn = Range(Selection.Address).Column
    If ActiveSheet.Name = "生产计划" And targetrow > 7 Then
       If delscjharray(1, 1) <> "" Then
           For crow = 1 To 100
              For col = 1 To 100
                  If delscjharray(crow, col) = "" Then Exit For
                  scjhsh_rowmax = wb.Worksheets("生产计划数据").Range("CZ65535").End(xlUp).row + 1
                     Set a = wb.Worksheets("生产计划数据").Columns("CZ").find(delscjharray(crow, col), LookIn:=xlValues, LookAt:=xlWhole)
                        If Not a Is Nothing Then
                             wb.Worksheets("生产计划数据").Rows(a.row).Copy Destination:=wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A")
                             If wb.ReadOnly = True Then color wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A"), 255
                             For i = targetrow + crow - 1 To targetrow + crow - 1 - 40 Step -1
                                 If Cells(i, "A") <> "" Then
                                     riqi = Cells(i, "A")
                                      Exit For
                                 End If
                              Next
                              wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A") = Format(riqi, "YYYY/MM/DD")
                             ' wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A") = wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A")
                              wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Q") = wb.Worksheets("生产计划").Cells(2, targetcolumn + col - 1)
                              wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "CZ") = Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & "-" & scjhsh_rowmax & Application.UserName  ' cong fu jian cha
                              wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Z") = ""
                              生产计划find位置 scjhsh_rowmax
                              erp_scjhsj_changed = True
                               scjharray(targetrow + crow - 1, targetcolumn - 1 + col) = scjhsh_rowmax
                                'If Cells(targetrow + crow - 1, targetcolumn - 1 + col) <> "" Then
                                    'del_生产计划数据 targetrow + crow - 1, targetcolumn - 1 + col
                               ' End If
                              '  wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Z") = targetrow + crow - 1 & "," & targetcolumn - 1 + col

                              ' Cells(targetrow + crow - 1, targetcolumn - 1 + col) = wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "X")
                        Else
                        
                        End If ' not a
               Next
           Next
exitfor:
          If chongxin_disp机床日期 = True Then
           If MsgBox("拷贝 好了 !! 可是 生产 计划 从新摆 (Y/N)", vbYesNo) = vbYes Then
              disp机床日期
           End If
          End If
          GoTo Endp
       End If
       
       For crow = 1 To selectedrange.Rows.Count
          For col = 1 To selectedrange.Columns.Count
             scjhsh_rowmax = wb.Worksheets("生产计划数据").Range("CZ65535").End(xlUp).row + 1
               Set a = wb.Worksheets("生产计划数据").Columns("Z").find(selectedrange(crow, col).row & "," & selectedrange(crow, col).Column, LookIn:=xlValues, LookAt:=xlWhole)
                If Not a Is Nothing Then
                'If scjharray(selectedrange(crow, col).row, selectedrange(crow, col).Column) = 0 Then GoTo next1
                    wb.Worksheets("生产计划数据").Rows(a.row).Copy Destination:=wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A")
                    If wb.ReadOnly = True Then color wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A"), 255
                    For i = targetrow + crow - 1 To targetrow + crow - 1 - 100 Step -1
                          If Cells(i, "A") <> "" Then
                             riqi = Cells(i, "A")
                             Exit For
                          End If
                    Next
                     
                   '  If Cells(targetrow + crow - 1, targetcolumn - 1 + col) <> "" Then
                                 '   del_生产计划数据 targetrow + crow - 1, targetcolumn - 1 + col
                       'End If
                                
                       wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "A") = Format(riqi, "YYYY/MM/DD")
                        wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Q") = wb.Worksheets("生产计划").Cells(2, targetcolumn + col - 1)
                        wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "CZ") = Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & "-" & scjhsh_rowmax & Application.UserName   ' cong fu jian cha
                        'wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Z") = targetrow + crow - 1 & "," & targetcolumn - 1 + col
                         wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "Z") = ""
                          scjharray(targetrow + crow - 1, targetcolumn - 1 + col) = scjhsh_rowmax
                          生产计划find位置 scjhsh_rowmax
                          erp_scjhsj_changed = True
                          'Cells(targetrow + crow - 1, targetcolumn - 1 + col) = wb.Worksheets("生产计划数据").Cells(scjhsh_rowmax, "X")
                     ' wb.Worksheets("生产计划数据").Rows(a.row).Copy Destination:=wb.Worksheets("sheet3").Cells(wb.Worksheets("sheet3").Range("X65535").End(xlUp).row + 1, "A")
                 End If ' a is
next1:
          Next
        Next
       ' selectedrange.Copy Destination:=Cells(targetrow, targetcolumn)
        If chongxin_disp机床日期 = True Then
           If MsgBox("拷贝 好了 !! 可是 生产 计划 从新摆 (Y/N)", vbYesNo) = vbYes Then
              disp机床日期
           End If
        End If
    Else
      On Error Resume Next
         Range(Selection.Address).Select
         ActiveSheet.Paste
         selectedrange.Copy Destination:=Cells(targetrow, targetcolumn)
      On Error GoTo 0
    End If
    
Endp:
   
   Application.EnableEvents = evt
End Sub

Sub ctrlxpress()
   If ActiveSheet.Name <> "生产计划" Then
      Application.OnKey "^X"
      GoTo Endp
    End If
    Application.EnableEvents = False
    Application.EnableEvents = False
    wbcheck
    On Error GoTo Endp
    On Error GoTo 0
    Erase delscjharray
    'wb.Worksheets("sheet3").Rows("5:" & wb.Worksheets("sheet3").Range("X65535").End(xlUp).row + 1).ClearContents
     Set currentselectedrange = Range(Selection.Address)
     For crow = 1 To currentselectedrange.Rows.Count
        For col = 1 To currentselectedrange.Columns.Count
            yclbz Cells(currentselectedrange(crow, col).row, currentselectedrange(crow, col).Column), False
            color wb.Worksheets("生产计划").Cells(currentselectedrange(crow, col).row, currentselectedrange(crow, col).Column), 0
             wb.Worksheets("生产计划").Cells(currentselectedrange(crow, col).row, currentselectedrange(crow, col).Column).Font.color = xlThemeColorLight1
            If currentselectedrange(crow, col) <> "" And currentselectedrange(crow, col).row > 7 Then
                delscjharray(crow, col) = currentselectedrange(crow, col).row & "," & currentselectedrange(crow, col).Column & "X"
               
                
                
                Set a = wb.Worksheets("生产计划数据").Columns("Z").find(currentselectedrange(crow, col).row & "," & currentselectedrange(crow, col).Column, LookIn:=xlValues, LookAt:=xlWhole)
                If Not a Is Nothing Then
                    delscjharray(crow, col) = wb.Worksheets("生产计划数据").Cells(a.row, "CZ")
                       
                       
                      If Application.UserName = "jyc" And wb.ReadOnly = False Then
                         'wb.Worksheets("生产计划数据").Rows(a.row).Delete Shift:=xlUp
                            wb.Worksheets("生产计划数据").Cells(a.row, "A") = ""
                            wb.Worksheets("生产计划数据").Cells(a.row, "AA") = Format(Date, "YYMMDD") & Format(Time, "HH:MM:SS") & Application.UserName
                            wb.Worksheets("生产计划数据").Cells(a.row, "Z") = wb.Worksheets("生产计划数据").Cells(a.row, "Z") & "X"
                            
                            
                      Else
                            wb.Worksheets("生产计划数据").Cells(a.row, "A") = ""
                             wb.Worksheets("生产计划数据").Cells(a.row, "AA") = Format(Date, "YYMMDD") & Format(Time, "HH:MM:SS") & Application.UserName
                            wb.Worksheets("生产计划数据").Cells(a.row, "Z") = wb.Worksheets("生产计划数据").Cells(a.row, "Z") & "X"
                            If wb.ReadOnly = True Then color wb.Worksheets("生产计划数据").Cells(a.row, "A"), 255
                            erp_scjhsj_changed = True
                     End If
                     'wb.Worksheets("生产计划数据").Rows(a.row).Copy Destination:=wb.Worksheets("sheet3").Cells(wb.Worksheets("sheet3").Range("X65535").End(xlUp).row + 1, "A")
                End If

               
             End If
        Next
     Next
    
      currentselectedrange.ClearContents
Endp:
   Application.EnableEvents = True
   
End Sub
Sub ctrlxpress1()
   If ActiveSheet.Name <> "生产计划" Then
      Application.OnKey "^X"
      GoTo Endp
    End If
    Application.EnableEvents = False
    Application.EnableEvents = False
    wbcheck
    On Error GoTo Endp
    On Error GoTo 0
    Erase delscjharray
    'wb.Worksheets("sheet3").Rows("5:" & wb.Worksheets("sheet3").Range("X65535").End(xlUp).row + 1).ClearContents
     Set currentselectedrange = Range(Selection.Address)
     For crow = 1 To currentselectedrange.Rows.Count
        For col = 1 To currentselectedrange.Columns.Count
            If currentselectedrange(crow, col) <> "" And currentselectedrange(crow, col).row > 7 Then
                delscjharray(crow, col) = currentselectedrange(crow, col).row & "," & currentselectedrange(crow, col).Column & "X"
                Set a = wb.Worksheets("生产计划数据").Columns("Z").find(currentselectedrange(crow, col).row & "," & currentselectedrange(crow, col).Column, LookIn:=xlValues, LookAt:=xlWhole)
                If Not a Is Nothing Then
                      If Application.UserName = "jyc" And wb.ReadOnly = False Then
                         'wb.Worksheets("生产计划数据").Rows(a.row).Delete Shift:=xlUp
                            wb.Worksheets("生产计划数据").Cells(a.row, "A") = ""
                            wb.Worksheets("生产计划数据").Cells(a.row, "AA") = Format(Date, "YYMMDD") & Format(Time, "HH:MM:SS") & Application.UserName
                            wb.Worksheets("生产计划数据").Cells(a.row, "Z") = wb.Worksheets("生产计划数据").Cells(a.row, "Z") & "X"
                            
                      Else
                            wb.Worksheets("生产计划数据").Cells(a.row, "A") = ""
                             wb.Worksheets("生产计划数据").Cells(a.row, "AA") = Format(Date, "YYMMDD") & Format(Time, "HH:MM:SS") & Application.UserName
                            wb.Worksheets("生产计划数据").Cells(a.row, "Z") = wb.Worksheets("生产计划数据").Cells(a.row, "Z") & "X"
                            If wb.ReadOnly = True Then color wb.Worksheets("生产计划数据").Cells(a.row, "A"), 255
                            erp_scjhsj_changed = True
                     End If
                     'wb.Worksheets("生产计划数据").Rows(a.row).Copy Destination:=wb.Worksheets("sheet3").Cells(wb.Worksheets("sheet3").Range("X65535").End(xlUp).row + 1, "A")
                End If

               
             End If
        Next
     Next
    
      currentselectedrange.ClearContents
Endp:
   Application.EnableEvents = True
   
End Sub


Sub f5jianyong()
Application.EnableEvents = False
    findws "加工情况", owsputh, "1"
    Application.EnableEvents = True
End Sub
Sub f5看查询()
Application.EnableEvents = False
    findws "查询", owsputh
    Application.EnableEvents = True
End Sub

Sub Fpresssub(wsname)
 
   wbcheck
Select Case ActiveWorkbook.Name
  
    Case "出入库单.xlsm"
        Set ows = ActiveSheet
            owsputh = wb.Worksheets("change").Cells(340, "C")
            pukehu = wb.Worksheets("change").Cells(340, "B")
            pugx = wb.Worksheets("change").Cells(340, "D")
            puycl = findyclbc(owsputh)
         '  If pugx <> "YCL" Then puycl = findyclbc(wb.Worksheets("change").Cells(340, "C"))
    Case "二楼库存表.xlsm"
    
     Set ows = ActiveSheet
     owsputh = wb.Worksheets("change").Cells(341, "C")
     pugx = "ZZ"
     puycl = findyclbc(wb.Worksheets("change").Cells(341, "C"))
  
    
    'Case "外协订单"
      'If wb.Worksheets("change").Cells(309, "B") = "YCL" Then
          'owsputh = ""
           'puycl = wb.Worksheets("change").Cells(309, "F")
       'Else
          'owsputh = wb.Worksheets("change").Cells(309, "F")
        '  puycl = ""
       'End If
        
    Case Else
      If at("mj.xlsm", ActiveWorkbook.Name) = 0 Then
          
            Set ows = ActiveSheet
           owsputh = wb.Worksheets("change").Cells(350, "C")
           pukehu = wb.Worksheets("change").Cells(350, "A")
           pugx = wb.Worksheets("change").Cells(350, "D")
           puycl = findyclbc(owsputh)
           If owsputh = "YCL" Or pugx = "YCL" Then
               puycl = wb.Worksheets("change").Cells(350, "F")
           End If
           move_stackarray -1
           j = 10
               stackarray(j, 0) = ActiveWorkbook.Name
               stackarray(j, 1) = ActiveSheet.Name
               stackarray(j, 2) = 0
               stackarray(j, 3) = owsputh
               stackarray(j, 4) = puycl
               stackarray(j, 5) = pugx
      End If
  End Select
  
 

  If ows Is Nothing Then Set ows = ActiveSheet
  On Error Resume Next
  Set ows = Workbooks(stackarray(10, 0)).Worksheets(stackarray(10, 1))
  owsputh = stackarray(10, 3)
  puycl = stackarray(10, 4)
  pugx = stackarray(10, 5)
  If ActiveSheet.Name = wsname Then
    
      If wsname = "加工情况" And ows.Name = wsname Then
         findws wsname, owsputh, "1"
       Else
         ows.Activate
        
        If ows.Name <> "生产计划" Then Application.GoTo Cells(stackarray(10, 2), "A"), Scroll:=True
          Rows(stackarray(10, 2)).Select
         
       End If
       On Error Resume Next
      ows.Activate
      GoTo Endp
     End If
   Select Case ows.Name
  
     
     
     
      Case "原材料情况"
           findws wsname, owsputh, pugx, True
     
         
       Case "欠单加工情况"
          findws wsname, owsputh, "", False
      
       Case "外协订单"
          
       
        
             findws wsname, owsputh, Worksheets("change").Cells(309, "D"), True
          
        Case "Err_list"
           
           findws wsname, owsputh, "GX", True
        
       Case "查询"
      
            findws wsname, owsputh, "", False
        Case "Sheet1"
             findws wsname, Worksheets("Sheet1").Cells(1, "H").Value, "", False

       Case Else
             findws wsname, owsputh, puycl, True
           
          
      End Select
Endp:
End Sub

Sub ALTf1press()
 Application.EnableEvents = False
 
  altpress = True
  Fpresssub "出入统计total"

Endp:
  Application.EnableEvents = True
  altpress = False
End Sub
Sub f2press()
 Application.EnableEvents = False
 On Error GoTo Endp
Dim ws As Worksheet

 Fpresssub "订单目录"
Endp:
   On Error GoTo 0
    Application.EnableEvents = True

End Sub

Sub f3press()
 Application.EnableEvents = False

Dim ws As Worksheet

   Fpresssub "整体计划"
   
Endp:
   Application.EnableEvents = True

End Sub

Sub f4press()
Dim ws As Worksheet
 Application.EnableEvents = False
 
 Fpresssub "原材料情况"
 
Endp:
   Application.EnableEvents = True

End Sub
Sub f5press()
    
 Application.EnableEvents = False
  Fpresssub "加工情况"
 
     
Endp:
     Application.EnableEvents = True
  
End Sub

Sub altf5press()
 Application.EnableEvents = False
 UserForm图号.Show
 
     
Endp:
     Application.EnableEvents = True
  
End Sub
Sub ALTf6press()
Application.EnableEvents = False
Dim ws As Worksheet
Dim awb As Workbook
On Error GoTo Endp:
On Error GoTo 0
 
  wbcheck
  Set awb = ActiveWorkbook
  Set aws = ActiveSheet
   If awb.Name = wb.Name Then
     'sheetjizhun ActiveSheet.Name
    
  End If
 If wb.Worksheets("sheet1").Cells(1, "L") = "" Then
   wb.Worksheets("sheet1").Cells(1, "AA") = ActiveWorkbook.Name
   wb.Sheets("sheet1").Cells(1, "AB") = ActiveSheet.Name
   
   Application.EnableEvents = True
   aa = Application.InputBox(prompt:="选择  美进图号 或则 客户 图号 ")
   Application.EnableEvents = False
     If aa <> "" And aa <> False Then
        If at("-", aa) > 0 Then
         puth = findtuhao(aa)
      Else
       puth = aa
      End If
      If at("MJ", VBA.UCase(wb.Name)) > 0 Then
         Set ows = ActiveSheet
         owsputh = puth
        findkehutuhaosform (puth)
     End If
    End If
 Else
    UserForm1.Show
 End If
Endp:
'awb.Activate
'aws.Activate
Application.EnableEvents = True
End Sub
Sub f6press()
 Application.EnableEvents = False
    Fpresssub "日抽检表"
  
Endp:
     Application.EnableEvents = True
  
End Sub
Sub f7press()
 Application.EnableEvents = False
 
  Fpresssub "外协订单"
     Application.EnableEvents = True
  
End Sub

Sub f8press()
 Application.EnableEvents = False
 
   Fpresssub "原材料"
     Application.EnableEvents = True
  
End Sub
Sub f9press()
 Application.EnableEvents = False
 
  
   Fpresssub "原材料顺序"
   Application.EnableEvents = True
End Sub

Sub f10press()
 Application.EnableEvents = False
 Dim ws As Worksheet
 
   Fpresssub "工序号码登记簿"
     Application.EnableEvents = True
  
End Sub
Sub f11press()
 Application.EnableEvents = False
 
   Fpresssub "工序号码"
     Application.EnableEvents = True
  
End Sub
Sub f12press()
 Application.EnableEvents = False

      Fpresssub "客户产品号码"
 Application.EnableEvents = True
  
End Sub
Sub altf12press()
 Application.EnableEvents = False
 
     If ActiveSheet.Name = "生产计划数据" Then
        Worksheets("生产计划").Activate
        'If activerow < 5 Then GoTo endP
       '  xstr = wb.Worksheets("生产计划数据").Cells(activerow, "Z")
         '  If xstr <> "" Then
                             '    x1 = Int(Val(Mid(xstr, 1, at(",", xstr) - 1)))
                                ' x2 = Int(Val(Mid(xstr, at(",", xstr) + 1, Len(xstr))))
                                ' wb.Worksheets("生产计划").Cells(x1, x2).Select
                                ' ' Application.Goto Cells(x1, x2), Scroll:=True
                                ' wb.Worksheets("生产计划").Rows(x1).Select
                                 
           'End If
     Else
        Fpresssub "生产计划数据"
     End If
Endp:
     Application.EnableEvents = True
  
End Sub

Sub altf10press()
   Application.EnableEvents = True
     Fpresssub "生产计划"
   
  
End Sub



Sub f7pressbuliang()
Dim mima As String
 evt = Application.EnableEvents
 Application.EnableEvents = False
  If Cells(1, "L").Value = "不能 修改" Then
 
  Else
    editmode = False

    If Cells(1, "M") < Time Then
     Cells(1, "L").Value = "不能 修改"
      color Cells(1, "L"), 3
      Cells(1, "M") = ""
       color Cells(1, "M"), 0
    End If
  End If
   Application.EnableEvents = evt


    
        
   
   Application.EnableEvents = True

End Sub

Sub f9pressmfjiagong()  ' BD clear
Dim a, b As Range
Dim i As Long
  evt = Application.EnableEvents
  
  
  wbcheck
'If ActiveWorkbook.Name <> "maching-flange-jiagong.xlsx" Then GoTo end1:
            
 Application.EnableEvents = False
  Set a = Range("O1:O" & Range("C1000000").End(xlUp).row).find(wb.owsputh & "-ZZ")
  If Not a Is Nothing Then
    For i = 0 To -20 Step -1
      If Cells(a.row + i, "C").Value <> wb.owsputh Then Exit For
        If Abs(Val(Cells(a.row + i, "AC")) - Val(Cells(a.row + i, "BM"))) > 5 And Cells(a.row + i, "E") = "" And Left(Cells(a.row + i, "F"), 1) <> "D" And Left(Cells(a.row + i, "F"), 1) <> "Q" And Left(Cells(a.row + i, "F"), 1) <> "W" Then Exit For
          Range(Cells(a.row + i, "AI"), Cells(a.row + i, "BD")).ClearContents
          Cells(a.row + i, "AI") = Cells(a.row + i, "AH") + Cells(a.row + i, "BO")
           Cells(a.row + i, "AH") = Cells(a.row + i, "AI")
           'Cells(a.Row + i, "BD") = ""
             
         
        Next
          If wbopencheck("mj.xlsm") > 0 Then
                    
                     banchengpingengxinxx20151113 a.row, True
                     
                  
                  Else
                    banchengpingengxinxx20151113 a.row, True
                   
                     
          End If
          If WorksheetFunction.SumIfs(Columns("BD"), Columns("C"), Cells(a.row, "C")) <> 0 Then
              For i = 0 To -20 Step -1
                 If Cells(a.row + i, "C").Value <> wb.owsputh Then Exit For
                 If Val(Cells(a.row + i, "BD")) <> 0 Then
         
                     Cells(a.row + i, "AI") = Cells(a.row + i, "AI") + Cells(a.row + i, "BD")
                     Cells(a.row + i, "AH") = Cells(a.row + i, "AH") + Cells(a.row + i, "BD")
                     Cells(a.row + i, "BD") = ""
                End If
             
            Next
          End If
    End If
    If Not a Is Nothing Then
         Application.GoTo Cells(a.row, "A"), Scroll:=True
    End If
end1:
 'Application.EnableEvents = True
 Application.EnableEvents = evt
End Sub

Sub f9pressmfjiagongmj20161226()
Dim a, b As Range
Dim i As Long
  evt = Application.EnableEvents
   
  wbcheck
'If ActiveWorkbook.Name <> "maching-flange-jiagong.xlsx" Then GoTo end1:
            
 Application.EnableEvents = False
  Set a = Range("O1:O" & Range("C1000000").End(xlUp).row).find(wb.owsputh & "-ZZ")
  If Not a Is Nothing Then
    For i = 0 To -20 Step -1
      If Cells(a.row + i, "C").Value <> wb.owsputh Then Exit For
        If Abs(Val(Cells(a.row + i, "AC")) - Val(Cells(a.row + i, "BM"))) > 5 And Cells(a.row + i, "E") = "" And Left(Cells(a.row + i, "F"), 1) <> "D" And Left(Cells(a.row + i, "F"), 1) <> "Q" And Left(Cells(a.row + i, "F"), 1) <> "W" Then Exit For
          Range(Cells(a.row + i, "AI"), Cells(a.row + i, "BD")).ClearContents
          Cells(a.row + i, "AI") = Cells(a.row + i, "AH") + Cells(a.row + i, "BO")
           Cells(a.row + i, "AH") = Cells(a.row + i, "AI")
         
             
         
        Next
          If wbopencheck("mj.xlsm") > 0 Then
                     Application.Run macro:=wb.Name & "!banchengpingengxinxx20151113", arg1:=a.row, arg2:=True
                     banchengpingengxinxx20151113 ", arg1:=Target.Row, arg2:=True"
                  Else
                   Application.Run macro:=wb.Name & "!banchengpingengxinxx20151113", arg1:=a.row, arg2:=True
                    banchengpingengxinxx20151113 ", arg1:=Target.Row, arg2:=True"
          End If
    End If
    If Not a Is Nothing Then
         If a.row > 20 Then Application.GoTo Cells(a.row, "A"), Scroll:=True - 20
    End If
end1:
 'Application.EnableEvents = True
 Application.EnableEvents = evt
End Sub
Sub f10pressjiagong()
Dim a, b As Range
Dim i As Long
  evt = Application.EnableEvents
            
 Application.EnableEvents = False
     For crow = 5 To Range("C1000000").End(xlUp).row

           
               Cells(crow + i, "DG") = Cells(crow, "DG") + Cells(crow, "BD") + Cells(crow + i, "AH").Value
               Cells(crow + i, "DB") = Cells(crow, "DB") + Cells(crow, "BD") + Cells(crow + i, "AH").Value
              
               Range(Cells(crow, "AH"), Cells(crow, "BD")).ClearContents
              
        Next
         ' If wbopencheck("mj.xlsm") > 0 Then
                     'Application.Run macro:=wb.name & "!banchengpingengxinxx", arg1:=a.Row
                 ' Else
                   'Application.Run macro:=wb.name & "!banchengpingengxinxx", arg1:=a.Row
         ' End If
   ' End If
    'If Not a Is Nothing Then
        ' If a.Row > 20 Then Application.Goto Cells(a.row, "A"),scroll:=true - 20
    'End If
end1:
 'Application.EnableEvents = True
 Application.EnableEvents = evt
End Sub



Function searchkey(key, wsname)
Dim a As Range
found = False
 With wb.Worksheets(wsname).Columns("X")
   Set a = .find(key, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious)
       If Not a Is Nothing Then
           firstrow = a.row
                                 Do
                                   
                                   If wb.Worksheets(wsname).Cells(a.row, "M") = "" Then
                                      found = True
                                      Exit Do
                                    End If
                                    Set a = .FindPrevious(a)
                                     If a Is Nothing Then Exit Do
                                     
                                Loop While firstrow > a.row
                          
            If found = False Then Set a = Cells(4, "A")
        End If
         End With
         
   If Not a Is Nothing Then
      searchkey = a.row
   Else
      searchkey = 4
   End If
End Function

Sub findws(wsname, key, Optional key1 = "", Optional cal = False)
Dim filterarray() As String
Dim a   As Range
Dim xa, xb As Long
  wbcheck
  If ows Is Nothing Then Set ows = ActiveSheet
     wb.Worksheets(wsname).Activate
    If wb.Worksheets(wsname).Visible = False Then wb.Worksheets(wsname).Visible = True
     
    If ows.Name = "原材料情况" And pugx = "YCL" Then
       colstr = ""
       Select Case ActiveSheet.Name
              
              Case "加工情况"
                    colstr = "C"
              Case "工序号码"
                    colstr = "E"
              Case "订单目录"
                    
                   ' colstr = "F"
                    colstr = ""
              Case Else
                    colstr = ""
           End Select
             If colstr <> "" Then
                      filterarray = chanpinjianyongfilterycl(key)
                       ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns(colstr).Column, Criteria1:=filterarray, Operator:=xlFilterValues
                       Application.GoTo Cells(5, "A"), Scroll:=True
                GoTo endselect:
             End If
    End If
    
     Select Case wsname
       
       Case "出入统计total", "日抽检表"
            found = False
            If wsname = "出入统计total" And altpress = False And pugx <> "YCL" Then
                 
                'wb.Worksheets(wsname).AutoFilterMode = False
                 'ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns("M").Column, Criteria1:=""
                 If pugx = "YCL" Then GoTo pugxycl
                 
                   If pugx <> "" And pugx <> "YCL" Then
                      
                       xb = 4
                        Select Case Left(pugx, 1)
                          
                        Case "S"
                           Set a = Cells(4, "A")
                           
                            xa = searchkey(key & "-" & "J", wsname)
                            If xa > xb Then xb = xa
                            xa = searchkey(key & "-" & "SZ", wsname)
                             If xa > xb Then xb = xa
                              xa = searchkey(key & "-" & pugx, wsname)
                              If xa > xb Then xb = xa
                               Set a = Cells(xb, "A")
                          Case "J"
                             Set a = Cells(4, "A")
                             If Left(pugx, 2) = "J1" Or Left(pugx, 2) = "J2" Then
                                  xb = searchkey(key & "-" & pugx, wsname)
                                  Set a = Cells(xb, "A")
                             End If
                             If a.row = 4 Then
                              
                                      xb = searchkey(key & "-" & "J", wsname)
                                      Set a = Cells(xb, "A")
                              
                              End If
                             
                           
                        Case Else
                            xb = searchkey(key & "-" & pugx, wsname)
                                 Set a = Cells(xb, "A")

                        End Select
                       
                      
                       If Not a Is Nothing Then found = True
                   End If
                   
                   If found = False And pugx <> "YCL" Then
                       With wb.Worksheets(wsname).Columns("C")
                       Set a = .find(key, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                          If Not a Is Nothing Then
                             firstrow = a.row
                           Do
                          
                              If wb.Worksheets(wsname).Cells(a.row, "M") = "" Then Exit Do
                              Set a = .FindPrevious(a)
                                If a Is Nothing Then Exit Do
                             Loop While firstrow > a.row
                         End If
                       End With
                    End If
                    
                  If Not a Is Nothing Then
                     If a.row <= 5 And ows.Name = "生产计划" Then
                       found = False
                        With wb.Worksheets(wsname).Columns("Q")
                         Set a = .find(wb.Worksheets("change").Cells(315, "C"), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                          If Not a Is Nothing Then
                             firstrow = a.row
                           Do
                          
                              If wb.Worksheets(wsname).Cells(a.row, "M") = "" And wb.Worksheets(wsname).Cells(a.row, "C") = key Then
                                 
                                 found = True
                                  Exit Do
                              End If
                              Set a = .FindPrevious(a)
                                If a Is Nothing Then Exit Do
                             Loop While firstrow > a.row
                         End If
                       End With
                       If found = True Then
                          Application.GoTo Cells(a.row, "A"), Scroll:=True
                          Rows(a.row).Select
                        End If
                     Else
                        Application.GoTo Cells(a.row, "A"), Scroll:=True
                        Rows(a.row).Select
                    End If
               Else
                    'MsgBox (wsname & "-" & key & " 找不到 ")
                    Application.GoTo Cells(Range("A1000000").End(xlUp).row, "A")
               End If
              GoTo endselect:
           
            End If
pugxycl:    If found = False And pugx = "YCL" Then
                      
                       Set a = wb.Worksheets(wsname).Columns("X").find(puycl & "-" & pugx, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                       
                          If a Is Nothing Then
                             Set a = wb.Worksheets(wsname).Columns("W").find(puycl, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                               If a Is Nothing Then Set a = Cells(4, "A")
                           End If
                          Application.GoTo Cells(a.row, "A"), Scroll:=True
                          Rows(a.row).Select
                       GoTo endselect:
                    End If
              If ows.Name = "原材料情况" Then
                       filterarray = chanpinjianyongfilterycl(key)
                       ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                       Application.GoTo Cells(5, "A"), Scroll:=True
                GoTo endselect:
            End If

               unprotectsub
                  
                  wb.Worksheets(wsname).AutoFilterMode = False
                '   If Cells(2, "A").Value <> "" And Cells(3, "A").Value = "" Then Range("A4:Z4").AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/")
                '   If Cells(2, "A").Value = "" And Cells(3, "A").Value <> "" Then Range("A4:Z4").AutoFilter Field:=1, Criteria1:="<" & "20" & datecode(Cells(3, "A").Value, "/")
        
                  '  If Cells(2, "A").Value <> "" And Cells(3, "A").Value <> "" Then Range("A4:Z4").AutoFilter Field:=1, Criteria1:=">=" & "20" & datecode(Cells(2, "A").Value, "/"), Criteria2:="<" & "20" & datecode(Cells(3, "A").Value, "/")
            
             If key1 = "" Then
                 'filterarray = chanpinjianyongfilter(key)
                 ReDim filterarray(0)
                 filterarray(0) = key
                ActiveSheet.Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                
                
             Else
                 If key1 = "YCL" Then
                    filterarray = chanpinjianyongfilterycl(key)
                      
                        ActiveSheet.Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                    
                 Else
                    If key1 = "GX" Then
                      filterarray = chanpinjianyongfilter(key)
                       ActiveSheet.Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                    Else
                       ActiveSheet.Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=key
                    End If
                 End If
             End If
         If cal = True And wsname = "出入统计total" Then
             crktjcal (False)
         End If
             protectsub
             Application.GoTo Cells(5, "A"), Scroll:=True
       Case "整体计划"
              wb.Worksheets(wsname).AutoFilterMode = False
              If pugx <> "YCL" Then
                Set a = Worksheets("整体计划").Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                If key = "YCL" Then Set a = Worksheets("整体计划").Columns("B").find(puycl, LookIn:=xlValues, LookAt:=xlWhole)
                          If Not a Is Nothing Then
                            Application.GoTo Cells(a.row, "A"), Scroll:=True
                            Rows(a.row).Select
                         Else
                             MsgBox (wsname & "-" & key & " Or " & key1 & " 找不到   ")
                        End If
                 Else
                       
                      
                        ActiveSheet.Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=Columns("B").Column, Criteria1:=puycl

                 End If
       Case "订单目录"
            If ows.Name = "原材料情况" Then
                       'filterarray = chanpinjianyongfilterycl(key)
                        ' ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("AF").Column, Criteria1:=key
                       Application.GoTo Cells(5, "A"), Scroll:=True
                GoTo endselect:
            End If
            If ows.Name <> "订单目录" Then
                  '‘wb.Worksheets("订单目录").AutoFilterMode = False
                  ' wb.Worksheets("订单目录").Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("W").Column, Criteria1:=""
             
            End If
           
           If pukehu = "" Then
             Set a = Worksheets(wsname).Columns("F").find(key, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
           Else
              Set a = Worksheets(wsname).Columns("AB").find(pukehu & key, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
           End If
           If pugx = "YCL" And puycl <> "" Then
              Set a = Worksheets(wsname).Columns("AF").find(puycl, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
           End If
              If Not a Is Nothing Then
                     Application.GoTo Cells(a.row, "A"), Scroll:=True
                        Rows(a.row).Select
                        
                Else
                       Set a = Worksheets(wsname).Columns("f").find(key, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                       If Not a Is Nothing Then
                         Application.GoTo Cells(a.row, "A"), Scroll:=True
                           Rows(a.row).Select
                        
                        Else
                            ActiveSheet.AutoFilterMode = False
                            Set a = Worksheets(wsname).Columns("E").find(key, LookIn:=xlValues, LookAt:=xlWhole, LookAt:=xlPrevious)
                            If Not a Is Nothing Then
                               Application.GoTo Cells(a.row, "A"), Scroll:=True
                               Rows(a.row).Select
                               MsgBox (wsname & "-" & key & " 已经 结束的订单 ")
                            Else
                             MsgBox (wsname & "-" & key & " Or " & key1 & " 没有未结束的订单 ")
                            End If
                        End If
                 End If
GoTo x1:
             unprotectsub
             ActiveSheet.AutoFilterMode = False
              filterarray = chanpinjianyongfilter(key)
             ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=filterarray, Operator:=xlFilterValues
              ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("W").Column, Criteria1:="="
            ' Application.Goto Cells(5, "A"),scroll:=true
             protectsub
x1:
        Case "原材料情况"
            filterarray = chanpinjianyongfilterycl(findyclbc(owsputh))
           
           ' If Mid(owsputh, Len(owsputh) - 1, 1) = "Z" Then filterarray(1) = Mid(owsputh, 1, Len(owsputh) - 1) & "*"
            atuofiltermode = False
            ActiveSheet.Range("A4:Z" & Range("A65535").End(xlUp).row).AutoFilter Field:=Columns("A").Column, Criteria1:=filterarray, Operator:=xlFilterValues
            GoTo endyclqk
            Set a = Worksheets(wsname).Columns("A").find(puycl, LookIn:=xlValues, LookAt:=xlWhole)
            
              If Not a Is Nothing Then
                 'Set a = Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, lookat:=xlWhole)
                Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                      
                Else
                             MsgBox (wsname & "-" & key & " Or " & findyclbc(key) & " YCL 找不到   ")
                 End If
endyclqk:
        Case "外协订单"
            ' unprotectsub
        '  GoTo x2:
            Set a = Nothing
            ActiveSheet.AutoFilterMode = False
             ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("W").Column, Criteria1:="="
             If key1 <> "" Then
               ReDim filterarray(2)
             filterarray(0) = key
             filterarray(1) = key1
                ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                Application.GoTo Cells(5, "A"), Scroll:=True
                GoTo xx2
             End If
             If dbclick = True Then
                 ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=owsputh
                  ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("D").Column, Criteria1:="<>YCL"
                  Application.GoTo Cells(5, "A"), Scroll:=True
                  GoTo xx2
             End If
            If puycl <> "" Then puycl = findyclbc(puth)
             Set a = Worksheets(wsname).Columns("F").find(puycl, LookIn:=xlValues, LookAt:=xlWhole)
              If Not a Is Nothing Then
                  'Application.Goto Cells(a.row, "A"),scroll:=true
                      '  Rows(a.row).Select
                   ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=puycl
                   Application.GoTo Cells(5, "A"), Scroll:=True
                       
                Else
                   Set a = Worksheets(wsname).Columns("F").find(owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                      If Not a Is Nothing Then
                         ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=owsputh
                       '  Application.Goto Cells(a.row + 10, "A")
                        ' Rows(a.row).Select
                      End If
                End If
                Application.GoTo Cells(5, "A"), Scroll:=True
                If a Is Nothing Then
             
                             MsgBox (wsname & "-" & puycl & " YCL Or " & owsputh & " 没有未结束的订单 ")
                 End If
x2:
              GoTo xx2
               ActiveSheet.AutoFilterMode = False
              ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("F").Column, Criteria1:=key
             ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("W").Column, Criteria1:="="
              
             protectsub
xx2:
         Case "查询"
          
              ActiveSheet.AutoFilterMode = False
              
               If key1 <> "YCL" Then
                     'chanpinjianyongfilter (key)
                     Set a = Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                     If Not a Is Nothing Then
                         Application.GoTo Cells(a.row, "A"), Scroll:=True
                        Rows(a.row).Select
                     End If
        
                Else
                   If key1 = "YCL" Then
                       filterarray = chanpinjianyongfilterycl(key)
                      
                        ActiveSheet.Range("A4:Z4").AutoFilter Field:=Columns("A").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                    End If
             End If
             
        Case "加工情况"
            If pugx = "YCL" Or ows.Name = "原材料情况" Then
                       filterarray = chanpinjianyongfilterycl(puycl)
                       ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns("C").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                       Application.GoTo Cells(5, "C")
                GoTo endselect:
            End If
              ActiveSheet.AutoFilterMode = False
              ' If key1 = "" Then GoTo xxx1:
              
                 '‘chanpinjianyongfilter (key)
xxx1:                Set a = wb.Worksheets(wsname).Columns("C").find(owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                 If Not a Is Nothing Then
                
                 Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
               End If
          Case "生产计划数据"
             ActiveSheet.AutoFilterMode = False
              If ows.Name = "生产计划" Then
                  Set a = wb.Worksheets("生产计划数据").Columns("Z").find(targetrow & "," & targetcolumn, LookIn:=xlValues, LookAt:=xlWhole)
                     If Not a Is Nothing Then
                        wb.Worksheets("生产计划数据").Activate
                        wb.Worksheets("生产计划数据").Rows(a.row).Select
                     End If
                   GoTo exit生产计划数据
             
               End If

               Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=">=" & Format(Date - 1, "YYYY/MM/DD")
               Range("A4:Z" & Range("a1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=owsputh
               
exit生产计划数据:
             
       Case "欠单加工情况"
          
              ActiveSheet.AutoFilterMode = False
             
        
              Set a = wb.Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
              If Not a Is Nothing Then
                 Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                  
             Else
                 MsgBox key & " 没有  欠单 "
             End If
      
            
            Case "工序号码登记簿"
                  If ows.Name = "原材料情况" And pugx = "YCL" Then
                       filterarray = chanpinjianyongfilterycl(key)
                       ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns("A").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                       Application.GoTo Cells(5, "A"), Scroll:=True
                        GoTo endselect:
                    End If

             
             If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
             If pugx = "YCL" Then
                 Set a = wb.Worksheets(wsname).Columns("I").find(puycl, LookIn:=xlValues, LookAt:=xlWhole)
              Else
                 Set a = wb.Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
              End If
              If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                 
             Else
                Set a = wb.Worksheets(wsname).Columns("E").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                Else
                  MsgBox key & " 工序号码登记簿  没有 "
                End If
              End If
             Case "工序号码"
                           If ows.Name = "原材料情况" Then
                       filterarray = chanpinjianyongfilterycl(key)
                       ActiveSheet.AutoFilterMode = False
                       ActiveSheet.Range("A4:Z" & Range("C1000000").End(xlUp).row).AutoFilter Field:=Columns("E").Column, Criteria1:=filterarray, Operator:=xlFilterValues
                       Application.GoTo Cells(5, "A"), Scroll:=True
                GoTo endselect:
            End If

                 wb.Worksheets(wsname).AutoFilterMode = False
              If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
              Set a = Worksheets(wsname).Columns("E").find(key, LookIn:=xlValues, LookAt:=xlWhole)
              If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "E")
                  Rows(a.row).Select
                 
             Else
                MsgBox key & " 工序号码  没有 "
              End If
              
             Case "原材料"
                Set a = Nothing
               If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
               Set a = Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
               
              If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                 
              Else
                MsgBox key & " 原材料  没有 "
              End If
              
            Case "原材料顺序"
              Set a = Nothing
              If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
             If pugx = "YCL" Then
                 Set a = Worksheets(wsname).Columns("A").find(puycl, LookIn:=xlValues, LookAt:=xlWhole)
             Else
                 Set a = Worksheets(wsname).Columns("A").find(findyclbc(owsputh), LookIn:=xlValues, LookAt:=xlWhole)
             
             End If
             
              If Not a Is Nothing Then
                 
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                 
              Else
                MsgBox puycl & " YCL or " & findyclbc(key) & "YCL  原材料顺序  没有 "
              End If
              Case "整体计划"
                  wb.Worksheets(wsname).AutoFilterMode = False
                   Set a = wb.Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                     If Not a Is Nothing Then
                        wb.Worksheets(wsname).Activate
                        wb.Worksheets(wsname).Rows(a.row).Select
                        Application.GoTo Cells(a.row, "A"), Scroll:=True
                     Else
                          MsgBox key & " 图号整体计划  没有 "
                     End If
                 
               Case "送货情况"
             
                  wb.Worksheets(wsname).AutoFilterMode = False
                   Set a = wb.Worksheets(wsname).Columns("A").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                     If Not a Is Nothing Then
                        wb.Worksheets(wsname).Activate
                        wb.Worksheets(wsname).Rows(a.row).Select
                        Application.GoTo Cells(a.row, "A"), Scroll:=True
                     Else
                          MsgBox key & " 送货计划  没有 "
                     End If
                Case "客户产品号码"
              If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
              
                 If pukehu <> "" Then
                     Set a = Worksheets(wsname).Columns("E").find(pukehu & owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                 Else
                 
                     Set a = Worksheets(wsname).Columns("K").find(owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                 End If
                
              If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                 
              Else
                 Set a = Worksheets(wsname).Columns("K").find(owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                 If Not a Is Nothing Then
                    Application.GoTo Cells(a.row, "A"), Scroll:=True
                    Rows(a.row).Select
                Else
                     Set a = Worksheets(wsname).Columns("B").find(key, LookIn:=xlValues, LookAt:=xlWhole)
                      If Not a Is Nothing Then
                          Application.GoTo Cells(a.row, "A"), Scroll:=True
                          Rows(a.row).Select
                 
                      Else
                           MsgBox key & " 产品客户号码  没有 "
                      End If
                  End If
              End If
              
              Case "客户价格表"
            
                If Worksheets(wsname).Visible = False Then Worksheets(wsname).Visible = True
              Set a = Worksheets(wsname).Columns("F").find(pukehu & owsputh, LookIn:=xlValues, LookAt:=xlWhole)
              If Not a Is Nothing Then
                  Application.GoTo Cells(a.row, "A"), Scroll:=True
                  Rows(a.row).Select
                 
              Else
                 Set a = Worksheets(wsname).Columns("E").find(owsputh, LookIn:=xlValues, LookAt:=xlWhole)
                  If Not a Is Nothing Then
                     Application.GoTo Cells(a.row, "A"), Scroll:=True
                     Rows(a.row).Select
                 
                 Else
              
                     res = MsgBox(owsputh & " 客户价格 没有 新输入 ? ", vbYesNo)
                     If res = vbYes Then
                        Application.GoTo Cells(Range("A1000000").End(xlUp).row + 1, "A")
                    
                     If key1 = "sheet1" Then
                         Cells(Range("A1000000").End(xlUp).row + 1, "A") = wb.Worksheets(key1).Cells(1, "J")
                         If wb.ReadOnly = True Then
                            color Cells(Range("A1000000").End(xlUp).row, "A"), 255
                             erp_khjg_changed = True
                          End If
                           Application.EnableEvents = True
                             Cells(Range("A1000000").End(xlUp).row, "B") = wb.Worksheets(key1).Cells(1, "K")
                               Application.EnableEvents = False
                             Unload UserForm1
                        End If
                   End If
                 End If
              End If
     End Select
endselect:
Set a = Nothing
End Sub

Sub ALT1press()
  Application.EnableEvents = False
Dim ws As Worksheet
Dim awb As Workbook
On Error GoTo Endp:
On Error GoTo 0


  
  
  
Endp:
Application.EnableEvents = True

End Sub

Sub ALTf2press()
Application.EnableEvents = False
Dim ws As Worksheet
Dim awb As Workbook
On Error GoTo Endp:
On Error GoTo 0
 Application.EnableEvents = False
  wbcheck
  Set awb = ActiveWorkbook
  Set aws = ActiveSheet
   If awb.Name = wb.Name Then
     'sheetjizhun ActiveSheet.Name
    
  End If
 If wb.Worksheets("sheet1").Cells(1, "L") = "" Then
   wb.Worksheets("sheet1").Cells(1, "AA") = ActiveWorkbook.Name
   wb.Sheets("sheet1").Cells(1, "AB") = ActiveSheet.Name
   
  '' Application.EnableEvents = True
   aa = Application.InputBox(prompt:="选择  美进图号 或则 客户 图号 ")
   Application.EnableEvents = False
     If aa <> "" And aa <> False Then
        If at("-", aa) > 0 Then
         puth = findtuhao(UCase(aa))
      Else
       puth = UCase(aa)
      End If
      If at("MJ", VBA.UCase(wb.Name)) > 0 Then
         Set ows = ActiveSheet
         owsputh = puth
        findkehutuhaosform (puth)
     End If
    End If
 Else
    UserForm1.Show
 End If
Endp:
'awb.Activate
'aws.Activate
Application.EnableEvents = True
End Sub
Sub f1press()
 Application.EnableEvents = False
  Application.EnableEvents = False

  altpress = False
  Fpresssub "出入统计total"


 
Endp:
Application.EnableEvents = True
End Sub
Sub ALTf3press()
Application.EnableEvents = False
wbcheck
 puth = owsputh
 pukehu = ""
 puddl = 0
 pujhq = 0
 If at("检验报告记录", ActiveWorkbook.path) = 0 Then
 
      ' If wb.Worksheets("sheet1").Cells(1, "J") <> "" Then GoTo xx1
       
     
         If at("mj.xlsm", ActiveWorkbook.Name) = 0 Then
             puth = wb.Worksheets("change").Cells(350, "C").Value
             pukehu = wb.Worksheets("change").Cells(350, "A")
             
         Else
          
              If ActiveSheet.Name = "生产计划" Or ActiveSheet.Name = "生产计划数据" Or ActiveSheet.Name = "订单目录" Then
                 
                  Set a = wb.Worksheets("订单目录").Columns("A").find(puddhm, LookIn:=xlValues, LookAt:=xlWhole)
                  If Not a Is Nothing Then
                    If wb.Worksheets("订单目录").Cells(a.row, "F") <> puth Then
                       
                       puddhm = ""
                   End If
                  Else
                     wb.Worksheets("sheet1").Cells(1, "K") = ""
                  End If
               ' If ActiveSheet.Name = "生产计划" Or ActiveSheet.Name = "生产计划数据" Then
                 ' wb.Worksheets("订单目录").AutoFilterMode = False
                '  wb.Worksheets("订单目录").Range("A4:AW" & wb.Worksheets("订单目录").Range("A1000000").End(xlUp).row).AutoFilter Field:=Columns("W").Column, Criteria1:="="
              ' End If
               'End If
             End If
              wb.Worksheets("sheet1").Cells(1, "K") = puddhm
              wb.Worksheets("sheet1").Cells(1, "J") = puth
       End If
xx1:         UserForm生产计划.Show
  Else
       Application.Run macro:=wb.Name & "!check_出厂检验记录"
  End If
Application.EnableEvents = True
End Sub
Sub ALTf4press()
Application.EnableEvents = False

  wbcheck
 Application.Run macro:=wb.Name & "!看图纸", arg1:=owsputh
Application.EnableEvents = True
End Sub

Sub ALTf4xxpress()
Application.EnableEvents = False
Dim aws, ws As Worksheet
Dim awb As Workbook
On Error GoTo Endp:
On Error GoTo 0
 
  wbcheck
  Set awb = ActiveWorkbook
  Set aws = ActiveSheet
  'Set aa = Application.InputBox(prompt:="选择  图号 ", Type:=8)
   aa = Application.InputBox(prompt:=" 图号 选择或则 写  ")
   If aa <> "" And aa <> False Then
      
         aa = findtuhao(aa)
    
     wb.Worksheets("出入统计total").Unprotect Password:="jyc0908"
     wb.Worksheets("出入统计total").Cells(3, "C") = aa
     wb.Activate
     Worksheets("出入统计total").Activate
     Application.Run macro:=wb.Name & "!看图纸", arg1:=wb.Worksheets("出入统计total").Cells(3, "C").Value, arg2:=wb.Worksheets("出入统计total").Cells(2, "M")
 End If

Endp:
On Error Resume Next
aws.Activate
Application.EnableEvents = True
End Sub
Sub ctrlQpress()
Application.EnableEvents = False
     Application.Calculation = xlCalculationManual
           Application.CalculateBeforeSave = False
          
                      ActiveWorkbook.Close savechanges = False
            
      
         
   
      'Application.CalculateBeforeSave = True
      Application.EnableEvents = evt
Application.EnableEvents = True
End Sub

Sub altxpress()
Dim owsold As Worksheet
Application.EnableEvents = False
 wbcheck
   If ActiveSheet.Name <> wb.Worksheets("change").Cells(2, "G").Value Then
      Set owsold = Workbooks(wb.Worksheets("change").Cells(2, "F").Value).Worksheets(wb.Worksheets("change").Cells(2, "G").Value)
       wb.Worksheets("change").Cells(3, "F") = ActiveWorkbook.Name
       wb.Worksheets("change").Cells(3, "g") = ActiveSheet.Name
   Else
       Set owsold = Workbooks(wb.Worksheets("change").Cells(3, "F").Value).Worksheets(wb.Worksheets("change").Cells(3, "G").Value)
   End If
  
   owsold.Activate
   
Application.EnableEvents = True
End Sub
Sub altSpress()
Application.EnableEvents = False
 
  wbcheck
  wb.Worksheets("change").Cells(2, "F") = ActiveWorkbook.Name
  wb.Worksheets("change").Cells(2, "g") = ActiveSheet.Name
Application.EnableEvents = True
End Sub

Sub kankanycltuhao()
 If fileopened("Z:\美进文当\shengchanbu\生产部\maching-flange-查询.xlsx", True) <> 0 Then GoTo Endp:
    Set wb1 = Workbooks("maching-flange-查询.xlsx")
  
       wb1.Worksheets("查询").Activate
       wb1.Worksheets("查询").AutoFilterMode = False
       wb1.Worksheets("查询").Range("A4:BZ4").AutoFilter Field:=3, Criteria1:="<>0"
       wb1.Worksheets("查询").Range("A4:BZ4").AutoFilter Field:=Columns("J").Column, Criteria1:=Worksheets("出入统计total").Cells(1, "C").Value
Endp:
End Sub
Sub buliangbiaogemove(xp)
Dim ws As Worksheet
On Error GoTo Endp:
 'If ActiveSheet.Name = "出入统计total" And ActiveWorkbook.Name = "mj.xlsm" Then
    Worksheets("日抽检表").Activate
      Worksheets("日抽检表").AutoFilterMode = False
     ' Worksheets("日抽检表").Range("A4:cZ4").AutoFilter Field:=3, Criteria1:=Worksheets("出入统计total").Cells(xp, "C").Value
     '  Worksheets("日抽检表").Range("A4:CZ4").AutoFilter Field:=2, Criteria1:=Worksheets("出入统计total").Cells(xp, "B").Value
     Set a = Columns("CZ").find(Sheets("出入统计total").Cells(xp, "CZ"), LookIn:=xlValues, LookAt:=xlWhole)
     If Not a Is Nothing Then
        Application.GoTo Cells(a.row, "A"), Scroll:=True
        Rows(a.row).Select
   
     Else
         MsgBox " 没有  不良记录  ", vbOKOnly
     End If
       '  Application.Goto Cells(a.row, "A"),scroll:=true
'Else
  ' Worksheets("出入统计total").Activate
'End If
Endp:
End Sub

Function chanpinjianyongfilter(th, Optional th1 As String)  'th1 zeng jiale weile gongxuhaoma xin zuo fangbian
Dim xx() As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
Dim evt As Boolean
Dim eth As String
''Dim wb As Workbook
On Error GoTo 0
evt = Application.EnableEvents
Application.EnableEvents = False
ReDim xx(0)

  
   wbcheck
   
    With wb.Worksheets("加工情况")
       Set rrj = Range(.Cells(1, "A"), .Cells(.Range("C1000000").End(xlUp).row, "CZ"))
       .AutoFilterMode = False
    End With
 
If th = "" Then GoTo Endp:
i = 0
If th1 <> "" Then
    xx(0) = th1
 
  i = 1
  ReDim Preserve xx(0 To 1)
End If
 
   xx(0) = th
    ReDim Preserve xx(0 To 1)
  GoTo end1
   Set a = rrj.Columns("C").find(th, LookIn:=xlValues, LookAt:=xlWhole)
    If a Is Nothing Then GoTo end1
        If rrj(a.row, "B") <> "" And Len(rrj(a.row, "B").Value) > 4 Then
           th = rrj(a.row, "B")
            xx(i) = th
            i = i + 1
         Else
            xx(i) = th
            i = i + 1
         End If
  
   
      
         Set a = rrj.Columns("B").find(th, LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
               firstaddress = a.row
               Do
                    If i = 0 Then
                       ReDim Preserve xx(0 To i)
                         xx(i) = rrj(a.row, "C").Value
                         i = i + 1
                    
                    Else
                      If xx(i - 1) <> rrj(a.row, "C").Value Then
                         ReDim Preserve xx(0 To i)
                         xx(i) = rrj(a.row, "C").Value
                         i = i + 1
                      End If
                    End If
                        Set a = rrj.Columns("B").findnext(a)
                        If a Is Nothing Then Exit Do
                        Debug.Print a.row, rrj(a.row, "C")
                Loop While firstaddress <> a.row
          End If
end1:
       
         chanpinjianyongfilter = xx
       '  GoTo endp:
        
       Select Case ActiveSheet.Name
           Case "加工情况"
            ' unprotectsub
              ActiveSheet.AutoFilterMode = False
                Range("C4:BZ" & Range("C1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
                'protectsub
          Case "工序号码"
            ' unprotectsub
                ActiveSheet.AutoFilterMode = False
                Range("E4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
             '   protectsub
           Case "查询"
                unprotectsub
                ActiveSheet.AutoFilterMode = False
                 Range("A4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
           Case "订单计划数据"
              unprotectsub
               ActiveSheet.AutoFilterMode = False
                 Range("A4:BZ4" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
                  protectsub
            Case "生产计划数据"
              
                 Range("A4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
            Case "出入统计total"
                  Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
            Case Else
              If ActiveWorkbook.Name = "各部门盘点数量" Then
                  ActiveSheet.AutoFilterMode = False
                  Range("A4:Z" & Range("A1000000").End(xlUp).row).AutoFilter Field:=2, Criteria1:=xx(), Operator:=xlFilterValues
              End If
           
       End Select
      
Endp:
  

     Application.EnableEvents = evt
     chanpinjianyongfilter = xx
End Function
Function chanpinkehufilter(kehu)  'kehu chanpin filter
Dim xx() As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
Dim evt As Boolean
Dim eth As String
''Dim wb As Workbook
On Error GoTo 0
evt = Application.EnableEvents
Application.EnableEvents = False
ReDim xx(0)

  
   wbcheck
   
    With wb.Worksheets("客户产品号码")
       Set rrj = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "AB"))
       .AutoFilterMode = False
    End With
 
If kehu = "" Then GoTo Endp:
i = 0
 
   For crow = 5 To wb.Worksheets("客户产品号码").Range("A1000000").End(xlUp).row
   
       If rrj(crow, "A") = kehu Then
           ReDim Preserve xx(0 To i)
            xx(i) = rrj(crow, "K")
            i = i + 1
         End If
  Next
   
      
      
Endp:
  

     Application.EnableEvents = evt
      chanpinkehufilter = xx
End Function
Function chanpinjianyongfilter200419(th, Optional th1 As String)  'th1 zeng jiale weile gongxuhaoma xin zuo fangbian
Dim xx() As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
Dim evt As Boolean
Dim eth As String
''Dim wb As Workbook
On Error GoTo 0
evt = Application.EnableEvents
Application.EnableEvents = False


  
   wbcheck
   
    With wb.Worksheets("加工情况")
       Set rrj = Range(.Cells(1, "A"), .Cells(.Range("C1000000").End(xlUp).row, "CZ"))
       .AutoFilterMode = False
    End With
   ReDim xx(0)
If th = "" Then GoTo Endp:
xx(0) = th
If th1 <> "" Then
  xx(1) = th1
 
  i = 2
   ReDim Preserve xx(0 To 2)

Else
   
  i = 1
  ReDim Preserve xx(0 To i)
End If
   Set a = rrj.Columns("C").find(th, LookIn:=xlValues, LookAt:=xlWhole)
    If Not a Is Nothing Then
        If rrj(a.row, "E") <> "" And Len(rrj(a.row, "E").Value) > 4 Then
          th = rrj(a.row, "E")
            
            xx(i) = th
          
           i = i + 1
           ReDim Preserve xx(0 To i)
        End If
    End If
   
      
         Set a = rrj.Columns("E").find(th, LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
               firstaddress = a.row
               Do
                 
                 For crow = 0 To i
                    If rrj(a.row, "C") = xx(crow) Then Exit For
                    If crow = i Then
                       'Debug.Print rrj(a.row, "A")
                        
                       xx(i) = rrj(a.row, "C").Value
                      
                       
                        i = i + 1
                        ReDim Preserve xx(0 To i)
                       Exit For
                     End If
                 Next
                 
                 
                   
                 Set a = rrj.Columns("E").findnext(a)
                If a Is Nothing Then Exit Do
                  Debug.Print a.row, rrj(a.row, "C")
                Loop While firstaddress <> a.row
          End If
          Set a = rrj.Columns("C").find(th, LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
              For crow = 0 To 15
                 If rrj(a.row + crow, "C").Value <> th Then Exit For
                  If rrj(a.row + crow, "E").Value <> "" And rrj(a.row + crow, "E").Value <> "X" Then
                      For xcrow = 0 To i
                         If rrj(a.row, "C") = xx(xcrow) Then Exit For
                           If xcrow = i Then
                               'Debug.Print rrj(a.row, "A")
                                 
                                  xx(i) = rrj(a.row, "C").Value
                                    
                                   i = i + 1
                                   ReDim Preserve xx(0 To i)
                               Exit For
                           End If
                      Next
                     
                 End If
             Next
            End If
       
        
       '  GoTo endp:
        
       Select Case ActiveSheet.Name
           Case "加工情况"
            ' unprotectsub
              ActiveSheet.AutoFilterMode = False
                Range("C4:BZ" & Range("C1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
                'protectsub
          Case "工序号码"
            ' unprotectsub
                ActiveSheet.AutoFilterMode = False
                Range("E4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
             '   protectsub
           Case "查询"
                unprotectsub
                ActiveSheet.AutoFilterMode = False
                 Range("A4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
           Case "订单计划数据"
              unprotectsub
               ActiveSheet.AutoFilterMode = False
                 Range("A4:BZ4" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
                  protectsub
            Case "生产计划数据"
              
                 Range("A4:BZ" & Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
                 
            
           ' Case "出入统计total"
               '  Range("A4:Z"& Range("A1000000").End(xlUp).row).AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
       End Select
      If xx(i) = "" Then ReDim Preserve xx(0 To i - 1)
       chanpinjianyongfilter = xx
Endp:
  
  Set rrj = Nothing
     Application.EnableEvents = evt
     chanpinjianyongfilter = xx
End Function


Function chanpinjianyongfilterycl(yclbc)
Dim xx() As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
''Dim wb As Workbook
'On Error Resume Next
evt = Application.EnableEvents
Application.EnableEvents = False
 
  wbcheck
   'unprotectsub
   
If yclbc = "" Then GoTo Endp:
ReDim xx(0)
xx(0) = yclbc

'ReDim xx(0 To 1)
i = 1
        wb.Worksheets("工序号码登记簿").AutoFilterMode = False
         Set a = wb.Worksheets("工序号码登记簿").Columns("I").find(yclbc, LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
               firstaddress = a.row
               Do
                
                  'Err.Clear
                ' x1 = WorksheetFunction.Match(rrj(a.row, "A").Value, xx, 0)
               
                        ReDim Preserve xx(0 To i)
                        xx(i) = wb.Worksheets("工序号码登记簿").Cells(a.row, "A")
                        i = i + 1
                  
                  ' End If
                   
                    
             
                 Set a = wb.Worksheets("工序号码登记簿").Columns("I").findnext(a)
                Loop While Not a Is Nothing And firstaddress <> a.row
          End If
        
       chanpinjianyongfilterycl = xx
Endp:
  'protectsub
     Application.EnableEvents = evt
End Function


Sub chanpinjianyongfilter1(th)
'
Dim xx(100) As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
''Dim wb As Workbook
evt = Application.EnableEvents
Application.EnableEvents = False

wbcheck
 ActiveSheet.Unprotect Password:="jyc0908"
    With wb.Worksheets("其他数据")
       Set rrq = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "CZ"))
   End With
If th = "" Then Exit Sub
xx(0) = th
i = 1
         Set a = rrq.Columns("B").find(th, LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
               firstaddress = a.row
               Do
                  For Each j In xx
                      If xx(j) = th Then Exit For
                      If xx(j) = "" Then
                         xx(j) = findtuhao(rrq(a.row, "A").Value)
                         i = j + 1
                         Exit Sub
                      End If
                  Next
                End If
                 Set a = rrq.Columns("B").findnext(a)
                Loop While Not a Is Nothing And firstaddress <> a.row
          End If
          Set a = rrq.Columns("A").find(th & "-A", LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
              For crow = 0 To 10
                 If findtuhao(rrq(a.row + crow, "A").Value) <> th Then Exit For
                  If rrq(a.row + crow, "B").Value <> "" And rrq(a.row + crow, "B").Value <> "X" Then
                     For Each j In xx
                      If xx(j) = findtuhao(rrq(a.row + crow, "B").Value) Then Exit For
                      If xx(j) = "" Then
                         xx(j) = findtuhao(rrq(a.row + crow, "B").Value)
                          i = j + 1
                         Exit Sub
                      End If
                  Next
                  End If
             Next
            End If
         For crow = i To 100
            xx(crow) = "XXX"
         Next
         
     Range("A4:BZ4").AutoFilter Field:=1, Criteria1:=xx(), Operator:=xlFilterValues
    ' ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
    Application.EnableEvents = evt
    Set rrq = Nothing
End Sub
Sub chazhaojianyonggx()
'
Dim xx(1000), xstr As String
Dim crow, i As Long
Dim a, b, c As Range
''Dim wb As Workbook
Dim firstaddress  As Long
evt = Application.EnableEvents
Application.EnableEvents = False
  ActiveSheet.Unprotect Password:="jyc0908"
   If fileopened("Z:\美进文当\shengchanbu\生产部\maching-flange-查询.xlsx", True) <> 0 Then GoTo Endp:
    Set wb1 = Workbooks("maching-flange-查询.xlsx")

  wbcheck
    With wb.Worksheets("其他数据")
       Set rrq = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "CZ"))
   End With
 With wb1.Worksheets("查询")
       Set rrc = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "CZ"))
   End With

  i = 0
    If Cells(3, "C").Value <> "" And Cells(3, "D") <> "" Then
        xx(i) = Cells(3, "C").Value
        i = i + 1
         Set a = rrq.Columns("A").find(Cells(3, "C").Value & "-" & Cells(3, "D").Value, LookIn:=xlValues, LookAt:=xlWhole)
         If Not a Is Nothing Then
             If rrq(a.row, "B").Value <> "" And rrq(a.row, "B").Value <> "X" Then
                xx(i) = rrq(a.row, "B").Value
                 i = i + 1
                xstr = rrq(a.row, "B").Value & "-" & Cells(3, "D").Value
                 Set a = rrq.Columns("B").find(xstr, LookIn:=xlValues, LookAt:=xlWhole)
                  If Not a Is Nothing Then
                     firstaddress = a.row
                      Do
                       xx(i) = findtuhao(rrq(a.row, "A").Value)
                       i = i + 1
                       Set a = rrq.Columns("B").findnext(a)
                     Loop While Not a Is Nothing And firstaddress <> a.row
                  End If
           Else
                  'xstr = Cells(3, "C").Value & "-" & Cells(3, "D").Value
                  Set a = rrq.Columns("B").find(Cells(3, "C").Value, LookIn:=xlValues, LookAt:=xlWhole)
                  If Not a Is Nothing Then
                     firstaddress = a.row
                      Do
                          For j = 0 To 20
                             If rrq(a.row + j, "B").Value = "" Or rrq(a.row + j, "B").Value = "X" Then Exit For
                             If findgongxu(rrq(a.row + j, "A").Value) = Cells(3, "D").Value Then
                                xx(i) = findtuhao(rrq(a.row + j, "A").Value)
                                 i = i + 1
                                End If
                             Next
                       Set a = rrq.Columns("B").findnext(a)
                     Loop While Not a Is Nothing And firstaddress <> a.row
                  End If
           
           End If
      End If
      Else
        If Cells(3, "C").Value <> "" And Cells(3, "D") = "" Then
            xx(i) = Cells(3, "C").Value
            i = i + 1
            Set a = rrc.Columns("A").find(Cells(3, "C").Value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not a Is Nothing Then
                For j = 0 To 20
                  If rrc(a.row + j, "C").Value = "ZZ" Then Exit For
                  If rrc(a.row + j, "B").Value <> "" And rrq(a.row + j, "B").Value <> "X" Then
                     xx(i) = rrc(a.row + j, "B").Value
                      i = i + 1
                      xstr = rrc(a.row + j, "B").Value
                      Set b = rrc.Columns("B").find(xstr, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not a Is Nothing Then
                               firstaddress = b.row
                              Do
                               xx(i) = rrc(b.row, "A").Value
                               i = i + 1
                                Set a = rrc.Columns("B").findnext(b)
                                Loop While Not a Is Nothing And firstaddress <> b.row
                         End If
                     Else
                          Exit For
                      End If
                 Next
                  Set b = rrc.Columns("B").find(Cells(3, "C").Value, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not b Is Nothing Then
                            firstaddress = b.row
                            Do
                               xx(i) = rrc(b.row, "A").Value
                               i = i + 1
                                Set a = rrc.Columns("B").findnext(b)
                                Loop While Not a Is Nothing And firstaddress <> b.row
                         End If
              Else
     
           End If
      End If
   End If
    'If Cells(2, "D").Value <> "" Then
        ' Set a = rrq.Columns("A").Find(Cells(3, "C").Value & "-" & Cells(2, "D").Value, LookIn:=xlValues, LookAt:=xlWhole)
        ' If Not a Is Nothing Then
          '   If rrq(a.Row, "B").Value <> "" And rrq(a.Row, "B").Value <> "X" Then
           '     xx(i) = rrq(a.Row, "B").Value
           '      i = i + 1
             '   xstr = rrq(a.Row, "B").Value & "-" & Cells(2, "D").Value
              '   Set a = rrq.Columns("B").Find(xstr, LookIn:=xlValues, LookAt:=xlWhole)
              '    If Not a Is Nothing Then
                '     firstaddress = a.Row
                '      Do
                  '     xx(i) = findtuhao(rrq(a.Row, "A").Value)
                  '     i = i + 1
                  '     Set a = rrq.Columns("B").FindNext(a)
                  '   Loop While Not a Is Nothing And firstaddress <> a.Row
                  'End If
       '    End If
     ' End If
   ' End If



         For crow = i To 1000
            xx(crow) = "XXX"
         Next
         
     Range("A4:BZ4").AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
    ' ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
Endp:
Set rrq = Nothing
Set rrc = Nothing
Application.EnableEvents = evt
End Sub
Function chaxunjianyonggontuhao(xx, th, gx)
evt = Application.EnableEvents
Application.EnableEvents = False
 ActiveSheet.Unprotect Password:="jyc0908"
  For i = 0 To 1000
    If xx(i) = "" Then Exit For
  Next
        xx(i) = th
         Set a = rrc.Columns("A").find(th, LookIn:=xlValues, LookAt:=xlWhole)
            If Not a Is Nothing Then
                For j = 0 To 20
                  If rrc(a.row + j, "C").Value = "ZZ" Then Exit For
                  If rrc(a.row + j, "B").Value <> "" And rrq(a.row + j, "B").Value <> "X" Then
                     xx(i) = rrc(a.row + j, "A").Value
                      i = i + 1
                      xstr = rrc(a.row + j, "B").Value
                      Set b = rrc.Columns("B").find(xstr, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not a Is Nothing Then
                               firstaddress = b.row
                              Do
                               xx(i) = rrc(b.row, "A").Value
                               i = i + 1
                                Set a = rrc.Columns("B").findnext(b)
                                Loop While Not a Is Nothing And firstaddress <> b.row
                         End If
                     Else
                          Exit For
                      End If
                 Next
    chaxunjianyonggontuhao = xx
    'ActiveSheet.Protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
  Application.EnableEvents = evt
  Set rrc = Nothing
End Function

Sub chanpinjianyongfilterjiagong(Optional tuhao As String = "")
Dim xx(1000) As String
Dim crow, i As Long
Dim a, b, c As Range
Dim firstaddress  As Long
evt = Application.EnableEvents
Application.EnableEvents = False
If tuhao = "" Then
 th = Cells(1, "D").Value
Else
  th = tuhao
End If
 If wbopencheck("mj.xlsm") > 0 Then
    With wb.Worksheets("其他数据")
       Set rrq = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "CZ"))
End With
Else
    With wb.Worksheets("其他数据")
       Set rrq = Range(.Cells(1, "A"), .Cells(.Range("A1000000").End(xlUp).row, "CZ"))
End With
End If
xx(0) = th
i = 1
    Set a = rrq.Columns("A").find(th & "-A", LookIn:=xlValues, LookAt:=xlWhole)
    If Not a Is Nothing Then
        If rrq(a.row, "E") <> "" And Len(rrq(a.row, "B").Value) > 4 Then th = rrq(a.row, "B")
        xx(1) = th
        i = i + 1
    End If
         Set a = rrq.Columns("B").find(th, LookIn:=xlValues, LookAt:=xlWhole)
           If Not a Is Nothing Then
             
               firstaddress = a.row
               Do
                  xx(i) = findtuhao(rrq(a.row, "A").Value)
                  i = i + 1
                 Set a = rrq.Columns("B").findnext(a)
                Loop While Not a Is Nothing And firstaddress <> a.row
          End If
          Set a = rrq.Columns("A").find(th & "-A", LookIn:=xlValues, LookAt:=xlWhole)
          If Not a Is Nothing Then
              For crow = 0 To 10
                 If findtuhao(rrq(a.row + crow, "A").Value) <> th Then Exit For
                  If rrq(a.row + crow, "B").Value <> "" And rrq(a.row + crow, "B").Value <> "X" Then
        
                     xx(i) = findtuhao(rrq(a.row + crow, "B").Value)
                                  i = i + 1
                  End If
             Next
            End If
        
         For crow = i To 1000
            xx(crow) = "XXX"
         Next
    'If ActiveSheet.ProtectScenarios = True Then
      'protect = True
      'ActiveSheet.unprotect Password:="jyc0908"
  ' End If
      
     ActiveSheet.AutoFilterMode = False
     Range("A4:BZ4").AutoFilter Field:=3, Criteria1:=xx(), Operator:=xlFilterValues
    ' If protect = True Then ActiveSheet.protect Password:="jyc0908", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
   Application.EnableEvents = evt
   Set rrq = Nothing
End Sub



Sub lastcolumn()
 evt = Application.EnableEvents
  Application.EnableEvents = False
    colstr = "A"
  If ActiveSheet.Name = "加工情况" Then colstr = "C"
   If ActiveSheet.Name = "工序号码" Then colstr = "E"
    Cells(Range(colstr & "1000000").End(xlUp).row + 1, "A").Select
 Application.EnableEvents = evt
End Sub

Sub crktjcal(calok)
On Error Resume Next
 'Cells(2, "F").Value = WorksheetFunction.Subtotal(109, Range("f5:F" & Range("A1000000").End(xlUp).Row))
            ' Cells(2, "G").Value = WorksheetFunction.Subtotal(109, Range("g5:g" & Range("A1000000").End(xlUp).Row))
            ' Cells(2, "H").Value = WorksheetFunction.Subtotal(109, Range("h5:h" & Range("A1000000").End(xlUp).Row))
           '  Cells(2, "I").Value = Cells(2, "G") - Cells(2, "F") - Cells(2, "H")
             'Cells(3, "I").Value = WorksheetFunction.Subtotal(109, Range("I5:I" & Range("A1000000").End(xlUp).Row))
             Cells(2, "t").Value = WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("B"), ">ZZZZ", Columns("M"), "")
            ' Cells(3, "N").Value = WorksheetFunction.Subtotal(109, Range("n5:n" & Range("A1000000").End(xlUp).Row))
            ' Cells(3, "O").Value = WorksheetFunction.Subtotal(109, Range("O5:O" & Range("A1000000").End(xlUp).Row))
            ' Cells(3, "U").Value = WorksheetFunction.Subtotal(109, Range("u5:u" & Range("A1000000").End(xlUp).Row))
          
          '  Cells(3, "R").Value = WorksheetFunction.Subtotal(109, Range("r5:r" & Range("A1000000").End(xlUp).Row))
        If calok = True Then
              Cells(1, "S").Value = WorksheetFunction.SumIfs(Columns("R"), Columns("A"), Cells(1, "A"), Columns("Q"), ">=A")
              Cells(2, "N").Value = WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "J*")
              Cells(2, "O").Value = WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "S*")
               Cells(2, "P").Value = WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "Z*") + WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "Y*") + WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "X*") + WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "A*")
                Cells(1, "t").Value = Cells(2, "N") + Cells(2, "O") + Cells(2, "P")
               Cells(2, "R").Value = WorksheetFunction.SumIfs(Columns("T"), Columns("A"), Cells(1, "A"), Columns("Q"), "GY*")
               Cells(2, "Q").Value = Cells(2, "T") - (Cells(2, "N") + Cells(2, "O") + Cells(2, "P") + Cells(2, "R"))
               
              Cells(1, "s").Value = Format(Cells(1, "t") / Cells(1, "s"), "##.##") & " RMB/HR "
        End If
           
           '  Cells(2, "U").Value = WorksheetFunction.SumIfs(Columns("G"), Columns("A"), Cells(1, "A"), Columns("B"), ">ZZZZ", Columns("M"), "")
            '  Cells(1, "U").Value = WorksheetFunction.SumIfs(Columns("G"), Columns("A"), Cells(1, "A"), Columns("B"), ">ZZZZ", Columns("Q"), ">=A")
             ' Cells(1, "V").Value = Cells(1, "T") / Cells(1, "U")
            '  Cells(2, "V").Value = Cells(2, "T") / Cells(2, "U")
             Cells(3, "Y").Value = WorksheetFunction.Subtotal(109, Range("y5:y" & Range("A1000000").End(xlUp).row))
               Cells(3, "Z").Value = WorksheetFunction.Subtotal(109, Range("Z5:Z" & Range("A1000000").End(xlUp).row))
               Cells(2, "F").Value = WorksheetFunction.Subtotal(109, Range("f5:F" & Range("A1000000").End(xlUp).row))
                 Cells(3, "J").Value = WorksheetFunction.Subtotal(109, Range("J5:J" & Range("A1000000").End(xlUp).row))
         Cells(3, "K").Value = WorksheetFunction.Subtotal(105, Range("K5:K" & Range("A1000000").End(xlUp).row))
          
         'Cells(3, "L").Value = WorksheetFunction.Subtotal(105, Range("l5:l"& Range("A1000000").End(xlUp).Row))
          Cells(3, "R").Value = WorksheetFunction.Subtotal(109, Range("r5:r" & Range("A1000000").End(xlUp).row))
          Cells(3, "N").Value = WorksheetFunction.Subtotal(109, Range("n5:n" & Range("A1000000").End(xlUp).row))
          Cells(3, "O").Value = WorksheetFunction.Subtotal(109, Range("O5:O" & Range("A1000000").End(xlUp).row))
          Cells(3, "U").Value = WorksheetFunction.Subtotal(109, Range("u5:u" & Range("A1000000").End(xlUp).row))
          Cells(3, "V").Value = WorksheetFunction.Subtotal(109, Range("v5:v" & Range("A1000000").End(xlUp).row))
          Cells(3, "F").Value = WorksheetFunction.Subtotal(109, Range("f5:F" & Range("A1000000").End(xlUp).row))
        Cells(3, "G").Value = WorksheetFunction.Subtotal(109, Range("g5:g" & Range("A1000000").End(xlUp).row))
        Cells(3, "H").Value = WorksheetFunction.Subtotal(109, Range("h5:h" & Range("A1000000").End(xlUp).row))
        ' Cells(2, "I").Value = Cells(3, "G") - Cells(3, "F") - Cells(3, "H")
        Cells(3, "I").Value = WorksheetFunction.Subtotal(109, Range("I5:I" & Range("A1000000").End(xlUp).row))
         Cells(3, "J").Value = WorksheetFunction.Subtotal(109, Range("J5:J" & Range("A1000000").End(xlUp).row))
         Cells(3, "K").Value = WorksheetFunction.Subtotal(105, Range("K5:K" & Range("A1000000").End(xlUp).row))
          
        
          Cells(3, "R").Value = WorksheetFunction.Subtotal(109, Range("r5:r" & Range("A1000000").End(xlUp).row))
           Cells(3, "N").Value = WorksheetFunction.Subtotal(109, Range("n5:n" & Range("A1000000").End(xlUp).row))
            Cells(3, "O").Value = WorksheetFunction.Subtotal(109, Range("O5:O" & Range("A1000000").End(xlUp).row))
             Cells(3, "U").Value = WorksheetFunction.Subtotal(109, Range("u5:u" & Range("A1000000").End(xlUp).row))
            Cells(3, "t").Value = WorksheetFunction.Subtotal(109, Range("t5:t" & Range("A1000000").End(xlUp).row))
             Cells(2, "s").Value = Int(Cells(3, "t").Value / Cells(3, "r").Value) & "." & Int(Cells(3, "t").Value / Cells(3, "r").Value * 100) - Int(Cells(3, "t").Value / Cells(3, "r")) * 100 & " RMB/HR "
             
              If Cells(3, "B").Value > "ZZZ" Then
                   If Cells(3, "G") > 0.0001 Then Cells(2, "H").Value = Cells(3, "H").Value / Cells(3, "G").Value
              Else
                    If Cells(3, "G").Value > 0.00001 Then
                       Cells(2, "H").Value = Cells(3, "H").Value / Cells(3, "G").Value
                    Else
                        Cells(2, "H").Value = Cells(3, "H").Value / Cells(3, "F").Value
                    End If
              End If
              
              If Cells(3, "R").Value > 0.001 Then
                   Cells(2, "G").Value = Format(Cells(3, "G").Value / Cells(3, "R").Value, "###.#") & " /HR "
                 
                   Cells(2, "s").Value = Format(Cells(3, "t").Value / Cells(3, "R").Value, "###.##") & " RMB/HR "
                    If Cells(3, "G").Value > 0.001 Then
                       
                        
                           If Left(Cells(3, "D").Value, 1) = "J" Or Left(Cells(3, "D").Value, 1) = "S" Or Left(Cells(3, "D").Value, 2) = "ZD" Or Left(Cells(3, "Q").Value, 1) = "S" Then
                                Cells(3, "E").Value = Format((0.05 * 60 * 1.1) / (Cells(3, "G").Value / Cells(3, "r").Value), "##.####") & " RMB/EA"
                            Else
                                If Cells(3, "D").Value = "A" Then
                                     Cells(3, "E").Value = Format((0.05 * 60 * 1.05) / (Cells(3, "G").Value / Cells(3, "r").Value), "##.####") & " RMB/EA" ' RMB/ EA
                                  Else
                                        Cells(3, "E").Value = Format((0.05 * 60 * 1.65) / (Cells(3, "G").Value / Cells(3, "r").Value), "##.####") & " RMB/EA" ' RMB/ EA
                                  End If
                            End If
                          
                           If Cells(3, "B").Value < "ZZZ" And Cells(3, "G") < 0.001 Then
                               Cells(2, "E").Value = Cells(2, "t").Value / Cells(3, "F").Value
                            Else
                              Cells(2, "E").Value = Cells(2, "t").Value / Cells(3, "G").Value
                          End If
                     Else
                        
                     End If
                     Set a = Worksheets("其他数据").Columns("M").find(Cells(3, "C") & "-" & Cells(3, "D"))
                     Cells(2, "E") = ""
                     If Not a Is Nothing Then
                        Cells(2, "E") = Format(Worksheets("其他数据").Cells(a.row, "N"), "#####.#") & " EA/HR"
                     End If
                     
               Else  'Cells(3, "R").Value > 0.001 Then
             
                 
                 If Cells(3, "B").Value < "ZZZ" Then
                             If Cells(3, "G").Value > 0.001 Or Cells(3, "F") > 0.01 Then
                                  If Cells(3, "G") < 0.001 And Cells(3, "T") > 0.001 Then
                                        Cells(3, "E").Value = decimalstring(Cells(3, "t").Value / Cells(3, "F").Value, 2) & " /YCL:" & decimalstring(-1 * Cells(3, "U").Value / Cells(3, "F").Value, 2) & "占" & decimalstring(Int(-1 * Cells(3, "U").Value / (Cells(3, "T").Value) * 100), 1) & "%"
                                  Else
                                         Cells(3, "E").Value = Cells(3, "t").Value / Cells(3, "G").Value
                                  End If
                              End If
                           
                     End If
                     Cells(2, "G").Value = 0 & " /HR "
             End If
             
               
          Cells(2, "F").Value = Format(chanzhi(Cells(3, "C"), Cells(3, "D"), ""), "###.#") & " /HR "

End Sub
Sub MJ保存()
   MJ保存1 True
End Sub
Sub MJ保存1(Optional mjsheet = False)
   evt = Application.EnableEvents
    Application.EnableEvents = False
           Application.Calculation = xlCalculationManual
           Application.CalculateBeforeSave = False
           Application.DisplayAlerts = False
           
  
   Select Case ActiveWorkbook.Name
           
                 Case "订单连接"
                       Application.Run macro:="订单连接.xlsm!fadingdan1"
                 Case "外协订单"
                       fawaixiedingdan
                 Case "出入库单.xlsm"
                      If ActiveWorkbook.ReadOnly <> True And mjsheet = False Then
                         ActiveWorkbook.Save
                      Else
                         出入库单保存
                      End If
                 Case "mj.xlsm", "xmj.xlsm"
        
                     If ActiveWorkbook.ReadOnly <> True And mjsheet = False Then
                        ActiveWorkbook.Save
                        GoTo endsel
                      End If
                   Select Case ActiveSheet.Name
   
                       Case "出入统计total"
        
                            出入统计total保存
        
        
                       Case "订单目录"
                            If MsgBox(" 保存 订单目录 ? (要生产计划保存-转到 生产计划) ", vbOKCancel) = 1 Then 订单目录数据保存

                       Case "送货情况"
                            送货情况保存
                       Case "原材料情况"
                             原材料情况保存
                       Case "加工情况"
                             加工情况保存
                       Case "日抽检表"
                            日抽检数据保存
                       Case "外协订单"
                            外协订单数据保存
                       Case "欠单加工情况"
                             欠单加工情况保存
                       Case "生产计划数据", "生产计划"
                             生产计划数据保存
                       Case "原材料顺序"
                            原材料顺序list保存
                       Case "工序号码登记簿", "工序号码", "客户产品号码", "客户", "外协厂家", "原材料厂家"
                            工序号码登记簿保存
                       Case "客户价格表"
                             客户价格表保存
                       Case "原材料价格"
                             保存_ycl价格
                     Case Else
                       
                  End Select
endsel:
                 Case Else
                    If at("周计划表", ActiveWorkbook.Name) > 0 And ActiveWorkbook.ReadOnly = True And at("计划表", ActiveSheet.Name) > 0 Then
                       周计划表保存
                    Else
                      If at("周计划表", ActiveWorkbook.Name) > 0 Then
                         If at("mj.xlsm", wb.Name) > 0 Then
                           Application.Run macro:=wb.Name & "!计划日期DIsplay"
                          End If
                      End If
                     End If
                      If ActiveWorkbook.ReadOnly = False Then
                         ActiveWorkbook.Save
                         If ActiveWorkbook.Name = "机床性能巡检记录表.xlsm" Then
                           Dim fs As Object
                                
                              Set fs = CreateObject("Scripting.FileSystemObject")

                                fs.CopyFile ActiveWorkbook.path & "\" & ActiveWorkbook.Name, ActiveWorkbook.path & "\" & Replace(ActiveWorkbook.Name, ".xlsm", "") & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".xlsm"
                            
                                Set fs = Nothing

                         End If
                  End If
    End Select
      'Application.CalculateBeforeSave = True
      Application.EnableEvents = evt
End Sub


Sub select成本分析()
 
 wbcheck
   wb.Worksheets("成本分析").Visible = True
   wb.Worksheets("成本分析").Activate
End Sub
Sub select工序分析()

 

 wbcheck
 Workbooks("maching-flange-工序分析.xlsm").Sheets("工序分析").Visible = True
   Workbooks("maching-flange-工序分析.xlsm").Sheets("工序分析").Activate
End Sub
Sub select对账单()

 
 wbcheck
 wb.Worksheets("对账单").Visible = True
   wb.Worksheets("对账单").Activate
End Sub

Sub 周计划表保存()
 Dim rr, rr1 As Range
 Dim rowmax As Long
 Dim awb As Workbook
 Dim aws As Worksheet
'Dim wb As Workbook
 evt = Application.EnableEvents
 dvt = Application.DisplayAlerts
    Application.EnableEvents = False
     Application.DisplayAlerts = False
     wbcheck

    Set awb = ActiveWorkbook
     Set aws = ActiveSheet
    x1 = 0
    For crow = 5 To Range("A1000000").End(xlUp).row
       If Cells(crow, "A").Interior.color = 255 Then
          x1 = 1
          Exit For
       End If
    Next
    
      If x1 = 0 Then
          MsgBox ("计划  没有 修改的东西   ? (Yes/No) 。。。。。 ")
          ActiveSheet.AutoFilterMode = False
          GoTo Endp:
      End If
      
    If fileopened("Z:\美进文当\shengchanbu\生产部\周计划表.xlsx", False) <> 0 Then GoTo Endp:
    
     With awb.Worksheets("计划表")
     
         rowmax = .Range("A1000000").End(xlUp).row
        
     End With
    
    
     With Workbooks("周计划表.xlsx").Worksheets("计划表")
    
          Cells(1, "A") = khClosedFileValue("Z:\美进文当\shengchanbu\生产部\", "change.xlsx", "sheet1", Cells(2, 2).Address)
          Set rr = Range(.Cells(1, "A"), .Cells(10000, "Z"))
       
        If Format(Cells(1, "A"), "YY/MM/DD ") & Format(Cells(1, "A"), "HH:mm:ss") < Format(FileDateTime("Z:\美进文当\shengchanbu\生产部\周计划表.xlsx"), "YY/MM/DD ") & Format(FileDateTime("Z:\美进文当\shengchanbu\生产部\周计划表.xlsx"), "HH:mm:ss") Then
     
             rowmax = .Range("A1000000").End(xlUp).row + 1
              If rowmax < 4 Then rowmax = 5
          Else
    
              rr.Offset(4, 0).ClearContents
              rowmax = 5
         End If
      End With
      xrowmax = rowmax
       aws.Activate
       For crow = 5 To Range("A1000000").End(xlUp).row
           If Cells(crow, "A").Interior.color = 255 Then
             Rows(crow).Copy Destination:=Workbooks("周计划表.xlsx").Worksheets("计划表").Range("A" & rowmax)
             Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(rowmax, "J") = awb.path & "\" & awb.Name
             Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(rowmax, "Z") = Cells(crow, "Z")
             color Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(crow, "A"), 0
             rowmax = rowmax + 1
          End If
        Next
         If MsgBox("计划表 确定 保存  ? (Yes/No) 。。。。。 ", 1) = 1 Then
            Application.Calculation = xlCalculationManual
            Application.CalculateBeforeSave = False
            Workbooks("周计划表.xlsx").Save
            Workbooks("周计划表.xlsx").Close savechanges:=False
            Application.Calculation = xlCalculationManual
             
             MsgBox ("计划表 保存 成功。。。。。 ")
            

         Else
            Workbooks("周计划表.xlsx").Close savechanges:=False
         End If

Endp:
 Application.EnableEvents = evt
 Application.DisplayAlerts = False
End Sub
Sub read计划表()
  
Dim evt, dvt, alreadyopen As Boolean
Dim rr, rrjh, wbrrddml, a As Range
Dim col, crow, rowmax As Long
Dim filedatetimeold As String
Dim awb As Workbook
Dim aws As Worksheet
evt = Application.EnableEvents
dvt = Application.DisplayAlerts
On Error GoTo 0
Application.DisplayAlerts = False
 Application.EnableEvents = False
      alreadyopen = True
     Set aws = ActiveSheet
     Set awb = ActiveWorkbook
     
     If wbopencheck("周计划表.xlsx") = 0 Then
           
        If fileopened(wb.Worksheets("change").Cells(2, "A").Value, True) <> 0 Then GoTo Endp:
           alreadyopen = False

     End If
 
         Workbooks("周计划表.xlsx").Worksheets("计划表").Activate
         Set rrjh = Range(Cells(1, "A"), Cells(Range("A1000000").End(xlUp).row, "Z"))
         filedatetimeold = FileDateTime(wb.Worksheets("change").Cells(2, "A").Value)
         
      For crow = 5 To Workbooks("周计划表.xlsx").Worksheets("计划表").Range("A1000000").End(xlUp).row
          
          If fileopened(Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(crow, "J").Value, False) <> 0 Then
             Workbooks("周计划表.xlsx").Close savechanges:=False
            GoTo Endp:
          End If
             For i = 10 To Len(Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(crow, "J").Value)
                If Mid(Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(crow, "J").Value, i, 1) = "\" Then x1 = i
             Next
               
                
             Set awb = Workbooks(Mid(Workbooks("周计划表.xlsx").Worksheets("计划表").Cells(crow, "J"), x1 + 1, 20))
             aws.Activate
          Set a = Columns("Z").find(rrjh(crow, "Z"), LookIn:=xlValues, LookAt:=xlWhole)
         
          If a Is Nothing Then
              Rows(awb.Range("Z65535").End(xlUp).row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
             Set a = Cells(Range("Z65535").End(xlUp).row + 1, "A")
          End If
             Cells(a.row, "A") = rrjh(crow, "A")
             Cells(a.row, "B") = rrjh(crow, "B")
             Cells(a.row, "C") = rrjh(crow, "C")
              Cells(a.row, "D") = rrjh(crow, "D")
              Cells(a.row, "E") = rrjh(crow, "E")
              Cells(a.row, "F") = rrjh(crow, "F")
              Cells(a.row, "G") = rrjh(crow, "G")
              Cells(a.row, "H") = rrjh(crow, "H")
              Cells(a.row, "I") = rrjh(crow, "I")
              Cells(a.row, "J") = rrjh(crow, "J")
         If rrjh(crow, "J") <> rrjh(crow + 1, "J") Then
            Application.Calculation = xlCalculationManual
            Application.CalculateBeforeSave = False
            Workbooks(awb.Name).Save
            Workbooks(awb.Name).Close savechanges:=False
         End If
      Next
     
     If alreadyopen = False Then Workbooks("周计划表.xlsx").Close savechanges:=False
    
     
     wb.Worksheets("change").Cells(2, "B").Value = filedatetimeold
    
     Set rrjh = Nothing
     Set rrs = Nothing
     Set rrq = Nothing
Endp:
     Application.DisplayAlerts = False
     Application.EnableEvents = evt
     On Error GoTo 0
End Sub

Sub 原材料加工费修改()
    Application.EnableEvents = False
   For crow = 5 To Range("A65535").End(xlUp).row
      If at("TAC", Cells(crow, "I")) <> 0 Then GoTo endnext
      If Right(Cells(crow, "J"), 2) = "-K" Then
          If at("6061", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 8.35 / 1.13
          If at("6063", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 8.15 / 1.13
      Else
           If at("3003", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 7.02 / 1.13
           If at("6061", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 7.15 / 1.13
           If at("6005", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 6.95 / 1.13
           If at("G77", Cells(crow, "K")) > 0 Then Cells(crow, "O") = 7.5 / 1.13
      
      End If
endnext:
     Next
End Sub

Sub reg_stackarray(crow)
          
           If stackarray(10, 1) = ActiveSheet.Name And stackarray(1, 0) = ActiveWorkbook.Name Then
                  
           Else
              move_stackarray -1
           End If
           j = 10
               stackarray(j, 0) = ActiveWorkbook.Name
               stackarray(j, 1) = ActiveSheet.Name
               stackarray(j, 2) = crow
               stackarray(j, 3) = owsputh
               stackarray(j, 4) = puycl
               stackarray(j, 5) = pugx
                
    
End Sub
Sub move_stackarray(inc As Integer)
    If inc < 0 Then
         For j = 0 To 0
             If stackarray(10, 1) = ActiveSheet.Name And stackarray(10, 0) = ActiveWorkbook.Name Then Exit For
             If stackarray(10, 0) = "" Then Exit For
             
             For K = 0 To 9
                
                stackarray(K, 0) = stackarray(K + 1, 0)
                stackarray(K, 1) = stackarray(K + 1, 1)
                stackarray(K, 2) = stackarray(K + 1, 2)
                stackarray(K, 3) = stackarray(K + 1, 3)
                stackarray(K, 4) = stackarray(K + 1, 4)
                stackarray(K, 5) = stackarray(K + 1, 5)
             Next ' k
          Next
    Else
       
       For j = 0 To 10
             If stackarray(10, 1) = ActiveSheet.Name And stackarray(10, 0) = ActiveWorkbook.Name And stackarray(9, 1) <> ActiveSheet.Name And stackarray(9, 0) <> ActiveWorkbook.Name Then
                     Application.GoTo Cells(stackarray(10, 2), "A"), Scroll:=True
                     Rows(stackarray(10, 2)).Select
                   End If
                  ' If stackarray(9, 0) = "" Then
                    '   Application.Goto Cells(stackarray(10, 2), "A"), Scroll:=True
                      ' Rows(stackarray(10, 2)).Select
                   'End If
             If Not (stackarray(10, 1) = ActiveSheet.Name And stackarray(10, 0) = ActiveWorkbook.Name) Then Exit For
             If stackarray(9, 0) = "" Then Exit For
             
              For K = 10 To 1 Step -1
                  
                   stackarray(K, 0) = stackarray(K - 1, 0)
                   stackarray(K, 1) = stackarray(K - 1, 1)
                   stackarray(K, 2) = stackarray(K - 1, 2)
                   stackarray(K, 3) = stackarray(K - 1, 3)
                   stackarray(K, 4) = stackarray(K - 1, 4)
                    stackarray(K, 5) = stackarray(K - 1, 5)
              Next ' k
      Next

    End If
End Sub
Sub del_stackarray(crow)
    For i = 10 To 0 Step -1
        If stackarray(i, 0) <> "" Then
           If stackarray(i, 1) = ActiveSheet.Name And stackarray(i, 0) = ActiveWorkbook.Name Then
               stackarray(i, 0) = ""
               stackarray(j, 1) = ""
           End If
           
               Exit For
         End If
      Next
      
End Sub
