VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HolidayJP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public yearVal As Integer
Private MyDic As Dictionary

Private Const SUN = 1
Private Const MON = 2
Private Const TUE = 3
Private Const WED = 4
Private Const THU = 5
Private Const FRU = 6
Private Const SAT = 7
Private Const WEEK1 = 1
Private Const WEEK2 = 8
Private Const WEEK3 = 15
Private Const WEEK4 = 22
Private Const FURIKAE_START = #4/12/1974#


Private Sub Class_Initialize()

    yearVal = 0
    Set MyDic = Nothing

End Sub

Sub SetYear(year As Integer)

    yearVal = year
    Set MyDic = GetHolidayDic(year)

End Sub

Public Function GetDate(year, holiday As String)
    Dim theDay As Date
    
    Select Case holiday
        Case "元日"
            theDay = DateSerial(year, 1, 1)
        Case "成人の日"
            If year >= 2000 Then
                '1月の第2月曜日
                theDay = GetWeekNumYobi(year, 1, WEEK2, MON)
            ElseIf year >= 1949 Then
                theDay = DateSerial(year, 1, 15)
            End If
        Case "建国記念の日"
            If year >= 1967 Then
                theDay = DateSerial(year, 2, 11)
            End If
        Case "昭和の日"
            If year >= 2007 Then
                theDay = DateSerial(year, 4, 29)
            End If
        Case "憲法記念日"
            If year >= 1949 Then
                theDay = DateSerial(year, 5, 3)
            End If
        Case "みどりの日"
            If year >= 2007 Then
                theDay = DateSerial(year, 5, 4)
            ElseIf year >= 1989 Then
                theDay = DateSerial(year, 4, 29)
            End If
        Case "こどもの日"
            If year >= 1949 Then
                theDay = DateSerial(year, 5, 5)
            End If
        Case "海の日"
            If year = 2021 Then
                theDay = DateSerial(year, 7, 22)
            ElseIf year = 2020 Then
                theDay = DateSerial(year, 7, 23)
            ElseIf year >= 2003 Then
                '7月の第3月曜日
                theDay = GetWeekNumYobi(year, 7, WEEK3, MON)
            ElseIf year >= 1996 Then
                theDay = DateSerial(year, 7, 20)
            End If
        Case "山の日"
            If year = 2021 Then
                theDay = DateSerial(year, 8, 8)
            ElseIf year = 2020 Then
                theDay = DateSerial(year, 8, 10)
            ElseIf year >= 2016 Then
                theDay = DateSerial(year, 8, 11)
            End If
        Case "敬老の日"
            If year >= 2003 Then
                '9月の第3月曜日
                theDay = GetWeekNumYobi(year, 9, WEEK3, MON)
            ElseIf year >= 1966 Then
                theDay = DateSerial(year, 9, 15)
            End If
        Case "体育の日"
            If year >= 2000 And year <= 2019 Then
                '10月の第2月曜日
                theDay = GetWeekNumYobi(year, 10, WEEK2, MON)
            ElseIf year >= 1966 And year <= 1999 Then
                theDay = DateSerial(year, 10, 10)
            End If
        Case "スポーツの日"
            If year >= 2022 Then
                '10月の第2月曜日
                theDay = GetWeekNumYobi(year, 10, WEEK2, MON)
            ElseIf year = 2021 Then
                theDay = DateSerial(year, 7, 23)
            ElseIf year = 2020 Then
                theDay = DateSerial(year, 7, 24)
            End If
        Case "文化の日"
            If year >= 1948 Then
                theDay = DateSerial(year, 11, 3)
            End If
        Case "勤労感謝の日"
            If year >= 1948 Then
                theDay = DateSerial(year, 11, 23)
            End If
        Case "天皇誕生日"
            If year >= 2020 Then
                theDay = DateSerial(year, 2, 23)
            ElseIf year >= 1989 And year <= 2018 Then
                theDay = DateSerial(year, 12, 23)
            ElseIf year >= 1949 And year <= 1988 Then
                theDay = DateSerial(year, 4, 29)
            End If
        Case "春分の日"
            If year >= 1980 And year <= 2099 Then
                theDay = DateSerial(year, 3, Int(20.8431 + 0.242194 * (year - 1980)) - Int((year - 1980) / 4))
            ElseIf year >= 1949 And year <= 1979 Then
                theDay = DateSerial(year, 3, Int(20.8357 + 0.242194 * (year - 1980)) - Int((year - 1983) / 4))
            ElseIf year >= 2100 And year <= 2150 Then
                theDay = DateSerial(year, 3, Int(21.851 + 0.242194 * (year - 1980)) - Int((year - 1980) / 4))
            End If
        Case "秋分の日"
            If year >= 1980 And year <= 2099 Then
                theDay = DateSerial(year, 9, Int(23.2488 + 0.242194 * (year - 1980)) - Int((year - 1980) / 4))
            ElseIf year >= 1948 And year <= 1979 Then
                theDay = DateSerial(year, 9, Int(23.2588 + 0.242194 * (year - 1980)) - Int((year - 1983) / 4))
            ElseIf year >= 2100 And year <= 2150 Then
                theDay = DateSerial(year, 9, Int(24.2488 + 0.242194 * (year - 1980)) - Int((year - 1980) / 4))
            End If
        Case "即位礼正殿の儀"
            If year = 2019 Then
                theDay = DateSerial(year, 10, 22)
            ElseIf year = 1990 Then
                theDay = DateSerial(year, 11, 12)
            End If
        Case "天皇の即位の日"
            If year = 2019 Then
                theDay = DateSerial(year, 5, 1)
            End If
        Case "皇太子徳仁親王の結婚の儀"
            If year = 1993 Then
                theDay = DateSerial(year, 6, 9)
            End If
        Case "昭和天皇の大喪の礼"
            If year = 1989 Then
                theDay = DateSerial(year, 2, 24)
            End If
        Case "皇太子明仁親王の結婚の儀"
            If year = 1959 Then
                theDay = DateSerial(year, 4, 10)
            End If
    End Select
    
    GetDate = theDay
    
End Function

Function GetWeekNumYobi(year, month, WeekNum, wday)
    Dim wday0, day
    
    wday0 = Weekday(DateSerial(year, month, WeekNum))
    day = WeekNum + (wday - wday0 + 7) Mod 7
    
    GetWeekNumYobi = DateSerial(year, month, day)

End Function

Property Get Holidays()
    Holidays = Array("元日", "成人の日", "建国記念の日", "昭和の日", "憲法記念日", "みどりの日", "こどもの日", "海の日", "山の日", "敬老の日", "体育の日", "スポーツの日", "文化の日", "勤労感謝の日", "天皇誕生日", "春分の日", "秋分の日", "即位礼正殿の儀", "天皇の即位の日", "皇太子徳仁親王の結婚の儀", "昭和天皇の大喪の礼", "皇太子明仁親王の結婚の儀")
End Property

Function GetHolidayDic(Optional year As Integer) As Dictionary
    Dim MyDic As Dictionary
    Dim i
    Dim holiday As Date
    Dim HolidayName As String
    
    If year = 0 Then year = yearVal
    Set MyDic = New Dictionary
    
    For i = 0 To UBound(Me.Holidays) - 1
    
        HolidayName = CStr(Me.Holidays(i))
        holiday = GetDate(year, HolidayName)
        
        If holiday > 0 Then 'Date型変数の初期値は0
            MyDic.Add holiday, HolidayName
        End If
    
    Next
    
    If year >= 2007 Then
        For i = 0 To MyDic.Count - 1
            holiday = MyDic.Keys(i)
            If Weekday(holiday) = SUN Then
                Do While MyDic.Exists(holiday)
                    holiday = holiday + 1
                Loop
                MyDic.Add holiday, "振替休日"
            End If
        Next
    ElseIf year >= 1973 Then
        For i = 0 To MyDic.Count - 1
            holiday = MyDic.Keys(i)
            If Weekday(holiday) = SUN And holiday >= FURIKAE_START Then
                    MyDic.Add holiday + 1, "振替休日"
            End If
        Next
    End If
    
    Set MyDic = sortDic(MyDic)
        
    If year >= 1986 Then
        For i = 0 To MyDic.Count - 1
            holiday = MyDic.Keys(i)
            If MyDic.Exists(holiday + 2) And Not MyDic.Exists(holiday + 1) And Weekday(holiday) <> SAT Then
                MyDic.Add holiday + 1, "国民の休日"
            End If
        Next
    End If
    
    Set GetHolidayDic = sortDic(MyDic)
    
End Function

Function GetHolidayArray(year As Integer)
    Dim dic As Dictionary
    
    If MyDic Is Nothing Then
        Set dic = GetHolidayDic(year)
    Else
        se dic = MyDic
    End If
    
    GetHolidayArray = dic.Keys
End Function

Function isHoliday(dateVal) As Boolean
    Dim dic As Dictionary
    
    If Len(CStr(dateVal)) <= 5 Then
        If yearVal = 0 Then
            dateVal = CDate(dateVal)
            yearVal = year(dateVal)
        Else
            dateVal = DateSerial(yearVal, month(CDate(dateVal)), day(CDate(dateVal)))
        End If
    ElseIf Format(CDate(dateVal), "yyyy") = CStr(yearVal) Then
        dateVal = CDate(dateVal)
    Else
        MsgBox "年が一致しません。", vbCritical
        isHoliday = False
        Exit Function
    End If
    
    If MyDic Is Nothing Then
        Set dic = GetHolidayDic(yearVal)
    Else
        Set dic = MyDic
    End If

    isHoliday = dic.Exists(dateVal)
    
End Function

Function HolidayName(dateVal)
    Dim dic As Dictionary
    
    If Len(CStr(dateVal)) <= 5 Then
        If yearVal = 0 Then
            dateVal = CDate(dateVal)
            yearVal = year(dateVal)
        Else
            dateVal = DateSerial(yearVal, month(CDate(dateVal)), day(CDate(dateVal)))
        End If
    ElseIf Format(CDate(dateVal), "yyyy") = CStr(yearVal) Then
        dateVal = CDate(dateVal)
    Else
        HolidayName = "年が一致しません。"
        Exit Function
    End If
    
    If MyDic Is Nothing Then
        Set dic = GetHolidayDic(yearVal)
    Else
        Set dic = MyDic
    End If
    
    If dic.Exists(dateVal) Then
        HolidayName = dic.Item(dateVal)
    Else
        HolidayName = ""
    End If

End Function

Sub dispHolidays(Optional year As Integer)
    
    Dim dic As Dictionary
    Dim days As Variant
    Dim yobi As Variant
    Dim str As String
    Dim i As Integer

    If year = 0 Then
        If yearVal > 0 Then
            year = yearVal
        Else
            year = Int(Format(Date, "yyyy"))
        End If
    End If
    
    If MyDic Is Nothing Then
        Set dic = GetHolidayDic(year)
    Else
        Set dic = MyDic
    End If

    days = dic.Keys

    yobi = Array("", "日", "月", "火", "水", "木", "金", "土")
    
    str = ""
    For i = 0 To UBound(days)
        str = str & Format(days(i), "yyyy/MM/dd") & " (" & yobi(Weekday(days(i))) & ") " & dic.Items(i) & vbCrLf
    Next
    MsgBox Trim(str), vbOKOnly + vbInformation, CStr(year) & "年 祝日一覧"

End Sub

Function sortDic(dic As Dictionary) As Dictionary
    
    Dim daysArr As Variant
    Dim holiday As Date, i
    Dim newDic As Dictionary
    
    daysArr = dic.Keys
    Call QuickSort(daysArr, LBound(daysArr), UBound(daysArr))
    
    Set newDic = New Dictionary
    For i = 0 To UBound(daysArr)
        holiday = daysArr(i)
        newDic.Add holiday, dic.Item(holiday)
    Next
    
    Set sortDic = newDic

End Function

Sub QuickSort(ByRef arr As Variant, low As Long, high As Long)

    Dim i As Long, j As Long
    Dim pivot As Long
    Dim temp
    
    If low < high Then
        For i = low To high - 1
            If arr(i) > arr(high) Then
                For j = high - 1 To low Step -1
                    If i = j Then
                        temp = arr(i)
                        arr(i) = arr(high)
                        arr(high) = temp
                        pivot = i
                        GoTo NextPart
                    ElseIf arr(j) < arr(high) Then
                        temp = arr(i)
                        arr(i) = arr(j)
                        arr(j) = temp
                        Exit For
                    End If
                Next
            ElseIf i = high - 1 Then
                pivot = high
            End If
        Next
NextPart:
        Call QuickSort(arr, low, pivot - 1)
        Call QuickSort(arr, pivot + 1, high)
    End If

End Sub
