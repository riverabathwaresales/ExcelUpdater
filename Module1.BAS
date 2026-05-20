Attribute VB_Name = "Module1"
Option Explicit

Private globalData As Variant
Private dataLoaded As Boolean

Public isRebuildingSummary As Boolean

Public imgWB As Workbook
Public vbaFeaturesEnabled As Boolean

Public Const APP_NAME As String = "MyApp"

Private Const USERS_URL As String = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQuUH611zbdHjPZlnfi_Efs8LZJIxukZLuaFjUAm3iHpfy1q1pZNzS9HmQpBSDz9DNxK_vBYOAq3bxu/pub?gid=0&single=true&output=csv"

Private Const DATA_URL As String = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQuUH611zbdHjPZlnfi_Efs8LZJIxukZLuaFjUAm3iHpfy1q1pZNzS9HmQpBSDz9DNxK_vBYOAq3bxu/pub?gid=1481842660&single=true&output=csv"

Private Const IMAGE_URL As String = "https://docs.google.com/spreadsheets/d/1UPMKkbkClj_FZitYOx6XyxVzevAxLHZv/export?format=xlsx"

Private Const API_URL As String = "https://script.google.com/macros/s/AKfycbzeVREjtiBOYKYIrKvLVvlScaWV3Qhi_DGc6PVL3L-KF6pR4SDavSKXfBl4pFYTqh4y/exec"

Private Const SUMMARY_KEY_PREFIX As String = "__BATH__|"
Private Const SUMMARY_TOTAL_KEY As String = "__TOTAL__"

Private Function CacheFolder() As String
    CacheFolder = Environ$("LOCALAPPDATA") & "\MyAppCache\"
End Function

Public Function HasCacheFolder() As Boolean
    HasCacheFolder = (Dir(CacheFolder(), vbDirectory) <> "")
End Function

Private Sub EnsureCacheFolder()
    If Dir(CacheFolder(), vbDirectory) = "" Then
        MkDir CacheFolder()
    End If
End Sub

Private Function TodayStamp() As String
    TodayStamp = Format(Date, "yyyymmdd")
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir(filePath)) > 0)
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    If Not FileExists(filePath) Then Exit Function

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText
        .Close
    End With
End Function

Private Sub SaveTextFile(ByVal filePath As String, ByVal txt As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText txt
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Function GetLatestCachedFile(ByVal filePattern As String) As String
    Dim f As String
    Dim fullPath As String
    Dim newestPath As String
    Dim newestDate As Date

    f = Dir(CacheFolder() & filePattern)

    Do While f <> ""
        fullPath = CacheFolder() & f

        If newestPath = "" Then
            newestPath = fullPath
            newestDate = FileDateTime(fullPath)
        ElseIf FileDateTime(fullPath) > newestDate Then
            newestPath = fullPath
            newestDate = FileDateTime(fullPath)
        End If

        f = Dir()
    Loop

    GetLatestCachedFile = newestPath
End Function

Private Function NormalizeTextLines(ByVal txt As String) As String
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    NormalizeTextLines = txt
End Function

Private Function CleanCsvValue(ByVal txt As String) As String
    Dim s As String

    s = txt
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim$(s)

    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If

    CleanCsvValue = Trim$(s)
End Function

Private Function AddNoCacheQuery(ByVal fileUrl As String) As String
    If InStr(1, fileUrl, "?", vbTextCompare) > 0 Then
        AddNoCacheQuery = fileUrl & "&_ts=" & Format$(Now, "yyyymmddhhnnss")
    Else
        AddNoCacheQuery = fileUrl & "?_ts=" & Format$(Now, "yyyymmddhhnnss")
    End If
End Function

Private Function DownloadText(ByVal fileUrl As String) As String
    On Error GoTo failed

    Dim http As Object
    Dim attempt As Long
    Dim requestUrl As String

    For attempt = 1 To 3
        requestUrl = AddNoCacheQuery(fileUrl)

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", requestUrl, False
        http.SetTimeouts 10000, 10000, 20000, 20000
        http.SetRequestHeader "Cache-Control", "no-cache"
        http.SetRequestHeader "Pragma", "no-cache"
        http.SetRequestHeader "User-Agent", "Mozilla/5.0"
        http.Option(6) = True
        http.Send

        If http.Status = 200 Then
            DownloadText = http.ResponseText
            Exit Function
        End If

        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt

failed:
    DownloadText = ""
End Function

Private Function DownloadBinaryFile(ByVal fileUrl As String, ByVal savePath As String) As Boolean
    On Error GoTo failed

    Dim http As Object
    Dim stm As Object
    Dim attempt As Long
    Dim requestUrl As String

    For attempt = 1 To 3
        requestUrl = AddNoCacheQuery(fileUrl)

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", requestUrl, False
        http.SetTimeouts 10000, 10000, 20000, 20000
        http.SetRequestHeader "Cache-Control", "no-cache"
        http.SetRequestHeader "Pragma", "no-cache"
        http.SetRequestHeader "User-Agent", "Mozilla/5.0"
        http.Option(6) = True
        http.Send

        If http.Status = 200 Then
            Set stm = CreateObject("ADODB.Stream")
            With stm
                .Type = 1
                .Open
                .Write http.ResponseBody
                .SaveToFile savePath, 2
                .Close
            End With

            DownloadBinaryFile = True
            Exit Function
        End If

        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt

failed:
    DownloadBinaryFile = False
End Function

Private Function GetDailyCachedText(ByVal fileUrl As String, ByVal cachePrefix As String) As String
    EnsureCacheFolder

    Dim todayFile As String
    Dim latestFile As String
    Dim txt As String

    todayFile = CacheFolder() & cachePrefix & "_" & TodayStamp() & ".csv"

    If FileExists(todayFile) Then
        GetDailyCachedText = ReadTextFile(todayFile)
        Exit Function
    End If

    txt = DownloadText(fileUrl)

    If txt <> "" Then
        SaveTextFile todayFile, txt
        GetDailyCachedText = txt
        Exit Function
    End If

    latestFile = GetLatestCachedFile(cachePrefix & "_*.csv")
    If latestFile <> "" Then
        GetDailyCachedText = ReadTextFile(latestFile)
    End If
End Function

Public Function GetUsersCSV() As String
    GetUsersCSV = DownloadText(USERS_URL)
End Function

Public Function GetDataCSV() As String
    GetDataCSV = GetDailyCachedText(DATA_URL, "data")
End Function

Private Function SessionFilePath() As String
    SessionFilePath = CacheFolder() & "session_user.txt"
End Function

Public Function HasSessionCache() As Boolean
    If Not HasCacheFolder() Then
        HasSessionCache = False
    Else
        HasSessionCache = FileExists(SessionFilePath())
    End If
End Function

Private Sub SaveSessionUser(ByVal userName As String)
    EnsureCacheFolder
    SaveTextFile SessionFilePath(), Trim$(userName)
End Sub

Private Function SafeToken(ByVal txt As String) As String
    Dim s As String

    s = UCase$(Trim$(txt))
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    s = Replace(s, " ", "_")

    SafeToken = s
End Function

Private Function ValidationCacheFile(ByVal userName As String) As String
    ValidationCacheFile = CacheFolder() & "validate_" & SafeToken(userName) & "_" & TodayStamp() & ".txt"
End Function

Private Sub SaveValidationCache(ByVal userName As String, ByVal statusText As String)
    EnsureCacheFolder
    SaveTextFile ValidationCacheFile(userName), Trim$(statusText)
End Sub

Private Function GetTodayValidationCache(ByVal userName As String) As String
    Dim fp As String

    fp = ValidationCacheFile(userName)

    If FileExists(fp) Then
        GetTodayValidationCache = Trim$(ReadTextFile(fp))
    End If
End Function

Private Sub DeleteUserValidationCache(ByVal userName As String)
    On Error Resume Next

    Dim f As String
    Dim patt As String

    If Not HasCacheFolder() Then Exit Sub
    If Trim$(userName) = "" Then Exit Sub

    patt = "validate_" & SafeToken(userName) & "_*.txt"
    f = Dir(CacheFolder() & patt)

    Do While f <> ""
        Kill CacheFolder() & f
        f = Dir()
    Loop
End Sub

Public Sub ClearSavedLogin(Optional ByVal userName As String = "")
    On Error Resume Next
    SaveSetting APP_NAME, "Login", "User", vbNullString
    If FileExists(SessionFilePath()) Then Kill SessionFilePath()
    If Trim$(userName) <> "" Then DeleteUserValidationCache userName
End Sub

Public Function GetPCID() As String
    Dim pcid As String
    Dim machineName As String
    Dim userName As String

    pcid = Trim$(GetSetting(APP_NAME, "System", "PCID", vbNullString))

    If pcid = vbNullString Then
        machineName = Trim$(Environ$("COMPUTERNAME"))
        userName = Trim$(Environ$("USERNAME"))

        If machineName = vbNullString Then machineName = "PC"
        If userName = vbNullString Then userName = "USER"

        pcid = machineName & "_" & userName
        SaveSetting APP_NAME, "System", "PCID", pcid
    End If

    GetPCID = pcid
End Function

Private Function UrlEncode(ByVal txt As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim outTxt As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        code = Asc(ch)

        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122
                outTxt = outTxt & ch
            Case 45, 46, 95, 126
                outTxt = outTxt & ch
            Case 32
                outTxt = outTxt & "%20"
            Case Else
                outTxt = outTxt & "%" & Right$("0" & Hex(code), 2)
        End Select
    Next i

    UrlEncode = outTxt
End Function

Private Function CallServer(ByVal actionName As String, ByVal userName As String, ByVal passValue As String, ByVal pcid As String) As String
    On Error GoTo failed

    Dim url As String

    url = API_URL & _
          "?action=" & UrlEncode(actionName) & _
          "&user=" & UrlEncode(userName) & _
          "&pass=" & UrlEncode(passValue) & _
          "&pc=" & UrlEncode(pcid)

    CallServer = Trim$(DownloadText(url))
    Exit Function

failed:
    CallServer = ""
End Function

Public Function ServerLogin(ByVal userName As String, ByVal passValue As String, ByVal pcid As String) As String
    ServerLogin = CallServer("login", userName, passValue, pcid)
End Function

Public Function ServerValidate(ByVal userName As String, ByVal pcid As String) As String
    ServerValidate = CallServer("validate", userName, vbNullString, pcid)
End Function

Public Function ServerLogout(ByVal userName As String, ByVal pcid As String) As String
    ServerLogout = CallServer("logout", userName, vbNullString, pcid)
End Function

Private Function SafeStatus(ByVal txt As String) As String
    SafeStatus = UCase$(Trim$(txt))
End Function

Private Function IsForceLogout(ByVal txt As String) As Boolean
    IsForceLogout = (Left$(SafeStatus(txt), 12) = "FORCE_LOGOUT")
End Function

Private Function IsInactiveResult(ByVal txt As String) As Boolean
    txt = SafeStatus(txt)
    IsInactiveResult = (txt = "LOGIN_DENIED|INACTIVE" Or InStr(1, txt, "INACTIVE", vbTextCompare) > 0)
End Function

Private Function IsValidResult(ByVal txt As String) As Boolean
    txt = SafeStatus(txt)
    IsValidResult = (txt = "VALID|ACTIVE" Or txt = "LOGIN_SUCCESS|ACTIVE")
End Function

Public Sub DisableVBAFeatures()
    vbaFeaturesEnabled = False
End Sub

Public Sub EnableVBAFeatures()
    vbaFeaturesEnabled = True
End Sub

Private Function FindUserInUsersCsv(ByVal userName As String, ByRef foundPassword As String, ByRef foundStatus As String) As Boolean
    Dim data As String
    Dim arr As Variant
    Dim i As Long
    Dim c As Variant
    Dim rowUser As String

    data = GetUsersCSV()
    If data = vbNullString Then Exit Function

    data = NormalizeTextLines(data)
    arr = Split(data, vbLf)

    For i = 1 To UBound(arr)
        If Trim$(arr(i)) <> vbNullString Then
            c = Split(arr(i), ",")

            If UBound(c) >= 2 Then
                rowUser = CleanCsvValue(CStr(c(0)))

                If UCase$(rowUser) = UCase$(Trim$(userName)) Then
                    foundPassword = CleanCsvValue(CStr(c(1)))
                    foundStatus = UCase$(CleanCsvValue(CStr(c(2))))
                    FindUserInUsersCsv = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function ValidateLoginFromFreshCSV(ByVal userName As String, ByVal passValue As String) As Long
    Dim sheetPassword As String
    Dim sheetStatus As String

    If Not FindUserInUsersCsv(userName, sheetPassword, sheetStatus) Then
        ValidateLoginFromFreshCSV = -2
        Exit Function
    End If

    If StrComp(sheetPassword, Trim$(passValue), vbBinaryCompare) <> 0 Then
        ValidateLoginFromFreshCSV = 0
        Exit Function
    End If

    If sheetStatus <> "ACTIVE" Then
        ValidateLoginFromFreshCSV = -1
        Exit Function
    End If

    ValidateLoginFromFreshCSV = 1
End Function

Private Sub CompleteLoginSuccess(ByVal userName As String, ByVal validationStatus As String)
    SaveSetting APP_NAME, "Login", "User", Trim$(userName)
    SaveSessionUser Trim$(userName)
    SaveValidationCache userName, validationStatus
    EnableVBAFeatures
End Sub

Public Function PerformLogin() As Boolean
    Dim u As String
    Dim p As String
    Dim response As String
    Dim pcid As String
    Dim localCheck As Long

    u = InputBox("Enter User ID")
    If Trim$(u) = vbNullString Then Exit Function

    p = InputBox("Enter Password")
    If Trim$(p) = vbNullString Then Exit Function

    pcid = GetPCID()
    response = SafeStatus(ServerLogin(u, p, pcid))

    Select Case response
        Case "LOGIN_SUCCESS|ACTIVE"
            CompleteLoginSuccess u, "VALID|ACTIVE"
            MsgBox "Login Success!", vbInformation
            PerformLogin = True
            Exit Function

        Case "LOGIN_DENIED|INACTIVE"
            ClearSavedLogin u
            DisableVBAFeatures
            MsgBox "User INACTIVE!", vbCritical
            Exit Function

        Case "LOGIN_DENIED|INVALID"
            MsgBox "Invalid Login!", vbCritical
            Exit Function
    End Select

    localCheck = ValidateLoginFromFreshCSV(u, p)

    Select Case localCheck
        Case 1
            CompleteLoginSuccess u, "VALID|OFFLINE"
            MsgBox "Offline login using fresh users data!", vbInformation
            PerformLogin = True

        Case -1
            ClearSavedLogin u
            DisableVBAFeatures
            MsgBox "User INACTIVE!", vbCritical

        Case 0
            MsgBox "Invalid Login!", vbCritical

        Case Else
            MsgBox "Internet/cache Error!", vbCritical
    End Select
End Function

Public Function ApplyDailyUserValidation(ByVal userName As String) As Boolean
    On Error GoTo fallbackCheck

    Dim statusText As String
    Dim freshStatus As String
    Dim u As String

    u = Trim$(userName)
    If u = vbNullString Then GoTo denyUser

    statusText = SafeStatus(ServerValidate(u, GetPCID()))

    If statusText <> vbNullString Then
        SaveValidationCache u, statusText

        If IsForceLogout(statusText) Then
            ClearSavedLogin u
            DisableVBAFeatures
            MsgBox "This user was logged in from another PC. This PC is now logged out.", vbExclamation
            ApplyDailyUserValidation = False
            Exit Function
        End If

        If IsInactiveResult(statusText) Then
            ClearSavedLogin u
            DisableVBAFeatures
            MsgBox "User INACTIVE. VBA automation is disabled.", vbCritical
            ApplyDailyUserValidation = False
            Exit Function
        End If

        If IsValidResult(statusText) Then
            EnableVBAFeatures
            ApplyDailyUserValidation = True
            Exit Function
        End If
    End If

    If FindUserStatusFresh(u, freshStatus) Then
        If freshStatus <> "ACTIVE" Then
            GoTo denyUser
        End If
    ElseIf statusText = vbNullString Then
        GoTo fallbackCheck
    End If

    If statusText <> vbNullString Then
        EnableVBAFeatures
        ApplyDailyUserValidation = True
        Exit Function
    End If

fallbackCheck:
    statusText = SafeStatus(GetTodayValidationCache(userName))

    If statusText = vbNullString Then GoTo denyUser
    If IsForceLogout(statusText) Then GoTo denyUser
    If IsInactiveResult(statusText) Then GoTo denyUser

    EnableVBAFeatures
    ApplyDailyUserValidation = True
    Exit Function

denyUser:
    ClearSavedLogin u
    DisableVBAFeatures
    ApplyDailyUserValidation = False
End Function

Public Sub LoadDataOnce()
    If dataLoaded Then Exit Sub

    Dim raw As String
    Dim arr As Variant
    Dim i As Long

    raw = GetDataCSV()
    If raw = vbNullString Then Exit Sub

    raw = NormalizeTextLines(raw)
    arr = Split(raw, vbLf)
    ReDim globalData(1 To UBound(arr))

    For i = 1 To UBound(arr)
        If Trim$(arr(i)) <> vbNullString Then
            globalData(i) = Split(arr(i), ",")
        End If
    Next i

    dataLoaded = True
End Sub

Private Function GetDailyImageWorkbookPath() As String
    EnsureCacheFolder

    Dim todayFile As String
    Dim latestFile As String

    todayFile = CacheFolder() & "images_" & TodayStamp() & ".xlsx"

    If FileExists(todayFile) Then
        GetDailyImageWorkbookPath = todayFile
        Exit Function
    End If

    If DownloadBinaryFile(IMAGE_URL, todayFile) Then
        GetDailyImageWorkbookPath = todayFile
        Exit Function
    End If

    latestFile = GetLatestCachedFile("images_*.xlsx")
    GetDailyImageWorkbookPath = latestFile
End Function

Private Function GetOpenWorkbookByFullName(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function GetHiddenImageWorkbook() As Workbook
    Dim filePath As String
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim alreadyOpen As Workbook

    If Not imgWB Is Nothing Then
        On Error Resume Next
        Dim testName As String
        testName = imgWB.Name
        If Err.Number = 0 Then
            imgWB.Saved = True
            Set GetHiddenImageWorkbook = imgWB
            Exit Function
        End If
        Err.Clear
        Set imgWB = Nothing
        On Error GoTo 0
    End If

    filePath = GetDailyImageWorkbookPath()
    If filePath = vbNullString Then Exit Function

    Set alreadyOpen = GetOpenWorkbookByFullName(filePath)

    If Not alreadyOpen Is Nothing Then
        Set imgWB = alreadyOpen
        On Error Resume Next
        imgWB.Windows(1).Visible = False
        imgWB.Saved = True
        On Error GoTo 0
        Set GetHiddenImageWorkbook = imgWB
        Exit Function
    End If

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo failed

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set imgWB = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    imgWB.Windows(1).Visible = False
    imgWB.Saved = True

    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen

    Set GetHiddenImageWorkbook = imgWB
    Exit Function

failed:
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Set imgWB = Nothing
End Function

Private Function IsSummaryFormulaCell(ByVal c As Range) As Boolean
    On Error GoTo failed

    If c Is Nothing Then Exit Function
    If Not c.HasFormula Then Exit Function

    Dim f As String
    f = UCase$(c.Formula)

    If InStr(1, f, "SUM(", vbTextCompare) > 0 Or InStr(1, f, "SUBTOTAL(", vbTextCompare) > 0 Then
        IsSummaryFormulaCell = True
    End If

    Exit Function

failed:
    IsSummaryFormulaCell = False
End Function

Public Function IsProtectedTotalRow(ByVal sh As Worksheet, ByVal r As Long) As Boolean
    On Error GoTo safeExit

    If Not IsBathSheet(sh) Then Exit Function

    If UCase$(Trim$(CStr(sh.Cells(r, 6).Value))) = "TOTAL" Then
        IsProtectedTotalRow = True
        Exit Function
    End If

    If IsSummaryFormulaCell(sh.Cells(r, 12)) Or IsSummaryFormulaCell(sh.Cells(r, 13)) Then
        IsProtectedTotalRow = True
        Exit Function
    End If

safeExit:
End Function

Public Sub ClearRow(ByVal cell As Range)
    Dim r As Long
    Dim sh As Worksheet
    Dim shp As Shape

    r = cell.Row
    Set sh = cell.Worksheet

    If IsProtectedTotalRow(sh, r) Then Exit Sub

    sh.Cells(r, 3).ClearContents
    sh.Cells(r, 5).ClearContents
    sh.Cells(r, 11).ClearContents
    sh.Cells(r, 14).ClearContents
    sh.Cells(r, 12).ClearContents
    sh.Cells(r, 13).ClearContents

    For Each shp In sh.Shapes
        If shp.TopLeftCell.Row = r And shp.TopLeftCell.Column = 2 Then
            shp.Delete
        End If
    Next shp
End Sub

Public Sub RecalculateRow(ByVal sh As Worksheet, ByVal r As Long)
    If vbaFeaturesEnabled = False Then Exit Sub
    If IsProtectedTotalRow(sh, r) Then Exit Sub

    If Not sh.Cells(r, 12).HasFormula Then
        sh.Cells(r, 12).Formula = "=D" & r & "*E" & r
    End If

    If Not sh.Cells(r, 13).HasFormula Then
        sh.Cells(r, 13).Formula = "=D" & r & "*N" & r
    End If
End Sub

Public Sub FillDataAndImage(ByVal cell As Range)
    If vbaFeaturesEnabled = False Then Exit Sub

    Dim code As String
    Dim sh As Worksheet
    Dim r As Long
    Dim i As Long
    Dim wbImg As Workbook
    Dim ws As Worksheet
    Dim found As Range
    Dim shp As Shape
    Dim c As Range

    code = UCase$(Trim$(cell.Value))
    Set sh = cell.Worksheet
    r = cell.Row

    If IsProtectedTotalRow(sh, r) Then Exit Sub
    If code = "TOTAL" Or code = "CODE" Then Exit Sub

    ClearRow cell
    LoadDataOnce

    If Not dataLoaded Then Exit Sub

    For i = LBound(globalData) To UBound(globalData)
        If Not IsEmpty(globalData(i)) Then
            If UCase$(Trim$(globalData(i)(0))) = code Then
                sh.Cells(r, 3).Value = globalData(i)(1)
                sh.Cells(r, 5).Value = Val(globalData(i)(2))
                sh.Cells(r, 11).Value = globalData(i)(8)
                sh.Cells(r, 14).Value = Val(globalData(i)(3))
                Exit For
            End If
        End If
    Next i

    Set wbImg = GetHiddenImageWorkbook()
    If wbImg Is Nothing Then
        RecalculateRow sh, r
        Exit Sub
    End If

    Set ws = wbImg.Sheets(1)
    Set found = ws.Range("A:A").Find(code, LookAt:=xlWhole)

    If Not found Is Nothing Then
        For Each shp In ws.Shapes
            If Not Intersect(shp.TopLeftCell, found.Offset(0, 1)) Is Nothing Then
                shp.Copy
                sh.Paste

                Set c = sh.Cells(r, 2)

                With sh.Shapes(sh.Shapes.Count)
                    .LockAspectRatio = msoTrue

                    If .Width > c.Width Then .Width = c.Width - 2
                    If .Height > c.Height Then .Height = c.Height - 2

                    .Left = c.Left + (c.Width - .Width) / 2
                    .Top = c.Top + (c.Height - .Height) / 2
                End With

                Application.CutCopyMode = False
                wbImg.Saved = True
                Exit For
            End If
        Next shp
    End If

    RecalculateRow sh, r
End Sub

Public Function IsSummarySheet(ByVal sh As Object) As Boolean
    IsSummarySheet = (UCase$(Trim$(sh.Name)) = "SUMMERY")
End Function

Public Function IsBathSheet(ByVal sh As Object) As Boolean
    IsBathSheet = (InStr(1, UCase$(sh.Name), "BATH", vbTextCompare) > 0)
End Function

Private Function GetBathStartRow(ByVal sheetName As String) As Long
    Dim s As String

    s = UCase$(Trim$(sheetName))

    If s Like "*BATH 1*" Or s Like "*BATH1*" Then
        GetBathStartRow = 10
    Else
        GetBathStartRow = 3
    End If
End Function

Private Function GetSummarySheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsSummarySheet(ws) Then
            Set GetSummarySheet = ws
            Exit Function
        End If
    Next ws
End Function

Public Sub RebuildSummaryFast()
    If vbaFeaturesEnabled = False Then Exit Sub

    On Error GoTo safeExit

    Dim sumWs As Worksheet
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim ws As Worksheet
    Dim startRow As Long
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim outArr() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim codeText As String
    Dim lastClearRow As Long

    Set sumWs = GetSummarySheet()
    If sumWs Is Nothing Then Exit Sub

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    lastClearRow = sumWs.Cells(sumWs.Rows.Count, "J").End(xlUp).Row
    If lastClearRow < 2 Then lastClearRow = 2

    rowCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If IsBathSheet(ws) Then
            startRow = GetBathStartRow(ws.Name)
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            If lastRow >= startRow Then
                dataArr = ws.Range("A" & startRow & ":D" & lastRow).Value2

                For i = 1 To UBound(dataArr, 1)
                    codeText = UCase$(Trim$(CStr(dataArr(i, 1))))
                    If codeText <> vbNullString And codeText <> "CODE" And codeText <> "TOTAL" Then
                        rowCount = rowCount + 1
                    End If
                Next i
            End If
        End If
    Next ws

    sumWs.Range("J2:K" & lastClearRow).ClearContents

    If rowCount = 0 Then GoTo safeExit

    ReDim outArr(1 To rowCount, 1 To 2)
    rowCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If IsBathSheet(ws) Then
            startRow = GetBathStartRow(ws.Name)
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            If lastRow >= startRow Then
                dataArr = ws.Range("A" & startRow & ":D" & lastRow).Value2

                For i = 1 To UBound(dataArr, 1)
                    codeText = UCase$(Trim$(CStr(dataArr(i, 1))))
                    If codeText <> vbNullString And codeText <> "CODE" And codeText <> "TOTAL" Then
                        rowCount = rowCount + 1
                        outArr(rowCount, 1) = dataArr(i, 1)
                        outArr(rowCount, 2) = dataArr(i, 4)
                    End If
                Next i
            End If
        End If
    Next ws

    sumWs.Range("J2").Resize(rowCount, 2).Value = outArr

safeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Sub

Private Function SummaryBathKey(ByVal sheetName As String) As String
    SummaryBathKey = SUMMARY_KEY_PREFIX & UCase$(Trim$(sheetName))
End Function

Private Function IsSummaryBathKey(ByVal txt As String) As Boolean
    IsSummaryBathKey = (Left$(UCase$(Trim$(txt)), Len(SUMMARY_KEY_PREFIX)) = SUMMARY_KEY_PREFIX)
End Function

Private Function FindBathSheetByKey(ByVal keyText As String) As Worksheet
    Dim keyName As String
    Dim ws As Worksheet

    If Not IsSummaryBathKey(keyText) Then Exit Function

    keyName = Mid$(keyText, Len(SUMMARY_KEY_PREFIX) + 1)

    For Each ws In ThisWorkbook.Worksheets
        If IsBathSheet(ws) Then
            If UCase$(Trim$(ws.Name)) = keyName Then
                Set FindBathSheetByKey = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function GetSummaryTermsRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    On Error Resume Next
    Set foundCell = ws.Columns("A").Find(What:="TERMS AND CONDITIONS", After:=ws.Cells(ws.Rows.Count, "A"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0

    If Not foundCell Is Nothing Then
        GetSummaryTermsRow = foundCell.Row
    Else
        GetSummaryTermsRow = 10
    End If
End Function

Private Function GetSummaryTotalRow(ByVal ws As Worksheet, ByVal termsRow As Long) As Long
    Dim r As Long

    For r = 2 To termsRow - 1
        If UCase$(Trim$(CStr(ws.Cells(r, "G").Value))) = SUMMARY_TOTAL_KEY Then
            GetSummaryTotalRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindSummaryRowByKey(ByVal ws As Worksheet, ByVal keyText As String, ByVal searchStartRow As Long, ByVal searchEndRow As Long) As Long
    Dim r As Long

    For r = searchStartRow To searchEndRow
        If UCase$(Trim$(CStr(ws.Cells(r, "G").Value))) = UCase$(Trim$(keyText)) Then
            FindSummaryRowByKey = r
            Exit Function
        End If
    Next r
End Function

Private Sub MoveSummaryRow(ByVal ws As Worksheet, ByVal sourceRow As Long, ByVal targetRow As Long)
    If sourceRow = targetRow Then Exit Sub
    ws.Rows(sourceRow).Cut
    ws.Rows(targetRow).Insert Shift:=xlDown
    Application.CutCopyMode = False
End Sub

Private Sub SetBathDescription(ByVal ws As Worksheet, ByVal newText As String)
    If UCase$(Trim$(ws.Name)) = "BATH 1" Then
        ws.Range("C8").Value = newText
    Else
        ws.Range("C1").Value = newText
    End If
End Sub

Private Function GetBathDescription(ByVal ws As Worksheet) As String
    If UCase$(Trim$(ws.Name)) = "BATH 1" Then
        GetBathDescription = Trim$(CStr(ws.Range("C8").Value))
    Else
        GetBathDescription = Trim$(CStr(ws.Range("C1").Value))
    End If
End Function

Private Function GetBathTotalRowNumber(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = ws.Columns("F").Find(What:="TOTAL", After:=ws.Cells(ws.Rows.Count, "F"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If foundCell Is Nothing Then
        Set foundCell = ws.UsedRange.Find(What:="TOTAL", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End If

    On Error GoTo 0

    If Not foundCell Is Nothing Then
        GetBathTotalRowNumber = foundCell.Row
    Else
        GetBathTotalRowNumber = 0
    End If
End Function

Private Sub UpdateSummaryBathRow(ByVal sumWs As Worksheet, ByVal rowNum As Long, ByVal bathWs As Worksheet, ByVal serialNum As Long)
    Dim totalRowNum As Long

    sumWs.Cells(rowNum, "A").Value = serialNum
    sumWs.Cells(rowNum, "B").Value = GetBathDescription(bathWs)
    sumWs.Cells(rowNum, "G").Value = SummaryBathKey(bathWs.Name)

    totalRowNum = GetBathTotalRowNumber(bathWs)

    If totalRowNum > 0 Then
        sumWs.Cells(rowNum, "C").Value = bathWs.Cells(totalRowNum, "G").Value
        sumWs.Cells(rowNum, "E").Value = bathWs.Cells(totalRowNum, "L").Value
        sumWs.Cells(rowNum, "F").Value = bathWs.Cells(totalRowNum, "M").Value
    Else
        sumWs.Cells(rowNum, "C").ClearContents
        sumWs.Cells(rowNum, "E").ClearContents
        sumWs.Cells(rowNum, "F").ClearContents
    End If
End Sub

Private Sub RefreshSummaryTotalRow(ByVal sumWs As Worksheet, ByVal totalRow As Long)
    If totalRow < 2 Then Exit Sub

    sumWs.Cells(totalRow, "A").ClearContents
    sumWs.Cells(totalRow, "B").Value = "TOTAL"
    sumWs.Cells(totalRow, "G").Value = SUMMARY_TOTAL_KEY
    sumWs.Cells(totalRow, "C").Formula = "=SUM(C2:C" & totalRow - 1 & ")"
    sumWs.Cells(totalRow, "E").Formula = "=SUM(E2:E" & totalRow - 1 & ")"
    sumWs.Cells(totalRow, "F").Formula = "=SUM(F2:F" & totalRow - 1 & ")"
End Sub

Public Sub RebuildSummaryAdvanced()

    If vbaFeaturesEnabled = False Then Exit Sub

    If isRebuildingSummary Then Exit Sub
    isRebuildingSummary = True

    On Error GoTo safeExit

    Dim sumWs As Worksheet
    Dim bathWs As Worksheet
    Dim termsRow As Long
    Dim totalRow As Long
    Dim desiredRow As Long
    Dim existingRow As Long
    Dim scanRow As Long
    Dim keyText As String
    Dim linkedWs As Worksheet
    Dim serialNum As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    Set sumWs = GetSummarySheet()
    If sumWs Is Nothing Then GoTo safeExit

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    termsRow = GetSummaryTermsRow(sumWs)
    desiredRow = 2
    serialNum = 1

    For Each bathWs In ThisWorkbook.Worksheets
        If IsBathSheet(bathWs) Then

            existingRow = FindSummaryRowByKey(sumWs, SummaryBathKey(bathWs.Name), desiredRow, termsRow - 1)

            If existingRow = 0 Then
                sumWs.Rows(desiredRow).Insert Shift:=xlDown
                totalRow = totalRow + 1
                termsRow = termsRow + 1
            ElseIf existingRow <> desiredRow Then
                MoveSummaryRow sumWs, existingRow, desiredRow
            End If

            UpdateSummaryBathRow sumWs, desiredRow, bathWs, serialNum

            desiredRow = desiredRow + 1
            serialNum = serialNum + 1
        End If
    Next bathWs

    For scanRow = termsRow - 1 To desiredRow Step -1
        keyText = Trim$(CStr(sumWs.Cells(scanRow, "G").Value))

        If IsSummaryBathKey(keyText) Then
            Set linkedWs = FindBathSheetByKey(keyText)

            If linkedWs Is Nothing Then
                sumWs.Rows(scanRow).Delete
                totalRow = totalRow - 1
                termsRow = termsRow - 1
            End If
        End If
    Next scanRow

    For scanRow = 2 To termsRow - 1
        keyText = Trim$(CStr(sumWs.Cells(scanRow, "G").Value))
        If IsSummaryBathKey(keyText) Then
            sumWs.Cells(scanRow, "A").Value = scanRow - 1
        End If
    Next scanRow

    ApplyTotalRowFormula sumWs

safeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    isRebuildingSummary = False

End Sub

Public Function NeedAdvancedSummaryRefresh(ByVal ws As Worksheet, ByVal Target As Range) As Boolean
    Dim totalRowNum As Long

    If Not IsBathSheet(ws) Then Exit Function

    If UCase$(Trim$(ws.Name)) = "BATH 1" Then
        If Not Intersect(Target, ws.Range("C8")) Is Nothing Then
            NeedAdvancedSummaryRefresh = True
            Exit Function
        End If
    Else
        If Not Intersect(Target, ws.Range("C1")) Is Nothing Then
            NeedAdvancedSummaryRefresh = True
            Exit Function
        End If
    End If

    If Not Intersect(Target, ws.Columns("A")) Is Nothing Then
        NeedAdvancedSummaryRefresh = True
        Exit Function
    End If

    totalRowNum = GetBathTotalRowNumber(ws)
    If totalRowNum > 0 Then
        If Not Intersect(Target, ws.Range("F" & totalRowNum & ":M" & totalRowNum)) Is Nothing Then
            NeedAdvancedSummaryRefresh = True
            Exit Function
        End If
    End If
End Function

Public Sub HandleSummarySheetChange(ByVal sh As Worksheet, ByVal Target As Range)
    If vbaFeaturesEnabled = False Then Exit Sub
    If Not IsSummarySheet(sh) Then Exit Sub

    Dim changedB As Range
    Dim c As Range
    Dim keyText As String
    Dim bathWs As Worksheet
    Dim totalRow As Long
    Dim termsRow As Long

    Set changedB = Intersect(Target, sh.Columns("B"))

    If Not changedB Is Nothing Then
        For Each c In changedB.Cells
            keyText = Trim$(CStr(sh.Cells(c.Row, "G").Value))

            If IsSummaryBathKey(keyText) Then
                Set bathWs = FindBathSheetByKey(keyText)

                If Not bathWs Is Nothing Then
                    SetBathDescription bathWs, CStr(c.Value)
                End If
            End If
        Next c
    End If

    termsRow = GetSummaryTermsRow(sh)
    totalRow = GetSummaryTotalRow(sh, termsRow)

    If totalRow > 0 Then
        RefreshSummaryTotalRow sh, totalRow
    End If
End Sub

Public Sub ScheduleSummaryRefreshAfterStructureChange()
    On Error Resume Next
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!RefreshSummariesAfterStructureChange"
End Sub

Public Sub RefreshSummariesAfterStructureChange()
    On Error Resume Next

    If vbaFeaturesEnabled = False Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    RebuildSummaryFast
    RebuildSummaryAdvanced

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub RemoveDuplicateTotalRows(ByVal ws As Worksheet, ByRef totalRow As Long, ByRef termsRow As Long)
    Dim r As Long
    Dim firstFound As Boolean

    firstFound = False

    For r = termsRow - 1 To 2 Step -1
        If UCase$(Trim$(CStr(ws.Cells(r, "G").Value))) = SUMMARY_TOTAL_KEY Then

            If Not firstFound Then
                totalRow = r
                firstFound = True
            Else
                ws.Rows(r).Delete
                termsRow = termsRow - 1
            End If

        End If
    Next r
End Sub

Public Sub ApplyTotalRowFormula(ByVal ws As Worksheet)

    Dim r As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For r = 2 To lastRow

        If UCase(Trim(ws.Cells(r, "B").Value)) = "TOTAL" Then

            ws.Cells(r, "C").Formula = "=SUM(C2:C" & r - 1 & ")"
            ws.Cells(r, "E").Formula = "=SUM(E2:E" & r - 1 & ")"
            ws.Cells(r, "F").Formula = "=SUM(F2:F" & r - 1 & ")"

        End If

    Next r

End Sub

Private Function FindUserStatusFresh(ByVal userName As String, ByRef foundStatus As String) As Boolean
    Dim data As String
    Dim arr As Variant
    Dim i As Long
    Dim c As Variant
    Dim rowUser As String

    data = GetUsersCSV()
    If data = vbNullString Then Exit Function

    data = NormalizeTextLines(data)
    arr = Split(data, vbLf)

    For i = 1 To UBound(arr)
        If Trim$(arr(i)) <> vbNullString Then
            c = Split(arr(i), ",")

            If UBound(c) >= 2 Then
                rowUser = CleanCsvValue(CStr(c(0)))

                If UCase$(rowUser) = UCase$(Trim$(userName)) Then
                    foundStatus = UCase$(CleanCsvValue(CStr(c(2))))
                    FindUserStatusFresh = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


