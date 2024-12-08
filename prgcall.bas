Attribute VB_Name = "prgcall"

'
'
Option Base 0

Const REDIMCONTINGENCY As Integer = 10

Const TEAMDBVACANT As Integer = 1

    
    
Dim ServerName As String
'------------------------------
'--- from 大会設定 --
'---
Public EventName As String
Public EventDate As String
Public EventVenue As String
Public EventNo As Integer
Dim MaxLaneNo4Heat As Integer
Dim MaxLaneNo4TimeFinal As Integer
Dim MaxLaneNo4Final As Integer
Dim MaxLaneNo4SemiFinal As Integer
'-------------------------
' full file name of mdb
'----------------------------


Dim GlobalError As Integer
'------------
' class table
'-----------
Dim className() As String

'------------------------
'--- from 選手マスター ---


Dim NumRaces As Integer
Dim MaxPrgNo As Integer
Dim SwimmerName() As String
Dim SwimmerNameKANA() As String
Dim BelongsTo() As Integer
Dim NumSwimmers As Integer




'-------------------------
'--  from プログラム ---
'
Dim RaceNobyUID() As Integer
Dim ShumokubyUID() As Integer
Dim DistancebyUID() As Integer
Dim UIDFromRaceNo() As Integer
Dim ClassNumberbyUID() As Integer
Dim ClassNamebyUID() As String
Dim GenderbyUID() As Integer

Dim Phase() As String ' such as 予選/

Dim MaxClassNumber As Integer

Dim Winner() As Integer   ' winner(uid, swimOrder, position)
Dim WinnerTime() As Long  'winnerTime(uid, position)
Dim Rank() As Integer   ' rank(uid,1) is always 1 rank(uid,n) is normally "n"
                     ' but can be less than "n" if there are more
                     ' than two swimmers recorded the same time.
                     
Dim swimmer()



                    
'-------------------------
' from リレーチーム
'----------------------
Dim NumTeam As Integer  ' for relay team not swimming club.
Dim TeamName4Relay() As String
                    





Const FEMALE As Integer = 2
Const MALE As Integer = 1
Const KONSEI As Integer = 3
Const KONGOU As Integer = 4
Const TIME4DNS As Long = 999999
Const TIME4DQ As Long = 999998
Const DNS As Integer = 1
Const DQ As Integer = 2

Dim GenderStr(4) As String
Dim Yoketsu() As String
Const NUMSTYLE As Integer = 7
Dim ShumokuTable(NUMSTYLE) As String
Const NUMDISTANCE As Integer = 7
Dim DistanceTable(NUMDISTANCE) As String

'---------------------------------
'for LocateTeamID  since database table 所属 is not reliable.
'----------------------------------
Dim MaxTeamNum As Integer
Dim Team(100) As String

Dim lastPrgNo As Integer
Dim firstPrgNo As Integer


Sub ReadServer()


    ServerName = Range("serverName").Value

    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim myCon As ADODB.Connection
    Dim row As Long
    Dim col As Long
    
    On Error GoTo MyError
    Set myCon = New ADODB.Connection
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open

    
    myquery = "SELECT 大会番号, 大会名1, 始期間, 終期間, 開催地 FROM 大会設定"
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    row = Range("startRow").row
    col = Range("大会番号").Column
    Do Until myRecordSet.EOF
        
        Cells(row, col).Value = myRecordSet!大会番号
        Cells(row, col + 1).Value = myRecordSet!大会名1
        If myRecordSet!始期間 = myRecordSet!終期間 Then
            Cells(row, col + 2).Value = myRecordSet!始期間
        Else
            Cells(row, col + 2).Value = myRecordSet!始期間 + "～" + myRecordSet!終期間
        End If
        Cells(row, col + 3).Value = myRecordSet!開催地
        row = row + 1
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing

    Exit Sub
MyError:
    MsgBox ("cannot access server " & ServerName)

    
    
End Sub


Public Sub popup(message As String)
    Dim WSH As Object
    Set WSH = CreateObject("WScript.Shell")
    WSH.popup message, 1, "Information", vbInformation
    Set WSH = Nothing
End Sub



Sub clear_xref()
  Range("A6:F999").ClearContents
  ActiveWindow.ScrollRow = 1
  

End Sub

     
Sub create_mdb_event_xref()
    Dim buf As String
    Dim rowNumber As Integer
    Dim fileName As String
    Call clear_xref
    buf = Dir(Range("dataBasePath").Value + "\*.mdb")
    rowNumber = 6
    Do While Len(buf) > 0
      Cells(rowNumber, 2).Value = buf
      ReadEventTable ()
      Cells(rowNumber, 3).Value = EventName
      Cells(rowNumber, 4).Value = EventDate
      Cells(rowNumber, 5).Value = EventVenue
      rowNumber = rowNumber + 1
      buf = Dir()
    Loop
End Sub







Sub ReadDataBase()
      
    Call ReadEventTable
    Call ReadTeamTable
    Call ReadClassTable
    Call ReadSwimmerTable
    Call ReadProgramTable

End Sub




Sub GoAhead()

    EventNo = Cells(Selection.row, 2).Value

    Call InitTables
    Call ReadDataBase
    Range("大会名").Value = EventName
    frmMain.Caption = EventName
    frmMain.show
End Sub




'-------------
'
Function get_directory_path(init_path As String) As String
  Dim fd As FileDialog
  Set fd = Application.FileDialog(msoFileDialogFolderPicker)
  With fd
    .ButtonName = "Select"
    .InitialFileName = init_path
    With .Filters
      .Clear
    End With

    If .show = True Then
      get_directory_path = .SelectedItems(1)
    End If
  End With
End Function






Private Sub init_class_db_array(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    myquery = "select クラス番号  from クラス where 大会番号=" & _
    EventNo & " and クラス番号=(select max(クラス番号) from クラス where 大会番号= " & _
     EventNo & ")"

    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    If myRecordSet.RecordCount > 0 Then
        MaxClassNumber = myRecordSet!クラス番号
        ReDim className(MaxClassNumber)
    Else
        MaxClassNumber = 1
        ReDim className(1)
    End If
    
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Sub ReadClassTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
   
    Dim myquery As String
    
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    
    Call init_class_db_array(myCon)
        
    myquery = "SELECT クラス番号,クラス名称 FROM クラス where 大会番号= " & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    
    Do Until myRecordSet.EOF
      className(myRecordSet!クラス番号) = myRecordSet!クラス名称

      myRecordSet.MoveNext
    Loop
        
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub

Private Sub init_swimmer_db_array(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    myquery = "select 選手番号  from 選手 where 大会番号= " & _
      EventNo & " and 選手番号=(select max(選手番号) from 選手 where " & _
       "大会番号= " & EventNo & ");"

    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    NumSwimmers = myRecordSet!選手番号
    ReDim SwimmerName(NumSwimmers)
    ReDim SwimmerNameKANA(NumSwimmers)
    ReDim BelongsTo(NumSwimmers)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Sub ReadSwimmerTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
   
    Dim myquery As String
    Dim clubName As String
    
    Dim clubNo As Integer
        
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    Call init_swimmer_db_array(myCon)
        
    myquery = "SELECT 選手番号, 氏名, 氏名カナ, 所属名称１ FROM 選手 where 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic

    SwimmerName(0) = ""
    BelongsTo(0) = 0
    MaxTeamNum = 0

    myRecordSet.MoveFirst
    Do Until myRecordSet.EOF
      
      clubNo = LocateTeamID(myRecordSet("所属名称1"))
     
      SwimmerName(myRecordSet!選手番号) = RTrim(myRecordSet!氏名)
      SwimmerNameKANA(myRecordSet!選手番号) = RTrim(myRecordSet!氏名カナ)

      BelongsTo(myRecordSet!選手番号) = clubNo

      myRecordSet.MoveNext
    Loop
      
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub




Sub ReadEventTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
       
    Dim myquery As String
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"

    myCon.Open
        
    myquery = "SELECT 大会名１,開催地,始期間,終期間,使用水路予選,使用水路タイム決勝,使用水路決勝,使用水路準決勝 " & _
     " FROM 大会設定 where 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
  
    EventName = Object2String(myRecordSet!大会名1)
    If EventName = "" Then
      EventDate = ""
      EventVenue = ""
    Else
    
      If myRecordSet!始期間 <> myRecordSet!終期間 Then
        EventDate = myRecordSet!始期間 & "～" & myRecordSet!終期間
      Else
        EventDate = myRecordSet!始期間
      End If
      EventVenue = myRecordSet!開催地
      MaxLaneNo4Heat = myRecordSet!使用水路予選
      MaxLaneNo4TimeFinal = myRecordSet!使用水路タイム決勝
      MaxLaneNo4Final = myRecordSet!使用水路決勝
      MaxLaneNo4SemiFinal = myRecordSet!使用水路準決勝
      
    End If
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub










Function is_relay(uid As Integer) As Boolean
 
  If (ShumokubyUID(uid) > 5) Then
    is_relay = True
  Else
    is_relay = False
  End If
  
End Function





Private Sub RedimProgramDBArray(maxuid As Integer)

    ReDim ClassNumberbyUID(maxuid)
    ReDim Phase(maxuid)
    ReDim DistancebyUID(maxuid)
    ReDim GenderbyUID(maxuid)
    ReDim ShumokubyUID(maxuid)
    ReDim RaceNobyUID(maxuid)
End Sub



Private Sub InitProgramDBArray(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim maxuid As Integer
    myquery = "select 競技番号 from プログラム where " & _
     "大会番号= " & EventNo & " and 競技番号=(select max(競技番号) from プログラム " & _
      "where 大会番号= " & EventNo & ");"

    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    NumRaces = myRecordSet!競技番号

    Call RedimProgramDBArray(NumRaces)
    myRecordSet.Close
    Set myRecordSet = Nothing
    myquery = "select 表示用競技番号 from プログラム where " & _
     "大会番号=" & EventNo & " and 表示用競技番号=(select max(表示用競技番号) from プログラム" & _
      " where 大会番号= " & EventNo & ");"
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    MaxPrgNo = myRecordSet!表示用競技番号
    ReDim UIDFromRaceNo(MaxPrgNo)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Sub ReadProgramTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim uid As Integer
    Dim myquery As String
    
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    Call InitProgramDBArray(myCon)
    myquery = "SELECT 競技番号 as uid, 表示用競技番号, 種目コード, 距離コード,  " + _
              "性別コード, 予決コード, クラス番号  " + _
              "FROM プログラム where 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    
    Do Until myRecordSet.EOF
        uid = myRecordSet!uid
      RaceNobyUID(uid) = CInt(myRecordSet!表示用競技番号)
      UIDFromRaceNo(CInt(myRecordSet("表示用競技番号"))) = uid

      ShumokubyUID(uid) = myRecordSet!種目コード
      GenderbyUID(uid) = CInt(myRecordSet!性別コード)
      DistancebyUID(uid) = myRecordSet!距離コード
      Phase(uid) = Yoketsu(myRecordSet!予決コード)
      ClassNumberbyUID(uid) = CInt(myRecordSet!クラス番号)
      myRecordSet.MoveNext
    Loop
    
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub


Function GetMaxLaneNo(uid As Integer) As Integer

    If Phase(uid) = "タイム決勝" Then
        GetMaxLaneNo = MaxLaneNo4TimeFinal
        Exit Function
    End If
    If Phase(uid) = "予選" Then
        GetMaxLaneNo = MaxLaneNo4Heat
        Exit Function
    End If
    If Phase(uid) = "準決勝" Then
        GetMaxLaneNo = MaxLaneNo4SemiFinal
        Exit Function
    End If
    GetMaxLaneNo = MaxLaneNo4Final
End Function




Sub InitGenderString()
  GenderStr(MALE) = "男子"
  GenderStr(FEMALE) = "女子"
  GenderStr(KONSEI) = "混成"
  GenderStr(KONGOU) = "混合"
End Sub

Function LocateStyleNumber(thisShumoku As String) As Integer
  Dim cnt As Integer
  
  For cnt = 1 To NUMSTYLE
    If ShumokuTable(cnt) = thisShumoku Then
      LocateStyleNumber = cnt
      Exit Function
    End If
  Next cnt
  MsgBox ("error in LocateStyleNumber")
  LocateStyleNumber = 0
End Function


Function LocateDistanceNumber(thisDistance As String) As Integer
  Dim cnt As Integer
  
  For cnt = 1 To NUMSTYLE
    If DistanceTable(cnt) = thisDistance Then
      LocateDistanceNumber = cnt
      Exit Function
    End If
  Next cnt
  MsgBox ("error in LocateDistanceNumber")
  LocateDistanceNumber = 0
End Function

Sub InitStyleTable()
  ShumokuTable(1) = "自由形"
  ShumokuTable(2) = "背泳ぎ"
  ShumokuTable(3) = "平泳ぎ"
  ShumokuTable(4) = "バタフライ"
  ShumokuTable(5) = "個人メドレー"
  ShumokuTable(6) = "リレー"
  ShumokuTable(7) = "メドレーリレー"
End Sub
Sub InitDistanceTable()
  DistanceTable(1) = "  25m"
  DistanceTable(2) = "  50m"
  DistanceTable(3) = " 100m"
  DistanceTable(4) = " 200m"
  DistanceTable(5) = " 400m"
  DistanceTable(6) = " 800m"
  DistanceTable(7) = "1500m"
End Sub

Sub InitTables()
  Call InitStyleTable
  Call InitDistanceTable
  Call InitGenderString
  Call InitYoketsu
  
  
End Sub

Private Sub InitYoketsuArray(myCon As ADODB.Connection)

    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    myquery = "select 予決コード, 予決 from 予決 where 予決コード=(select max(予決コード) from 予決);"

    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    
    If myRecordSet.RecordCount > 0 Then
      
      ReDim Yoketsu(myRecordSet!予決コード)
    Else
      GlobalError = GlobalError Or TEAMDBVACANT
      
    End If
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub


Sub InitYoketsu()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim mySQL As String

      
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    
    Call InitYoketsuArray(myCon)

      mySQL = "SELECT 予決コード,予決 FROM 予決"
      
      myRecordSet.Open mySQL, myCon, adOpenStatic, adLockOptimistic

      Yoketsu(0) = ""
      Do Until myRecordSet.EOF

        Yoketsu(myRecordSet!予決コード) = myRecordSet!予決

        myRecordSet.MoveNext
      Loop
       
      myRecordSet.Close
      Set myRecordSet = Nothing
    
    myCon.Close
    Set myCon = Nothing

End Sub



Private Sub InitTeamDbArray(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    myquery = "select チーム番号 from リレーチーム where 大会番号= " & EventNo & _
      " and チーム番号=(select max(チーム番号) from リレーチーム " & _
       " where 大会番号= " & EventNo & ");"

    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    
    If myRecordSet.RecordCount > 0 Then
      NumTeam = myRecordSet!チーム番号
      ReDim TeamName4Relay(NumTeam)
    Else
      GlobalError = GlobalError Or TEAMDBVACANT
      
    End If
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub


Sub ReadTeamTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim mySQL As String

      
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    
    Call InitTeamDbArray(myCon)
    If (GlobalError And TEAMDBVACANT) = 0 Then
      mySQL = "SELECT チーム番号,チーム名 FROM リレーチーム where 大会番号=" & EventNo
      
      myRecordSet.Open mySQL, myCon, adOpenStatic, adLockOptimistic

      TeamName4Relay(0) = ""
      Do Until myRecordSet.EOF

        TeamName4Relay(myRecordSet!チーム番号) = myRecordSet!チーム名

        myRecordSet.MoveNext
      Loop
       
      myRecordSet.Close
      Set myRecordSet = Nothing
    End If
    myCon.Close
    Set myCon = Nothing

End Sub




Function LocateTeamID(teamName As String) As Integer
  Dim teamNum As Integer
  Team(0) = ""
  For teamNum = 1 To MaxTeamNum
    If Team(teamNum) = teamName Then
      LocateTeamID = teamNum
      Exit Function
    End If
  Next teamNum
  Team(teamNum) = teamName

  MaxTeamNum = teamNum
  LocateTeamID = teamNum
End Function





Sub ClearSwimmerList()
    Dim lastRow As Long
    Sheets("選手登録番号リスト").Select
    lastRow = Cells(Rows.Count, 1).End(xlup).row
    Range("A3:C" + CStr(lastRow)).Select
    Selection.ClearContents
    
End Sub



Sub CreateSwimmerList()

    Dim row As Long
    Call ClearSwimmerList



    Call ReadEventTable
    Call ReadSwimmerTable
    Cells(1, 1).Value = EventName
    For i = 1 To NumSwimmers
      row = i + 2
      Cells(row, 1).Value = i
      Cells(row, 2).Value = SwimmerName(i)
      Cells(row, 3).Value = Team(BelongsTo(i))
    Next i
    Cells(3, 1).Select
    
    
End Sub

Sub SetSwimmerListHeader()

    Cells(2, 1).Value = "登録番号"
    Cells(2, 2).Value = "選手名"
    Cells(2, 3).Value = "所属"
    Rows("2:2").AutoFilter

End Sub




Sub CloseSQLConnection(myRecordSet As ADODB.Recordset, myCon As ADODB.Connection)
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub




Function RaceExist(uid As Integer, kumi As Integer) As Boolean

    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
     
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
           
    myquery = "SELECT 競技番号, 選手番号,  第１泳者, 第２泳者, 第３泳者, 第４泳者  " & _
              ", 組, 水路, 事由入力ステータス " & _
              "FROM 記録 WHERE 組=" & kumi & " AND 競技番号=" & uid & _
              " and 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    
    If myRecordSet.EOF Then
      RaceExist = False
    Else
      RaceExist = True
    End If
    Call CloseSQLConnection(myRecordSet, myCon)

End Function
Function Object2Int(obj As Variant) As Integer
    If IsNull(obj) Then
        Object2Int = 0
    Else
        Object2Int = obj
    End If
End Function
Function Object2String(obj As Variant) As String
    If IsNull(obj) Then
        Object2String = ""
    Else
        Object2String = obj
    End If
End Function
Private Sub show(uid As Integer, kumi As Integer)
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
   
    Dim myquery As String
    Dim laneNoStr As String
    Dim lane0NoStr As String
    Dim furiganaStr As String
    Dim furigana0Str As String
    Dim swimmer As Integer
    Dim relaySwimmer(5) As Integer
    Dim maxLaneNumber As Integer
    Dim minLaneNumber As Integer
    Dim laneNo As Integer
    maxLaneNumber = 0
    minLaneNumber = GetMaxLaneNo(uid)
    Call frmMain.hide_unused_lane(minLaneNumber)
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
         
    myquery = "SELECT 競技番号, 選手番号,  " & _
              "第１泳者, 第２泳者, 第３泳者, 第４泳者  " & _
              ", 組, 水路, 事由入力ステータス " & _
              "FROM 記録 WHERE 組=" & kumi & " AND 競技番号=" & uid & _
              " and 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    frmMain.clearMe

    frmMain.lblClassName.Caption = className(ClassNumberbyUID(uid))
    frmMain.lblGender.Caption = GenderStr(GenderbyUID(uid))
    frmMain.lblRaceName.Caption = DistanceTable(DistancebyUID(uid)) + ShumokuTable(ShumokubyUID(uid))
    frmMain.lblPhase.Caption = Phase(uid)
    Do Until myRecordSet.EOF
      laneNo = CInt(myRecordSet!水路)
      If laneNo = 0 Then
        laneNo = 10
      End If

      laneNoStr = "lblName" & laneNo
      furiganaStr = "lblKana" & laneNo

      swimmer = Object2Int(myRecordSet!選手番号)
      If swimmer > 0 Then
        If laneNo > maxLaneNumber Then
          maxLaneNumber = laneNo
        End If
        If laneNo < minLaneNumber Then
          minLaneNumber = laneNo
        End If
        If is_relay(uid) Then
          frmMain.Controls(laneNoStr).Caption = TeamName4Relay(swimmer)
          If myRecordSet!事由入力ステータス = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "棄権"
          Else
            
            relaySwimmer(1) = Object2Int(myRecordSet!第1泳者)
            relaySwimmer(2) = Object2Int(myRecordSet!第2泳者)
            relaySwimmer(3) = Object2Int(myRecordSet!第3泳者)
            relaySwimmer(4) = Object2Int(myRecordSet!第4泳者)
            For swimOrder = 1 To 4
              lane0NoStr = laneNoStr & CStr(swimOrder)
              frmMain.Controls(lane0NoStr).Caption = SwimmerName(relaySwimmer(swimOrder))
              furigana0Str = furiganaStr & CStr(swimOrder)
              frmMain.Controls(furigana0Str).Caption = SwimmerNameKANA(relaySwimmer(swimOrder))
            Next swimOrder
          End If
        Else
        
          frmMain.Controls(laneNoStr).Caption = SwimmerName(swimmer)
          frmMain.Controls(furiganaStr).Caption = SwimmerNameKANA(swimmer)
          lane0NoStr = laneNoStr & "1"
          frmMain.Controls(lane0NoStr).Caption = "(" + Trim(Team(BelongsTo(swimmer))) + ")"
          If myRecordSet!事由入力ステータス = 1 Then
            lane0NoStr = laneNoStr + "2"
            frmMain.Controls(lane0NoStr).Caption = "棄権"
          End If
        End If
      End If
      myRecordSet.MoveNext
    Loop
    Call CloseSQLConnection(myRecordSet, myCon)
    If can_go_with_next(uid, kumi, maxLaneNumber) Then
      frmMain.cmdNextTogether.Visible = True
    Else
      frmMain.cmdNextTogether.Visible = False
    End If
    If CanGoWithPrev(minLaneNumber) Then
      frmMain.cmdPrvTogether.Visible = True
    Else
      frmMain.cmdPrvTogether.Visible = False
    End If
End Sub


Private Sub next_race_show(prevUID As Integer, uid As Integer, kumi As Integer)
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
   
    Dim myquery As String
    Dim laneNo As Integer
    Dim laneNoStr As String
    Dim lane0NoStr As String
    Dim furiganaStr As String
    Dim furigana0Str As String
    Dim swimmer As Integer
    Dim relaySwimmer(5) As Integer
    Dim maxLaneNumber As Integer
    Dim minLaneNumber As Integer
    
    maxLaneNumber = 0
    minLaneNumber = GetMaxLaneNo(uid)

    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
         
    myquery = "SELECT 競技番号, 選手番号,  " & _
              "第１泳者, 第２泳者, 第３泳者, 第４泳者  " & _
              ", 組, 水路, 事由入力ステータス " & _
              "FROM 記録 WHERE 組=" & kumi & " AND 競技番号=" & uid & _
              " and 大会番号=" & EventNo
      
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    If frmMain.lblClassName.Caption <> "合同" Then
    frmMain.lblClassName.Caption = "合同"
      If GenderbyUID(prevUID) <> GenderbyUID(uid) Then
        frmMain.lblGender.Caption = ""
        frmMain.lblClassName1 = className(ClassNumberbyUID(prevUID)) & " " & GenderStr(GenderbyUID(prevUID))
      Else
        frmMain.lblClassName1.Caption = className(ClassNumberbyUID(prevUID))
      End If
      frmMain.lblClassName1.Visible = True
    End If
    Do Until myRecordSet.EOF
      laneNo = CInt(myRecordSet!水路)
      If laneNo = 0 Then
        laneNo = 10
      End If
      laneNoStr = "lblName" & laneNo
      furiganaStr = "lblKana" & laneNo

      swimmer = Object2Int(myRecordSet!選手番号)
      If swimmer > 0 Then
        If laneNo > maxLaneNumber Then
          maxLaneNumber = laneNo
        End If
        If laneNo <= minLaneNumber Then
          minLaneNumber = myRecordSet!水路
          If GenderbyUID(prevUID) <> GenderbyUID(uid) Then
            frmMain.Controls("lblClassName" & minLaneNumber).Caption = className(ClassNumberbyUID(uid)) & " " & _
                                                           GenderStr(GenderbyUID(uid))
          Else
            frmMain.Controls("lblClassName" & minLaneNumber).Caption = className(ClassNumberbyUID(uid))
          End If
          frmMain.Controls("lblClassName" & minLaneNumber).Visible = True
        End If
        If is_relay(uid) Then
          frmMain.Controls(laneNoStr).Caption = TeamName4Relay(swimmer)
          If myRecordSet!事由入力ステータス = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "棄権"
          Else
            
            relaySwimmer(1) = Object2Int(myRecordSet!第1泳者)
            relaySwimmer(2) = Object2Int(myRecordSet!第2泳者)
            relaySwimmer(3) = Object2Int(myRecordSet!第3泳者)
            relaySwimmer(4) = Object2Int(myRecordSet!第4泳者)
            For swimOrder = 1 To 4
              lane0NoStr = laneNoStr & CStr(swimOrder)
              frmMain.Controls(lane0NoStr).Caption = SwimmerName(relaySwimmer(swimOrder))
              furigana0Str = furiganaStr & CStr(swimOrder)
              frmMain.Controls(furigana0Str).Caption = SwimmerNameKANA(relaySwimmer(swimOrder))
            Next swimOrder
          End If
        Else
        
          frmMain.Controls(laneNoStr).Caption = SwimmerName(swimmer)
          frmMain.Controls(furiganaStr).Caption = SwimmerNameKANA(swimmer)
          lane0NoStr = laneNoStr & "1"
          frmMain.Controls(lane0NoStr).Caption = "(" + Trim(Team(BelongsTo(swimmer))) + ")"
          If myRecordSet!事由入力ステータス = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "棄権"
          End If
        End If
      End If
      myRecordSet.MoveNext
    Loop
    Call CloseSQLConnection(myRecordSet, myCon)
    If can_go_with_next(uid, kumi, maxLaneNumber) Then
      frmMain.cmdNextTogether.Visible = True
    Else
      frmMain.cmdNextTogether.Visible = False
    End If
    If CanGoWithPrev(minLaneNumber) Then
      frmMain.cmdPrvTogether.Visible = True
    Else
      frmMain.cmdPrvTogether.Visible = False
    End If
End Sub


Function get_first_occupied_lane(uid As Integer) As Integer
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim minLane As Integer
    Dim swimmer As Integer
    
    minLane = GetMaxLaneNo(uid)
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
         
    myquery = "SELECT UID, 選手番号, 組, 水路 " + _
              "FROM 記録マスター WHERE 大会番号= " & EventNo & _
              " and 組=1 " + " AND UID=" & uid
    myRecordSet.Open myquery, myCon, adOpenStatic, adLockOptimistic
    Do Until myRecordSet.EOF
      swimmer = Object2Int(myRecordSet!選手番号)
      If swimmer > 0 Then
        If CInt(myRecordSet!水路) < minLane Then
          minLane = CInt(myRecordSet!水路)
        End If
      End If
      myRecordSet.MoveNext
    Loop
    get_first_occupied_lane = minLane
    Call CloseSQLConnection(myRecordSet, myCon)
End Function
Function can_go_with_next(uid As Integer, kumi As Integer, maxLaneNumber As Integer) As Boolean
    Dim prgno As Integer
    Dim nextUid As Integer
    can_go_with_next = False
    If maxLaneNumber = GetMaxLaneNo(uid) Then Exit Function
'    If kumi > 1 Then Exit Function
    
    prgno = lastPrgNo + 1
    If prgno > MaxPrgNo Then Exit Function
    
    nextUid = UIDFromRaceNo(prgno)
'    If RaceExist(uid, 2) Then Exit Function
    If DistancebyUID(uid) <> DistancebyUID(nextUid) Then Exit Function
    If ShumokubyUID(uid) <> ShumokubyUID(nextUid) Then Exit Function
    If Phase(uid) <> Phase(nextUid) Then Exit Function
    If maxLaneNumber < get_first_occupied_lane(nextUid) Then
        can_go_with_next = True
    End If
    
End Function
Function CanGoWithPrev(minLaneNumber As Integer) As Boolean
    CanGoWithPrev = False
End Function
Sub ClearArea(area As String)

    Range(area).Select
    Selection.ClearContents
End Sub

Public Sub GetNextRace()
  Dim rc As Boolean

  Dim loopCount As Integer
  Dim prgno As Integer
  Dim group As Integer
  Dim uid As Integer

  Const loopLimit As Integer = 10
  Call frmMain.hide_class
  
  prgno = lastPrgNo     ' frmMain.txtPrgNo.Value
  group = frmMain.txtKumi.Value

  Call InitAndReadDB
  uid = UIDFromRaceNo(prgno)
  loopCount = 0
  Do While rc = False
    group = group + 1
    If RaceExist(uid, group) Then
      frmMain.txtPrgNo.Value = prgno
      frmMain.txtKumi.Value = group
      firstPrgNo = prgno
      lastPrgNo = prgno
      Call show(uid, group)
      Exit Sub
    End If
    prgno = prgno + 1
    If prgno > MaxPrgNo Then
      Call popup("最終レースです。")
      Exit Sub
    End If
    uid = UIDFromRaceNo(prgno)
    group = 0
    loopCount = loopCount + 1
    If loopCount = loopLimit Then
      Call popup("最終レースです。(exceeds loop limit)")
      Exit Sub
    End If
  Loop

End Sub

Public Sub GetPrevRace()
  Dim rc As Boolean

  Dim prgno As Integer
  Dim group As Integer
  Dim uid As Integer
  
  Call frmMain.hide_class
  prgno = firstPrgNo    ' frmMain.txtPrgNo.Value
  group = frmMain.txtKumi.Value

  Call InitAndReadDB

  If (group > 1) Then
      group = group - 1
  Else
      prgno = prgno - 1
      If prgno = 0 Then
        popup ("最初のレースです。")
        Exit Sub
      End If
      If prgno > MaxPrgNo Then
        popup ("該当するレースはありません。最終レースを表示します。")
        prgno = MaxPrgNo
      End If
      uid = UIDFromRaceNo(prgno)
      rc = True
      Do While rc = True
        group = group + 1
        rc = RaceExist(uid, group)
      Loop
      group = group - 1
  End If
  uid = UIDFromRaceNo(prgno)
  
  frmMain.txtPrgNo.Value = prgno
  frmMain.txtKumi.Value = group
  lastPrgNo = prgno
  firstPrgNo = prgno
  Call show(uid, group)
      
End Sub

Public Sub GoWithNextRace()

  Dim loopCount As Integer
  Dim prgno As Integer
  Dim group As Integer
  Dim uid As Integer
  Dim thisUID As Integer
  
  thisUID = UIDFromRaceNo(lastPrgNo)
    
  prgno = lastPrgNo + 1
  lastPrgNo = prgno
  group = 1
  Call InitAndReadDB
  uid = UIDFromRaceNo(prgno)


  Call next_race_show(thisUID, uid, group)

End Sub

Public Sub reset()
    NumSwimmers = 0
    Call ShowLaneOrder
End Sub
Sub InitAndReadDB()
  If NumSwimmers = 0 Then
    Call InitTables
    Call ReadDataBase

  End If
End Sub

Public Sub ShowLaneOrder()

  Dim prgno As Integer
  Dim kumi As Integer
  Dim uid As Integer
  
  prgno = frmMain.txtPrgNo.Value
  kumi = frmMain.txtKumi.Value
  Call frmMain.hide_class
  
  Call InitAndReadDB
  If prgno > MaxPrgNo Then
    popup ("該当のレースはありません。最終レースを表示します。")
    prgno = MaxPrgNo
    frmMain.txtPrgNo.Value = prgno
    frmMain.txtKumi.Value = 1
    kumi = 1
  End If
  uid = UIDFromRaceNo(prgno)

  If Not RaceExist(uid, kumi) Then
    popup ("該当のレースはありません。")
  Else
    lastPrgNo = prgno
    firstPrgNo = prgno
    Call show(uid, kumi)
  End If

End Sub



