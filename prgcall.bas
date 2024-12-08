Attribute VB_Name = "prgcall"

'
'
Option Base 0

Const REDIMCONTINGENCY As Integer = 10

Const TEAMDBVACANT As Integer = 1

    
    
Dim ServerName As String
'------------------------------
'--- from ���ݒ� --
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
' �p��̒�`
' ���Z�ԍ��@: �����N���X������ړ��������ň�̋��Z�B���Z�̒��ɑg������
' race�ԍ�  : ��̋��Z�ɕ����̑g������ꍇ��race�͂��̑g�̐���������B
'      �������A�����̏ꍇ�͕����̑g�����race�ɂȂ�
'
'----------------------------
Dim RealRace() As Integer
'----
' RealRace�͓񎟌��̃A���C�BRealRace(3, ���[�X�̐�)
' ���������[�X�̐��͍����𖳎��������BRealRace(2, rn) �� 0�łȂ��ꍇ�͍������[�X�ƂȂ�B
'  RealRace(2,rn)���������̃��[�X�������Ƃ������ƂɂȂ�B
Dim GlobalError As Integer
'---------    ---
' class table
'-----------
Dim className() As String

'------------------------
'--- from �I��}�X�^�[ ---


Dim numRaces As Integer
Dim MaxPrgNo As Integer
Dim SwimmerName() As String
Dim SwimmerNameKANA() As String
Dim BelongsTo() As Integer
Dim NumSwimmers As Integer




'-------------------------
'--  from �v���O���� ---
'
Dim RaceNobyUID() As Integer
Dim ShumokubyUID() As Integer
Dim DistancebyUID() As Integer
Dim UIDFromRaceNo() As Integer
Dim ClassNumberbyUID() As Integer
Dim ClassNamebyUID() As String
Dim GenderbyUID() As Integer

Dim Phase() As String ' such as �\�I/

Dim MaxClassNumber As Integer

Dim Winner() As Integer   ' winner(uid, swimOrder, position)
Dim WinnerTime() As Long  'winnerTime(uid, position)
Dim Rank() As Integer   ' rank(uid,1) is always 1 rank(uid,n) is normally "n"
                     ' but can be less than "n" if there are more
                     ' than two swimmers recorded the same time.
                     
Dim swimmer()



                    
'-------------------------
' from �����[�`�[��
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
'for LocateTeamID  since database table ���� is not reliable.
'----------------------------------
Dim MaxTeamNum As Integer
Dim Team(200) As String

Dim lastPrgNo As Integer
Dim firstPrgNo As Integer


Sub ReadServer()


    ServerName = Range("serverName").Value

    Dim myRecordSet As New ADODB.Recordset
    Dim myQuery As String
    Dim myCon As ADODB.Connection
    Dim row As Long
    Dim col As Long
    
    On Error GoTo MyError
    Set myCon = New ADODB.Connection
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open

    
    myQuery = "SELECT ���ԍ�, ��1, �n����, �I����, �J�Òn FROM ���ݒ�"
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    row = Range("startRow").row
    col = Range("���ԍ�").Column
    Do Until myRecordSet.EOF
        
        Cells(row, col).Value = myRecordSet!���ԍ�
        Cells(row, col + 1).Value = myRecordSet!��1
        If myRecordSet!�n���� = myRecordSet!�I���� Then
            Cells(row, col + 2).Value = myRecordSet!�n����
        Else
            Cells(row, col + 2).Value = myRecordSet!�n���� + "�`" + myRecordSet!�I����
        End If
        Cells(row, col + 3).Value = myRecordSet!�J�Òn
        row = row + 1
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
    myCon.Close
    Set myCon = Nothing

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
    Call CreateRaceArray
    Range("��").Value = EventName

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
    Dim myQuery As String

    myQuery = "select �N���X�ԍ�  from �N���X where ���ԍ�=" & _
    EventNo & " and �N���X�ԍ�=(select max(�N���X�ԍ�) from �N���X where ���ԍ�= " & _
     EventNo & ")"

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    If myRecordSet.recordCount > 0 Then
        MaxClassNumber = myRecordSet!�N���X�ԍ�
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
   
    Dim myQuery As String
    
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    
    Call init_class_db_array(myCon)
        
    myQuery = "SELECT �N���X�ԍ�,�N���X���� FROM �N���X where ���ԍ�= " & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    
    Do Until myRecordSet.EOF
      className(myRecordSet!�N���X�ԍ�) = myRecordSet!�N���X����

      myRecordSet.MoveNext
    Loop
        
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub

Private Sub init_swimmer_db_array(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myQuery As String

    myQuery = "select �I��ԍ�  from �I�� where ���ԍ�= " & _
      EventNo & " and �I��ԍ�=(select max(�I��ԍ�) from �I�� where " & _
       "���ԍ�= " & EventNo & ");"

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    NumSwimmers = myRecordSet!�I��ԍ�
    ReDim SwimmerName(NumSwimmers)
    ReDim SwimmerNameKANA(NumSwimmers)
    ReDim BelongsTo(NumSwimmers)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Sub ReadSwimmerTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
   
    Dim myQuery As String
    Dim clubName As String
    
    Dim clubNo As Integer
        
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    myCon.CursorLocation = adUseClient
    Call init_swimmer_db_array(myCon)
        
    myQuery = "SELECT �I��ԍ�, ����, �����J�i, �������̂P FROM �I�� where ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic

    SwimmerName(0) = ""
    BelongsTo(0) = 0
    MaxTeamNum = 0

    myRecordSet.MoveFirst
    Do Until myRecordSet.EOF
      
      clubNo = LocateTeamID(myRecordSet("��������1"))
     
      SwimmerName(myRecordSet!�I��ԍ�) = RTrim(myRecordSet!����)
      SwimmerNameKANA(myRecordSet!�I��ԍ�) = RTrim(myRecordSet!�����J�i)

      BelongsTo(myRecordSet!�I��ԍ�) = clubNo

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
       
    Dim myQuery As String
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"

    myCon.Open
        
    myQuery = "SELECT ���P,�J�Òn,�n����,�I����,�g�p���H�\�I,�g�p���H�^�C������,�g�p���H����,�g�p���H������ " & _
     " FROM ���ݒ� where ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
  
    EventName = Object2String(myRecordSet!��1)
    If EventName = "" Then
      EventDate = ""
      EventVenue = ""
    Else
    
      If myRecordSet!�n���� <> myRecordSet!�I���� Then
        EventDate = myRecordSet!�n���� & "�`" & myRecordSet!�I����
      Else
        EventDate = myRecordSet!�n����
      End If
      EventVenue = myRecordSet!�J�Òn
      MaxLaneNo4Heat = myRecordSet!�g�p���H�\�I
      MaxLaneNo4TimeFinal = myRecordSet!�g�p���H�^�C������
      MaxLaneNo4Final = myRecordSet!�g�p���H����
      MaxLaneNo4SemiFinal = myRecordSet!�g�p���H������
      
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
    Dim myQuery As String
    Dim maxuid As Integer
    myQuery = "select ���Z�ԍ� from �v���O���� where " & _
     "���ԍ�= " & EventNo & " and ���Z�ԍ�=(select max(���Z�ԍ�) from �v���O���� " & _
      "where ���ԍ�= " & EventNo & ");"

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    numRaces = myRecordSet!���Z�ԍ�

    Call RedimProgramDBArray(numRaces)
    myRecordSet.Close
    Set myRecordSet = Nothing
    myQuery = "select �\���p���Z�ԍ� from �v���O���� where " & _
     "���ԍ�=" & EventNo & " and �\���p���Z�ԍ�=(select max(�\���p���Z�ԍ�) from �v���O����" & _
      " where ���ԍ�= " & EventNo & ");"
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    MaxPrgNo = myRecordSet!�\���p���Z�ԍ�
    ReDim UIDFromRaceNo(MaxPrgNo)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Sub ReadProgramTable()
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim uid As Integer
    Dim myQuery As String
    
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
    Call InitProgramDBArray(myCon)
    myQuery = "SELECT ���Z�ԍ� as uid, �\���p���Z�ԍ�, ��ڃR�[�h, �����R�[�h,  " + _
              "���ʃR�[�h, �\���R�[�h, �N���X�ԍ�  " + _
              "FROM �v���O���� where ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    
    Do Until myRecordSet.EOF
        uid = myRecordSet!uid
      RaceNobyUID(uid) = CInt(myRecordSet!�\���p���Z�ԍ�)
      UIDFromRaceNo(CInt(myRecordSet("�\���p���Z�ԍ�"))) = uid

      ShumokubyUID(uid) = myRecordSet!��ڃR�[�h
      GenderbyUID(uid) = CInt(myRecordSet!���ʃR�[�h)
      DistancebyUID(uid) = myRecordSet!�����R�[�h
      Phase(uid) = Yoketsu(myRecordSet!�\���R�[�h)
      ClassNumberbyUID(uid) = CInt(myRecordSet!�N���X�ԍ�)
      myRecordSet.MoveNext
    Loop
    
    myRecordSet.Close
    myCon.Close
    Set myRecordSet = Nothing
    Set myCon = Nothing
End Sub


Function GetMaxLaneNo(uid As Integer) As Integer

    If Phase(uid) = "�^�C������" Then
        GetMaxLaneNo = MaxLaneNo4TimeFinal
        Exit Function
    End If
    If Phase(uid) = "�\�I" Then
        GetMaxLaneNo = MaxLaneNo4Heat
        Exit Function
    End If
    If Phase(uid) = "������" Then
        GetMaxLaneNo = MaxLaneNo4SemiFinal
        Exit Function
    End If
    GetMaxLaneNo = MaxLaneNo4Final
End Function




Sub InitGenderString()
  GenderStr(MALE) = "�j�q"
  GenderStr(FEMALE) = "���q"
  GenderStr(KONSEI) = "����"
  GenderStr(KONGOU) = "����"
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
  ShumokuTable(1) = "���R�`"
  ShumokuTable(2) = "�w�j��"
  ShumokuTable(3) = "���j��"
  ShumokuTable(4) = "�o�^�t���C"
  ShumokuTable(5) = "�l���h���["
  ShumokuTable(6) = "�����["
  ShumokuTable(7) = "���h���[�����["
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
    Dim myQuery As String

    myQuery = "select �\���R�[�h, �\�� from �\�� where �\���R�[�h=(select max(�\���R�[�h) from �\��);"

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    
    If myRecordSet.recordCount > 0 Then
      
      ReDim Yoketsu(myRecordSet!�\���R�[�h)
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

      mySQL = "SELECT �\���R�[�h,�\�� FROM �\��"
      
      myRecordSet.Open mySQL, myCon, adOpenStatic, adLockOptimistic

      Yoketsu(0) = ""
      Do Until myRecordSet.EOF

        Yoketsu(myRecordSet!�\���R�[�h) = myRecordSet!�\��

        myRecordSet.MoveNext
      Loop
       
      myRecordSet.Close
      Set myRecordSet = Nothing
    
    myCon.Close
    Set myCon = Nothing

End Sub



Private Sub InitTeamDbArray(myCon As ADODB.Connection)
    Dim myRecordSet As New ADODB.Recordset
    Dim myQuery As String

    myQuery = "select �`�[���ԍ� from �����[�`�[�� where ���ԍ�= " & EventNo & _
      " and �`�[���ԍ�=(select max(�`�[���ԍ�) from �����[�`�[�� " & _
       " where ���ԍ�= " & EventNo & ");"

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    
    If myRecordSet.recordCount > 0 Then
      NumTeam = myRecordSet!�`�[���ԍ�
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
      mySQL = "SELECT �`�[���ԍ�,�`�[���� FROM �����[�`�[�� where ���ԍ�=" & EventNo
      
      myRecordSet.Open mySQL, myCon, adOpenStatic, adLockOptimistic

      TeamName4Relay(0) = ""
      Do Until myRecordSet.EOF

        TeamName4Relay(myRecordSet!�`�[���ԍ�) = myRecordSet!�`�[����

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
    Sheets("�I��o�^�ԍ����X�g").Select
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

    Cells(2, 1).Value = "�o�^�ԍ�"
    Cells(2, 2).Value = "�I�薼"
    Cells(2, 3).Value = "����"
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
    Dim myQuery As String
     
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
           
    myQuery = "SELECT ���Z�ԍ�, �I��ԍ�,  ��P�j��, ��Q�j��, ��R�j��, ��S�j��  " & _
              ", �g, ���H, ���R���̓X�e�[�^�X " & _
              "FROM �L�^ WHERE �g=" & kumi & " AND ���Z�ԍ�=" & uid & _
              " and ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    
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
   
    Dim myQuery As String
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
    Call frmMain.HideUnusedLane(minLaneNumber)
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
         
    myQuery = "SELECT ���Z�ԍ�, �I��ԍ�,  " & _
              "��P�j��, ��Q�j��, ��R�j��, ��S�j��  " & _
              ", �g, ���H, ���R���̓X�e�[�^�X " & _
              "FROM �L�^ WHERE �g=" & kumi & " AND ���Z�ԍ�=" & uid & _
              " and ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    frmMain.clearMe

    frmMain.lblClassName.Caption = className(ClassNumberbyUID(uid))
    frmMain.lblGender.Caption = GenderStr(GenderbyUID(uid))
    frmMain.lblRaceName.Caption = DistanceTable(DistancebyUID(uid)) + ShumokuTable(ShumokubyUID(uid))
    frmMain.lblPhase.Caption = Phase(uid)
    frmMain.lblClassName.Visible = True
    Do Until myRecordSet.EOF
      laneNo = CInt(myRecordSet!���H)


      laneNoStr = "lblName" & laneNo
      furiganaStr = "lblKana" & laneNo

      swimmer = Object2Int(myRecordSet!�I��ԍ�)
      If swimmer > 0 Then
        If laneNo > maxLaneNumber Then
          maxLaneNumber = laneNo
        End If
        If laneNo < minLaneNumber Then
          minLaneNumber = laneNo
        End If
        If is_relay(uid) Then
          frmMain.Controls(laneNoStr).Caption = TeamName4Relay(swimmer)
          If myRecordSet!���R���̓X�e�[�^�X = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "����"
          Else
            
            relaySwimmer(1) = Object2Int(myRecordSet!��1�j��)
            relaySwimmer(2) = Object2Int(myRecordSet!��2�j��)
            relaySwimmer(3) = Object2Int(myRecordSet!��3�j��)
            relaySwimmer(4) = Object2Int(myRecordSet!��4�j��)
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
          If myRecordSet!���R���̓X�e�[�^�X = 1 Then
            lane0NoStr = laneNoStr + "2"
            frmMain.Controls(lane0NoStr).Caption = "����"
          End If
        End If
      End If
      myRecordSet.MoveNext
    Loop
    Call CloseSQLConnection(myRecordSet, myCon)
    If CanGoWithNext(uid, kumi, maxLaneNumber) Then
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
   
    Dim myQuery As String
    Dim laneNo As Integer
    Dim laneNoDisp As Integer
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
         
    myQuery = "SELECT ���Z�ԍ�, �I��ԍ�,  " & _
              "��P�j��, ��Q�j��, ��R�j��, ��S�j��  " & _
              ", �g, ���H, ���R���̓X�e�[�^�X " & _
              "FROM �L�^ WHERE �g=" & kumi & " AND ���Z�ԍ�=" & uid & _
              " and ���ԍ�=" & EventNo
      
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    If frmMain.lblClassName.Caption <> "����" Then
      frmMain.lblClassName.Caption = "����"
      If GenderbyUID(prevUID) <> GenderbyUID(uid) Then
        frmMain.lblGender.Caption = ""
        frmMain.lblClassName1 = className(ClassNumberbyUID(prevUID)) & " " & GenderStr(GenderbyUID(prevUID))
      Else
        frmMain.lblClassName1.Caption = className(ClassNumberbyUID(prevUID))
      End If
      frmMain.lblClassName1.Visible = True
    End If
    Do Until myRecordSet.EOF
      laneNo = CInt(myRecordSet!���H)
      If laneNo = 10 Then
        laneNo = 0

      End If
      

      laneNoStr = "lblName" & laneNo
      furiganaStr = "lblKana" & laneNo

      swimmer = Object2Int(myRecordSet!�I��ԍ�)
      If swimmer > 0 Then
        If laneNo > maxLaneNumber Then
          maxLaneNumber = laneNo
        End If
        If laneNo <= minLaneNumber Then
          minLaneNumber = myRecordSet!���H
          If GenderbyUID(prevUID) <> GenderbyUID(uid) Then
            If minLaneNumber = 10 Then
                laneNo
            frmMain.Controls("lblClassName" & minLaneNumber).Caption = className(ClassNumberbyUID(uid)) & " " & _
                                                           GenderStr(GenderbyUID(uid))
          Else
            frmMain.Controls("lblClassName" & minLaneNumber).Caption = className(ClassNumberbyUID(uid))
          End If
          frmMain.Controls("lblClassName" & minLaneNumber).Visible = True
        End If
        If is_relay(uid) Then
          frmMain.Controls(laneNoStr).Caption = TeamName4Relay(swimmer)
          If myRecordSet!���R���̓X�e�[�^�X = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "����"
          Else
            
            relaySwimmer(1) = Object2Int(myRecordSet!��1�j��)
            relaySwimmer(2) = Object2Int(myRecordSet!��2�j��)
            relaySwimmer(3) = Object2Int(myRecordSet!��3�j��)
            relaySwimmer(4) = Object2Int(myRecordSet!��4�j��)
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
          If myRecordSet!���R���̓X�e�[�^�X = 1 Then
            lane0NoStr = laneNoStr + "1"
            frmMain.Controls(lane0NoStr).Caption = "����"
          End If
        End If
      End If
      myRecordSet.MoveNext
    Loop
    Call CloseSQLConnection(myRecordSet, myCon)
    If CanGoWithNext(uid, kumi, maxLaneNumber) Then
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


Function GetFirstOccupiedLane(uid As Integer) As Integer
    Dim myCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim myQuery As String
    Dim minLane As Integer
    Dim swimmer As Integer
    
    minLane = GetMaxLaneNo(uid)
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open
         
    myQuery = "SELECT  �I��ԍ�, �g, ���H " & _
              "FROM �L�^ WHERE ���ԍ�= " & EventNo & _
              " and �g=1 " + " AND ���Z�ԍ�=" & uid
    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic
    Do Until myRecordSet.EOF
      swimmer = Object2Int(myRecordSet!�I��ԍ�)
      If swimmer > 0 Then
        If CInt(myRecordSet!���H) < minLane Then
          minLane = CInt(myRecordSet!���H)
        End If
      End If
      myRecordSet.MoveNext
    Loop
    GetFirstOccupiedLane = minLane
    Call CloseSQLConnection(myRecordSet, myCon)
End Function
Function CanGoWithNext(uid As Integer, kumi As Integer, maxLaneNumber As Integer) As Boolean
    Dim prgNo As Integer
    Dim nextUid As Integer
    CanGoWithNext = False
    If maxLaneNumber = GetMaxLaneNo(uid) Then Exit Function
'    If kumi > 1 Then Exit Function
    
    prgNo = lastPrgNo + 1
    If prgNo > MaxPrgNo Then Exit Function
    
    nextUid = UIDFromRaceNo(prgNo)
'    If RaceExist(uid, 2) Then Exit Function
    If DistancebyUID(uid) <> DistancebyUID(nextUid) Then Exit Function
    If ShumokubyUID(uid) <> ShumokubyUID(nextUid) Then Exit Function
    If Phase(uid) <> Phase(nextUid) Then Exit Function
    If maxLaneNumber < GetFirstOccupiedLane(nextUid) Then
        CanGoWithNext = True
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
  Dim prgNo As Integer
  Dim group As Integer
  Dim uid As Integer

  Const loopLimit As Integer = 10
  Call frmMain.hide_class
  
  prgNo = lastPrgNo     ' frmMain.txtPrgNo.Value
  group = frmMain.txtKumi.Value

  Call InitAndReadDB
  uid = UIDFromRaceNo(prgNo)
  loopCount = 0
  Do While rc = False
    group = group + 1
    If RaceExist(uid, group) Then
      frmMain.txtPrgNo.Value = prgNo
      frmMain.txtKumi.Value = group
      firstPrgNo = prgNo
      lastPrgNo = prgNo
      Call show(uid, group)
      Exit Sub
    End If
    prgNo = prgNo + 1
    If prgNo > MaxPrgNo Then
      Call popup("�ŏI���[�X�ł��B")
      Exit Sub
    End If
    uid = UIDFromRaceNo(prgNo)
    group = 0
    loopCount = loopCount + 1
    If loopCount = loopLimit Then
      Call popup("�ŏI���[�X�ł��B(exceeds loop limit)")
      Exit Sub
    End If
  Loop

End Sub

Public Sub GetPrevRace()
  Dim rc As Boolean

  Dim prgNo As Integer
  Dim group As Integer
  Dim uid As Integer
  
  Call frmMain.hide_class
  prgNo = firstPrgNo    ' frmMain.txtPrgNo.Value
  group = frmMain.txtKumi.Value

  Call InitAndReadDB

  If (group > 1) Then
      group = group - 1
  Else
      prgNo = prgNo - 1
      If prgNo = 0 Then
        popup ("�ŏ��̃��[�X�ł��B")
        Exit Sub
      End If
      If prgNo > MaxPrgNo Then
        popup ("�Y�����郌�[�X�͂���܂���B�ŏI���[�X��\�����܂��B")
        prgNo = MaxPrgNo
      End If
      uid = UIDFromRaceNo(prgNo)
      rc = True
      Do While rc = True
        group = group + 1
        rc = RaceExist(uid, group)
      Loop
      group = group - 1
  End If
  uid = UIDFromRaceNo(prgNo)
  
  frmMain.txtPrgNo.Value = prgNo
  frmMain.txtKumi.Value = group
  lastPrgNo = prgNo
  firstPrgNo = prgNo
  Call show(uid, group)
      
End Sub

Public Sub GoWithNextRace()

  Dim loopCount As Integer
  Dim prgNo As Integer
  Dim group As Integer
  Dim uid As Integer
  Dim thisUID As Integer
  
  thisUID = UIDFromRaceNo(lastPrgNo)
    
  prgNo = lastPrgNo + 1
  lastPrgNo = prgNo
  group = 1
  Call InitAndReadDB
  uid = UIDFromRaceNo(prgNo)


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

  Dim prgNo As Integer
  Dim kumi As Integer
  Dim uid As Integer
  
  prgNo = frmMain.txtPrgNo.Value
  kumi = frmMain.txtKumi.Value
  Call frmMain.hide_class
  
  Call InitAndReadDB
  If prgNo > MaxPrgNo Then
    popup ("�Y���̃��[�X�͂���܂���B�ŏI���[�X��\�����܂��B")
    prgNo = MaxPrgNo
    frmMain.txtPrgNo.Value = prgNo
    frmMain.txtKumi.Value = 1
    kumi = 1
  End If
  uid = UIDFromRaceNo(prgNo)

  If Not RaceExist(uid, kumi) Then
    popup ("�Y���̃��[�X�͂���܂���B")
  Else
    lastPrgNo = prgNo
    firstPrgNo = prgNo
    Call show(uid, kumi)
  End If

End Sub

'-----------------------
Function GetGodoNo(myCon As ADODB.Connection, prgNo As Integer, kumiNo As Integer)
    Dim myQuery As String
    Dim myRecordSet As New ADODB.Recordset
    myQuery = "select �������[�X�ԍ� from �������[�X�v���O���� where ���ԍ�= " & EventNo & _
        " and (���Z�ԍ��P=" & prgNo & " or ���Z�ԍ�2=" & prgNo & " or ���Z�ԍ�3=" & prgNo & _
        " or   ���Z�ԍ�4 =" & prgNo & " or ���Z�ԍ�5=" & prgNo & " or ���Z�ԍ�6=" & prgNo & _
        " or   ���Z�ԍ�7 =" & prgNo & " or ���Z�ԍ�8=" & prgNo & " or ���Z�ԍ�9=" & prgNo & _
        " or   ���Z�ԍ�10 =" & prgNo & ") " & _
        " and ( �g1=" & kumiNo & " or �g2=" & kumiNo & " or �g3 =" & kumiNo & " or �g4 =" & kumiNo & _
        " or �g5 = " & kumiNo & " or �g6=" & kumiNo & " or �g7=" & kumiNo & " or �g8=" & kumiNo & _
        " or �g9 = " & kumiNo & " or �g10=" & kumiNo & " )"
        
     myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic, adLockReadOnly
     If myRecordSet.recordCount > 0 Then
        GetGodoNo = CInt(myRecordSet!�������[�X�ԍ�)
     Else
        GetGodoNo = 0
     End If
     myRecordSet.Close
     Set myRecordSet = Nothing
End Function
    
    
    
    
Sub CreateRaceArray()

    
    Dim myRecordSet As New ADODB.Recordset
    Dim myQuery As String
    Dim myCon As ADODB.Connection


    Dim recordCount As Integer
    Dim prgNo As Integer
    Dim kumiNo As Integer
    Set myCon = New ADODB.Connection
    myCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    myCon.Open


    
    myQuery = "SELECT distinct ���Z�ԍ�, �g FROM �L�^ where ���ԍ�= " & EventNo
    

    myRecordSet.Open myQuery, myCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        recordCount = recordCount + 1
        myRecordSet.MoveNext
    Loop

    myRecordSet.MoveFirst

    ReDim RealRace(3, recordCount)
    recordCount = 0
    Do Until myRecordSet.EOF
        prgNo = CInt(myRecordSet!���Z�ԍ�)
        kumiNo = CInt(myRecordSet!�g)
        RealRace(0, recordCount) = prgNo
        RealRace(1, recordCount) = kumiNo
        RealRace(2, recordCount) = GetGodoNo(myCon, prgNo, kumiNo)
        recordCount = recordCount + 1
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
    myCon.Close

End Sub
'---- for debugging use only
Sub DebugShow(recordCount As Integer)
    Dim i As Integer
    For i = 0 To recordCount
        Debug.Print "" & RealRace(0, i) & ", " & RealRace(1, i) & " , " & RealRace(2, i)
    Next i
End Sub
