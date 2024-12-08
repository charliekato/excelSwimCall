VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "レーン順"
   ClientHeight    =   12320
   ClientLeft      =   84
   ClientTop       =   396
   ClientWidth     =   15432
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdNextRace_Click()
    Call prgcall.GetNextRace
End Sub





Private Sub cmdNextTogether_Click()
    Call prgcall.GoWithNextRace
End Sub

Private Sub cmdPrevRace_Click()
    Call prgcall.GetPrevRace
End Sub

Private Sub cmdShow_Click()
    Call prgcall.ShowLaneOrder
End Sub

Private Sub cmdUpdate_Click()
    Call prgcall.reset
End Sub

Private Sub UserForm_Initialize()
    Me.txtPrgNo.Value = 1
    Me.txtKumi.Value = 1
    Call alignMe(80, 10, 140, 50)
    Call prgcall.ShowLaneOrder
    
End Sub

Public Sub clearMe()
    Dim i As Integer
    Dim lane As Integer
    For lane = 0 To 9
      Me.Controls("lblName" & lane).Caption = ""
      Me.Controls("lblKana" & lane).Caption = ""
          
      For i = 1 To 4
        Me.Controls("lblName" & lane & i).Caption = ""
        Me.Controls("lblKana" & lane & i).Caption = ""
      Next i
    Next lane
    
End Sub

Public Sub HideUnusedLane(maxLane As Integer)
    Dim lane As Integer
    For lane = 1 To 9
        If (lane > maxLane) Then

            Me.Controls("lbl" & lane & "lane").Visible = False
        Else

            Me.Controls("lbl" & lane & "lane").Visible = True
        End If
    Next lane
    If maxLane = 10 Then
        Me.Controls("lbl0lane").Visible = True
    Else
        Me.Controls("lbl0lane").Visible = False
    End If
    
End Sub



Public Sub hide_class()
    Dim lane As Integer
    For lane = 0 To 9
        Me.Controls("lblClassName" & lane).Visible = False
    Next lane
End Sub
Public Sub alignMe(mostTop As Integer, mostLeft As Integer, swimmerNameWidth As Integer, heightInterval)
    Dim lane As Integer
    Dim order As Integer

    Dim top4ClassName As Integer
    Dim top4laneNo As Integer
    Dim top4Kana As Integer
    Dim laneLeft As Integer
    Dim nameLeft As Integer

    
    laneLeft = mostLeft + 10
    For lane = 0 To 9
      top4ClassName = mostTop + (heightInterval * lane)
      top4Kana = top4ClassName + 15
      top4laneNo = top4ClassName + 25

      nameLeft = laneLeft + 25
      Me.Controls("lbl" & lane & "Lane").Left = laneLeft
      Me.Controls("lbl" & lane & "Lane").Height = 20
      Me.Controls("lblClassName" & lane).Top = top4ClassName
      Me.Controls("lblClassName" & lane).Height = 15
      Me.Controls("lblClassName" & lane).Left = mostLeft
      Me.Controls("lblClassName" & lane).Visible = False
      Me.Controls("lbl" & lane & "Lane").Top = top4laneNo
      Me.Controls("lblName" & lane).Left = nameLeft
      Me.Controls("lblName" & lane).Height = 20
      Me.Controls("lblName" & lane).Top = top4laneNo
      Me.Controls("lblName" & lane).width = swimmerNameWidth
      Me.Controls("lblKana" & lane).Left = nameLeft
      Me.Controls("lblKana" & lane).width = swimmerNameWidth
      Me.Controls("lblKana" & lane).Top = top4Kana
      For order = 1 To 4
        nameLeft = nameLeft + swimmerNameWidth
        Me.Controls("lblName" & lane & order).Left = nameLeft
        Me.Controls("lblName" & lane & order).Height = 20
        Me.Controls("lblName" & lane & order).width = swimmerNameWidth
        Me.Controls("lblName" & lane & order).Top = top4laneNo
        Me.Controls("lblKana" & lane & order).Left = nameLeft
        Me.Controls("lblKana" & lane & order).width = swimmerNameWidth
        Me.Controls("lblKana" & lane & order).Top = top4Kana
      Next order
    Next lane
    
End Sub
