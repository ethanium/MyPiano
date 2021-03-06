Attribute VB_Name = "MidiBeat"
Public Type BeatData
    Name As String
    Rate As Integer
    Data As String
End Type

Public Beats(2) As BeatData

Public Sub LoadBeat()
    Beats(1).Name = "Rock"
    Beats(1).Rate = 150
    Beats(1).Data = "44,35,55||44||44,40||44|;" _
        & "44,35||44,35||44,40||44||44,35||44||44,40||44|;" _
        & "44,50,35|50|44,48,35|48|44,47,35|47|44,45,35|45|44,35,55||44||44,40||44|;" _
        & "44,35||46||40,44||42|;" _
        & "44||44||44||44,35,55"
    
    Beats(2).Name = "16 Beat"
    Beats(2).Rate = 100
    Beats(2).Data = "44|44|44|44|44|44|44|44;" _
        & "44,35|44|44|44|44,40|44|44|44;" _
        & "44,40|44,40|44,40|44,40|44,40|44,40|44,40|44,40|44,35,55|44|44|44|44,40|44|44|44;" _
        & "44,35|44|44|46|40|42|44|44;" _
        & "44,35,55"
        
    Beats(0).Name = "Disco"
    Beats(0).Rate = 170
    Beats(0).Data = "44,35,55|46|44,40,35|42;" _
        & "44,35|46|44,40,35|42;" _
        & "44,35|44,35,40|44,35,40|44,35,40|44,35,55|46|44,40,35|42;" _
        & "44,35|44,40|44,40|44,35;" _
        & "44,35,55"
        
    
'    Beats(0).Name = "Test"
'    Beats(0).Data = Array( _
'        Array(), _
'        Array(), _
'        Array(), _
'        Array(), _
'        Array() _
'    )
End Sub


