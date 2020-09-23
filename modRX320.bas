Attribute VB_Name = "modRX320"
    '*********************************************
    '
    '   Module to control Ten Tec RX320 Reciever
    '
    '   Mark Mokoski
    '   17-FEB-2005
    '
    '*********************************************

    Option Explicit

    'Public variables
    Public volSPK                    As Integer 'Speaker volume, 63-0
    Public volLINE                   As Integer 'Line level, 63-0
    Public valFilter                 As Integer 'IF filter width in Hz
    Public valAGC                    As String  'AGC decay, fast="3", med="2", slow="1"
    Public valMode                   As String  'RCV mode, AM,USB,LSB,CW
    Public valModeCor                As Integer 'RCV mode correction, AM=0,USB=1,LSB=-1,CW=-1
    Public valFreq                   As Long    'Displayed RCV frequency in Hz
    Public valAdjFreq                As Integer 'Adjustment frequency factor for radio in Hz
    Public valCWOffSet               As Integer 'Offset for LO ic CW mode in Hz
    Public valStep                   As Long    'Tuning frequency step value in Hz
    Public valSignalLevel            As Integer 'S Meter signal level, 0 - 10,000
    Public valMUTE                   As Boolean 'RCV muted true/false
    Public valBFO                    As Integer 'Tuning factor BFO correction
    'Public LevelList(5)              As Integer 'List of the 5 last "S" meter values
    Public strFreq                   As String  'String used to input nre frequency
    Public RadioComPort              As Integer 'COM Port number for radio communication

    
Public Function SetSMeter(MeterString As String)

    '
    '   Get "S" meter value and display
    '   Set "S" Meter on main form
    '

    On Error Resume Next

    Dim SigLevel            As Integer
    Dim s                   As Integer
    Dim Segments            As Integer

    'Static LevelList(5)            As Integer 'List of the 5 last "S" meter values
    
    SigLevel = (Asc(Mid(MeterString, 2, 1)) * 256) + Asc(Mid(MeterString, 3, 1))
    
    'CC_Comment Out (4/22/2005):
    '    'Shift old "S" mete values

    '        For s = 0 To 3
    '            LevelList(s) = LevelList(s + 1)
    '        Next s
    
    '    LevelList(4) = SigLevel
        
    '    'Average the last 5 "S" meter readings
    '    SigLevel = (LevelList(0) + LevelList(1) + LevelList(2) + LevelList(3) + LevelList(4)) / 5
        
    'End CC_Comment Out
    Segments = SigLevel / 40
    Segments = (Segments / 21) - 1

        If Segments > 20 Then Segments = 20
        
        For s = 0 To 20
            frmRX320.sUnit(s).Visible = False
        Next s

        For s = 0 To Segments
            frmRX320.sUnit(s).Visible = True
        Next s


End Function

Public Function SetAGC(strAGC As String)

    '
    '   Set the AGC decay time constant
    '

    strAGC = UCase(strAGC)

        Select Case strAGC
            Case "FAST"
                valAGC = "3"
            Case "MEDIUM"
                valAGC = "2"
            Case "SLOW"
                valAGC = "1"
        End Select
    
    frmRX320.RadioCOM.Output = "G" & valAGC & vbCr
    
End Function

Public Function SetFilter(intFilter As Integer)

    '
    '   Set the IF filter width in Hz
    '

        Select Case intFilter
            Case 8000
                valFilter = intFilter
                intFilter = 33
            Case 6000
                valFilter = intFilter
                intFilter = 0
            Case 5700
                valFilter = intFilter
                intFilter = 1
            Case 5400
                valFilter = intFilter
                intFilter = 2
            Case 5100
                valFilter = intFilter
                intFilter = 3
            Case 4800
                valFilter = intFilter
                intFilter = 4
            Case 4500
                valFilter = intFilter
                intFilter = 5
            Case 4200
                valFilter = intFilter
                intFilter = 6
            Case 3900
                valFilter = intFilter
                intFilter = 7
            Case 3600
                valFilter = intFilter
                intFilter = 8
            Case 3300
                valFilter = intFilter
                intFilter = 9
            Case 3000
                valFilter = intFilter
                intFilter = 10
            Case 2850
                valFilter = intFilter
                intFilter = 11
            Case 2700
                valFilter = intFilter
                intFilter = 12
            Case 2550
                valFilter = intFilter
                intFilter = 13
            Case 2400
                valFilter = intFilter
                intFilter = 14
            Case 2250
                valFilter = intFilter
                intFilter = 15
            Case 2100
                valFilter = intFilter
                intFilter = 16
            Case 1950
                valFilter = intFilter
                intFilter = 17
            Case 1800
                valFilter = intFilter
                intFilter = 18
            Case 1650
                valFilter = intFilter
                intFilter = 19
            Case 1500
                valFilter = intFilter
                intFilter = 20
            Case 1350
                valFilter = intFilter
                intFilter = 21
            Case 1200
                valFilter = intFilter
                intFilter = 22
            Case 1050
                valFilter = intFilter
                intFilter = 23
            Case 900
                valFilter = intFilter
                intFilter = 24
            Case 750
                valFilter = intFilter
                intFilter = 25
            Case 675
                valFilter = intFilter
                intFilter = 26
            Case 600
                valFilter = intFilter
                intFilter = 27
            Case 525
                valFilter = intFilter
                intFilter = 28
            Case 450
                valFilter = intFilter
                intFilter = 29
            Case 375
                valFilter = intFilter
                intFilter = 30
            Case 330
                valFilter = intFilter
                intFilter = 31
            Case 300
                valFilter = intFilter
                intFilter = 32
        End Select
    
    frmRX320.RadioCOM.Output = "W" & Chr$(intFilter) & vbCr
    SetFREQ (valFreq)

End Function

Public Function SetStep(intStep As Integer)

    '
    '   Set the frequency increment (Step) in Hz
    '

    valStep = intStep

End Function

Public Function SetSPKvol(valSPK As Integer)

    '
    '   Set the speaker volume 63=min, 0=max
    '


        If frmRX320.RadioCOM.PortOpen = True Then
            frmRX320.RadioCOM.Output = "V" + Chr$(0) + Chr$(valSPK) + vbCr
        End If

End Function

Public Function SetLINEvol(valLINE As Integer)

    '
    '   Set the Line Out level 63=min, 0=max

        If frmRX320.RadioCOM.PortOpen = True Then
            frmRX320.RadioCOM.Output = "A" + Chr$(0) + Chr$(valLINE) + vbCr
        End If

End Function

Public Function SetMode(strMode As String)

    '
    '   Set the receive mode and mode correction value
    '

        Select Case strMode
            Case "AM"
                valMode = "AM"
                strMode = "0"
                valModeCor = 0
                valBFO = 0
            Case "USB"
                valMode = "USB"
                strMode = "1"
                valModeCor = 1
                valBFO = 0
            Case "LSB"
                valMode = "LSB"
                strMode = "2"
                valModeCor = -1
                valBFO = 0
            Case "CW"
                valMode = "CW"
                strMode = "3"
                valModeCor = -1
                valBFO = valCWOffSet
        End Select

    frmRX320.RadioCOM.Output = "M" & strMode & vbCr
    SetFREQ (valFreq)

End Function

Public Function SetFREQ(newFREQ As Long)

    '
    '   Set the receive frequency
    '   This is also called after the following setting changes:
    '   MODE change
    '   FILTER change
    '   CW OFFSET change
    '   FREQUNCY ADJUST change
    '   And also changeing the RECEIVE FREQUNCY it self
    '
    '   Tuning Factors, from TenTec documentation.
    '   Converted to VB6 code from GW BASIC code in TenTec RX320 Programmers Reference.
    '   I followed the TenTec variable names so that it would be
    '   easy to follow the flow of the code and read the TenTec doc's
    '   so code can be compaired side by side.
    '
    '   Tfreq = Tuned Frequency in MHz.
    '   Mcor = Mode Correction = 0 for AM mode, +1 for USB, -1 for LSB, -1 for CW.
    '   Fcor = Filter Correction calculated using (Bandwidth/2)+200 where bandwidth is in Hz.
    '   Cbfo = Desired center frequency of filter in Hz for CW mode. Only needed in CW mode.
    '
    '   SETTING THE RADIO TUNING FACTORS
    '   calculation:
    '   AdjTfreq = Adjusted Tuned Frequency = Tfreq-0.00125+(Mcor * (Fcor+Cbfo))/1000000
    '   Ctf = Coarse tuning factor = (int) (AdjTfreq)/0.0025)+18000
    '   where (int) is the integer function to get the integer only portion of the division
    '   Ftf = Fine tuning factor = mod(AdjTfreq/0.0025) * 2500 * 5.46
    '   where mod is the modulus function used to get the fractional remainder of a division operation.
    '   Btf = Bfo Tuning Factor = (int) ((fcor+CWBFO+8000) * 2.73)
    '   where (int) is the integer function to get the integer only protion of the division
    '
    '
    'Declare local variables

    Dim Tfreq                      As Single
    Dim Mcor                       As Integer
    Dim Fcor                       As Integer
    Dim Cbfo                       As Integer
    Dim AdjTfreq                   As Single
    Dim Ctf                        As Long
    Dim Ftf                        As Long
    Dim Btf                        As Long
    
    
    '"High" and "Low" bytes of tuning factors
    Dim Ch                         As Long
    Dim Cl                         As Long
    Dim Fh                         As Long
    Dim Fl                         As Long
    Dim Bh                         As Long
    Dim Bl                         As Long

    'Setup the variables for calculation
    Tfreq = (valFreq + valAdjFreq) / 1000000
    Mcor = valModeCor
    Fcor = (valFilter / 2) + 200
    Cbfo = valBFO
    
    'Calculate the tuning factors
    AdjTfreq = Tfreq - 0.00125 + (Mcor * (Fcor + Cbfo)) / 1000000
    Ctf = Int(AdjTfreq / 0.0025)
    Ftf = (((AdjTfreq / 0.0025) - Ctf) * 2500 * 5.46)
    Btf = Int((Fcor + Cbfo + 8000) * 2.73)
    Ctf = Ctf + 18000
    
    'Now break up the factors into Bytes for sending to COM Port
    Ch = Int(Ctf / 256)
    Cl = Ctf - (Ch * 256)
    Fh = Int(Ftf / 256)
    Fl = Ftf - (Fh * 256)
    Bh = Int(Btf / 256)
    Bl = Btf - (Bh * 256)
    
    'Now send it to the radio
    frmRX320.RadioCOM.Output = "N" & Chr$(Ch) & Chr$(Cl) & Chr$(Fh) & Chr$(Fl) & Chr$(Bh) & Chr$(Bl) & vbCr
    
    
End Function

Public Function SetCWOFFSET(newOFFSET As Integer)

    '
    '   Set CW Mode offset in Hz
    '

    valCWOffSet = newOFFSET
    SetFREQ (valFreq)

End Function

Public Function SetFreqAdjust(newFreqAdjust As Integer)

    '
    '   Set value of frequncy adjustment (LO deviation) in Hz
    '

    valAdjFreq = newFreqAdjust
    SetFREQ (valFreq)

End Function

Public Function SetMUTE(newMUTE As Boolean)

    '
    '   Set value of the Mute flag
    '
    valMUTE = newMUTE

End Function
