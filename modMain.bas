Attribute VB_Name = "modMain"
    '*********************************************
    '
    '   Module to control Ten Tec RX320 Reciever
    '
    '   Mark Mokoski
    '   17-FEB-2005
    '
    '*********************************************
    
    'General settings
    Public DSPver                       As String  'DSP firmware version
    Public muteOnExit                   As Boolean 'Flag for muting the reciever on program exit, prompt user if false
    Public show1Hz                      As Boolean 'Show or not the 1 Hz digit and step selection
    Public CurrentVFO                   As Integer 'What is the "Main" VFO 0 = A, 1 = B
    Public FixLineLevel                 As Boolean 'Set Line out as fixed level(valFixedLevel)
    Public valFixedLevel                As Integer 'Value of fixed line out level ( 0 = max, 63 = min)
    Public WindowSize                   As Integer '"Full Window" = 1, "Basic Window" = 0
    Public RadioOK                      As Boolean 'Flag set if radio responds to a DSPver query or "Power ON" string received

    'VFO settings
    Public freqVFO(2)                   As Long    'Stored freqency for VFO
    Public VFOFilter(2)                 As Integer 'IF filter width in Hz
    Public VFOAGC(2)                    As String  'AGC decay, fast="3", med="2", slow="1"
    Public VFOMode(2)                   As String  'RCV mode, AM,USB,LSB,CW
    Public VFOModeCor(2)                As Integer 'RCV mode correction, AM=0,USB=1,LSB=-1,CW=-1
    Public VFOFreq(2)                   As Long    'Displayed RCV frequency in Hz
    Public VFOAdjFreq(2)                As Integer 'Adjustment frequency factor for radio in Hz
    Public VFOCWOffSet(2)               As Integer 'Offset for LO ic CW mode in Hz
    Public VFOStep(2)                   As Long    'Tuning frequency step VFO in Hz
    Public VFOBFO(2)                    As Integer 'Tuning factor BFO correction
    'Individual Mode Parameters
    Public AMfilter                     As Integer 'Last selected AMfilter
    Public AMquickfilter(6)             As Integer 'AM Quick filter selections
    Public AMstep                       As String  'AM tuning step
    Public SSBfilter                    As Integer 'Last selected SSB filter
    Public SSBquickfilter(6)            As Integer 'SSB Quick filter selections
    Public SSBstep                      As String  'SSB tuning step
    Public LSBoffset                    As Integer 'LSB/CW tuning correction
    Public USBoffset                    As Integer 'USB tuning correction
    Public CWfilter                     As Integer 'Last selected CW filter
    Public CWquickfilter(6)             As Integer 'CW Quick filter selection
    Public CWstep                       As String  'CW tuning step

Sub Main()

    ' ***************************************************************************
    ' * Test to see if App is allready running
    ' * If App is running, terminate copy
    ' ***************************************************************************

        If App.PrevInstance Then
            MsgBox App.Title & " application is already running." & vbCrLf & _
            "Only one instance (copy) of program this can be running" & vbCrLf & _
            "for proper operation.", vbCritical, "Application ERROR"
            End
        Else
            '  MsgBox "This is the first instance of your application."
        End If

    '
    '   Load saved/default values from registry
    '

    'frmRX320.Visible = True
    Load frmRX320
    
End Sub

Public Sub SaveRegSettings()


    'Save current radio settings to registry
    'General settings
    SaveSetting "RX320", "General", "RadioComPort", RadioComPort
    SaveSetting "RX320", "General", "volSPK", volSPK
    SaveSetting "RX320", "General", "volLINE", volLINE
    SaveSetting "RX320", "General", "valFilter", valFilter
    SaveSetting "RX320", "General", "valAGC", valAGC
    SaveSetting "RX320", "General", "valMode", valMode
    SaveSetting "RX320", "General", "valModeCor", valModeCor
    SaveSetting "RX320", "General", "valFreq", valFreq
    SaveSetting "RX320", "General", "valAdjFreq", valAdjFreq
    SaveSetting "RX320", "General", "valCWOffSet", valCWOffSet
    SaveSetting "RX320", "General", "valStep", valStep
    SaveSetting "RX320", "General", "valBFO", valBFO
    SaveSetting "RX320", "General", "muteOnExit", muteOnExit
    SaveSetting "RX320", "General", "show1Hz", show1Hz
    SaveSetting "RX320", "General", "FixLineLevel", FixLineLevel
    SaveSetting "RX320", "General", "valFixedLevel", valFixedLevel
    SaveSetting "RX320", "General", "WindowSize", WindowSize
    DoEvents
     
    'Mode settings
    SaveSetting "RX320", "AM", "AMfilter", AMfilter         'Last selected AMfilter
    SaveSetting "RX320", "AM", "AMstep", AMstep             'AM tuning step
    SaveSetting "RX320", "SSB", "SSBfilter", SSBfilter      'Last selected SSB filter
    SaveSetting "RX320", "SSB", "SSBstep", SSBstep          'SSB tuning step
    SaveSetting "RX320", "SSB", "LSBoffset", LSBoffset      'LSB/CW tuning correction
    SaveSetting "RX320", "SSB", "USBoffset", USBoffset      'USB tuning correction
    SaveSetting "RX320", "CW", "CWfilter", CWfilter         'Last selected CW filter
    SaveSetting "RX320", "CW", "CWstep", CWstep             'CW tuning step

    Dim x            As Integer

        For x = 0 To 5
            SaveSetting "RX320", "AM", "AMquickfilter_" & x, AMquickfilter(x)       'AM Quick filter selections
            SaveSetting "RX320", "SSB", "SSBquickfilter_" & x, SSBquickfilter(x)    'SSB Quick filter selections
            SaveSetting "RX320", "CW", "CWquickfilter_" & x, CWquickfilter(x)       'CW Quick filter selection
        Next x

    DoEvents
    
    'VFO settings
    SaveSetting "RX320", "General", "CurrentVFO", CurrentVFO

    'VFO A
    SaveSetting "RX320", "VFOA", "freqVFO", freqVFO(0)
    SaveSetting "RX320", "VFOA", "VFOFilter", VFOFilter(0)
    SaveSetting "RX320", "VFOA", "VFOAGC", VFOAGC(0)
    SaveSetting "RX320", "VFOA", "VFOMode", VFOMode(0)
    SaveSetting "RX320", "VFOA", "VFOModeCor", VFOModeCor(0)
    SaveSetting "RX320", "VFOA", "VFOFreq", VFOFreq(0)
    SaveSetting "RX320", "VFOA", "VFOAdjFreq", VFOAdjFreq(0)
    SaveSetting "RX320", "VFOA", "VFOCWOffSet", VFOCWOffSet(0)
    SaveSetting "RX320", "VFOA", "VFOStep", VFOStep(0)
    SaveSetting "RX320", "VFOA", "VFOBFO", VFOBFO(0)
    
    'VFO B
    SaveSetting "RX320", "VFOB", "freqVFO", freqVFO(1)
    SaveSetting "RX320", "VFOB", "VFOFilter", VFOFilter(1)
    SaveSetting "RX320", "VFOB", "VFOAGC", VFOAGC(1)
    SaveSetting "RX320", "VFOB", "VFOMode", VFOMode(1)
    SaveSetting "RX320", "VFOB", "VFOModeCor", VFOModeCor(1)
    SaveSetting "RX320", "VFOB", "VFOFreq", VFOFreq(1)
    SaveSetting "RX320", "VFOB", "VFOAdjFreq", VFOAdjFreq(1)
    SaveSetting "RX320", "VFOB", "VFOCWOffSet", VFOCWOffSet(1)
    SaveSetting "RX320", "VFOB", "VFOStep", VFOStep(1)
    SaveSetting "RX320", "VFOB", "VFOBFO", VFOBFO(1)

    DoEvents
 
End Sub
