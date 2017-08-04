Attribute VB_Name = "Ca310Lib"
Option Explicit

'===================================
' Application Data Type Definiition
'===================================
Public Type TypeColor
    color0 As Single
    color1 As Single
    color2 As Single
End Type

Type RefData
    Refx As Single
    Refy As Single
    RefLv As Single
    Mesx As Single
    Mesy As Single
    MesLv As Single
    MesFlicker As Single
    SelectDataName As String '"COLOR","FMA","JEITA"
End Type

Public Type TypeMeasurementData

    dateColorData As Date
    timeColorData As Date
    lColorStatus As Long
    ssx As Single
    ssy As Single
    sLv As Single
    sLvfL As Single
    Sx As Single
    Sy As Single
    Sz As Single
    sud As Single
    svd As Single
    sduv As Single
    LT As Long
    
    susUser As Single
    svsUser As Single
    sLsUser As Single
    sdEUser As Single
    
    dateFMAData As Date
    timeFMAData As Date
    lFMAStatus As Long
    sFMA As Single
    
    dateJEITAData As Date
    timeJEITAData As Date
    lJEITAStatus As Long
    sJEITA As Single

End Type

Public Type TypeReferenceData
    sRefx As Single
    sRefy As Single
    sRefLv As Single
End Type

'===================================
' SDK Object
'===================================
Public objCa200 As Ca200
Public objCa As Ca
Public objProbe As Probe
Public objMemory As Memory

'===================================
' SDK Constant
'===================================
'---------------
' CA Display Mode
'---------------
Public Const DSP_LXY As Long = 0
Public Const DSP_DUV As Long = 1
Public Const DSP_ANL As Long = 2
Public Const DSP_ANLG As Long = 3
Public Const DSP_ANLR As Long = 4
Public Const DSP_PUV As Long = 5
Public Const DSP_FMA As Long = 6
Public Const DSP_XYZ As Long = 7
Public Const DSP_JEITA As Long = 8

'---------------
' CA Sync. Mode
'---------------
Public Const SYNC_NTSC As Long = 0
Public Const SYNC_PAL As Long = 1
Public Const SYNC_EXT As Long = 2
Public Const SYNC_UNIV As Long = 3
Public Const SYNC_INT As Long = 4

'---------------
' CA Display Digits Mode
'---------------
Public Const DIGT_3 As Long = 0
Public Const DIGT_4 As Long = 1

'---------------
' CA Fas/Slow Mode
'---------------
Public Const AVRG_SLOW As Long = 0
Public Const AVRG_FAST As Long = 1
Public Const AVRG_AUTO As Long = 2

'---------------
' CA Brightness Unit Option
'---------------
Public Const BUNIT_FL As Long = 0
Public Const BUNIT_CD As Long = 1

'---------------
' CA Calibration Mode
'---------------
Public Const CAL_D65 As Long = 1
Public Const CAL_9300 As Long = 2
Public Const CAL_WHITE As Long = 10
Public Const CAL_MATRIX As Long = 11


'===================================
' Application Data
'===================================
Public typCurrentMeasurementData As TypeMeasurementData
Public typCurrentRefereceData As TypeReferenceData
Public gstrCADataFile As String
Public gstrVGDataFile As String

'===================================
' Application Constant
'===================================
Public Const STD_D65 As String = "Minolta D65"
Public Const STD_9300K As String = "Minolta 93K"
Public Const STD_WHITE As String = "User White"
Public Const STD_MATRIX As String = "User Matrix"
Public FormVisibleFlg As Boolean

Public CA_Type As String
Public gstrLvOrEv As String

Public Const FORMAT_SXY As String = "0.0000"
Public Const FORMAT_LV As String = "0.0###"



Public Sub StartMain()
    Dim i As Integer
    
    '===================================
    ' Set Error Trap
    '===================================
    On Error GoTo E
    
    '===================================
    ' Create SDK/Application Object
    '===================================
    Set objCa200 = New Ca200
    'DoEvents
    '===================================
    ' Set Configuration
    '===================================
    objCa200.AutoConnect
    DoEvents
    '===================================
    ' Initialize SDK Object
    '===================================
    Set objCa = objCa200.SingleCa
    Set objProbe = objCa.SingleProbe
    Set objMemory = objCa.Memory
    CA_Type = objCa.CAType
    '"CA-210"
    If CA_Type = "CA-210" Or CA_Type = "CA-310" Then
    End If
    
    '===================================
    ' Initialize CA and Application
    '===================================
    
    ' 021224
    gstrLvOrEv = "Lv"
    
    ' 0 Calibration
    MsgBox "0-Cal,CA310校准", vbOKOnly
    FormGamma.Label3.Caption = "CA310校准中，请静候校准完成..."
    DoEvents
    objCa.CalZero
    DoEvents
    Screen.MousePointer = vbHourglass
    
    objMemory.GetReferenceColor objProbe.ID, typCurrentRefereceData.sRefx, typCurrentRefereceData.sRefy, typCurrentRefereceData.sRefLv
    '=================================================================
'    Load FormGamma
    FormGamma.Show
    
    FormGamma.xForm_Initialize
        
    FormGamma.Enabled = False
    '==============================================================
    FormGamma.FrameRefData.Caption = "Ref. xyLv"
    
    DoEvents
    Set FormGamma.objCaControl.Ca = objCa
    Set FormGamma.objCaControl.Probe = objProbe
    Set FormGamma.objCaControl.Memory = objMemory
    Set FormGamma.xyControl1.Probe = objProbe
    FormGamma.objCaControl.UpdateCaInfo
    FormGamma.objCaControl.UpdateMemoryInfo
    
    '===========================================
    ' Set CA Display Mode and Measuring Pattern
    '===========================================
    If CA_Type = "CA-210" Or CA_Type = "CA-310" Then
        If objCa.DisplayMode = DSP_FMA Or objCa.DisplayMode = DSP_JEITA Then
            objCa.DisplayMode = DSP_LXY
        End If
    End If
    'FormGamma.objVGControl.Pattern = 1
    'FormGamma.objVGControl.SetGVideoLevel 255, 255, 255
   ' FormGamma.objVGControl.RedSW = True
    'FormGamma.objVGControl.GreenSW = True
   ' FormGamma.objVGControl.BlueSW = True
    
    '===================================
    ' Show Main Form
    '===================================
    FormGamma.Tag = "END"
    Screen.MousePointer = vbDefault
    FormGamma.Enabled = True
    'FormGamma.Show
    
    Exit Sub
    
E:
    '===================================
    ' Error Trap
    '===================================
    Dim strERR As String
    Dim iReturn As Integer
    
    strERR = "Error from " + Err.Source + Chr$(10) + Chr$(13)
    strERR = strERR + Err.Description + Chr$(10) + Chr$(13)
    strERR = strERR + "HRESULT " + CStr(Err.Number - vbObjectError)
    iReturn = MsgBox(strERR, vbRetryCancel)
    Select Case iReturn
        Case vbRetry: Resume
        Case Else:
            'objCa.RemoteMode = 0
            End
    End Select
End Sub
