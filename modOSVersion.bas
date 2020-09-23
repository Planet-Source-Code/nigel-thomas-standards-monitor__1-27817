Attribute VB_Name = "modOSVersion"
Option Explicit

'**********************************************************
'* Based on
'**********************************************************

Public Enum OSTypes
    osUnknown
    osWindows3x
    osWindows95
    osWindows95OSR2
    osWindows98
    osWindows98SE
    osWindowsME
    oswindowsNT3
    osWindowsNT31
    osWindowsNT35
    osWindowsNT351
    osWindowsNT4
    osWindows2000Professional
    osWindows2000Server
    osWindows2000AdvancedServer
    osWindows2000DataCenter
    osWindowsXPHomeEdition
    osWindowsXPProfessional
    osWindowsDOTNETEnterpriseServer
    osWindowsDOTNETServer
End Enum

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Function GetOSType() As OSTypes
   
   Dim osinfo As OSVERSIONINFOEX
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = Len(osinfo)

   With osinfo
   
   Select Case .dwPlatformId
      
      Case 0
        
        GetOSType = osWindows3x
        
      Case 1
      
         If .dwMinorVersion = 0 Then
            
            If InStr(UCase$(osinfo.szCSDVersion), "C") Then
                
                GetOSType = osWindows95OSR2
            Else
                
                GetOSType = osWindows95
            End If
         
         ElseIf .dwMinorVersion = 10 Then
                        
            If InStr(UCase(osinfo.szCSDVersion), "A") Then
                
                GetOSType = osWindows98SE
            Else
                
                GetOSType = osWindows98
            End If
         
         ElseIf .dwMinorVersion = 90 Then
         
            GetOSType = osWindowsME
         End If
      
      Case 2
         
         If .dwMajorVersion = 3 Then
            
            Select Case .dwMinorVersion
                
                Case 0:  GetOSType = oswindowsNT3
                Case 1:  GetOSType = osWindowsNT31
                Case 2:  GetOSType = osWindowsNT35
                Case 51: GetOSType = osWindowsNT351
            End Select
            
         ElseIf .dwMajorVersion = 4 Then
            
            GetOSType = osWindowsNT4
            
         ElseIf .dwMajorVersion = 5 Then
            
            Select Case .wProductType
            
                Case 1:  GetOSType = osWindows2000Professional
                Case 3:
                    
                    Select Case .wSuiteMask
                    
                        Case 128: GetOSType = osWindows2000DataCenter
                        Case 2:   GetOSType = osWindows2000AdvancedServer
                        Case Else: GetOSType = osWindows2000Server
                        
                    End Select
                    
            End Select
            
         ElseIf .dwMajorVersion = 1 Then
            
            Select Case .wProductType
            
                Case 1 'win XP
    
                    If .wSuiteMask = 512 Then
                    
                        GetOSType = osWindowsXPHomeEdition
                    Else
                        
                        GetOSType = osWindowsXPProfessional
                    End If
                
                Case Else
                
                    If .wSuiteMask = 2 Then
                        
                        GetOSType = osWindowsDOTNETEnterpriseServer
                    Else
                        
                        GetOSType = osWindowsDOTNETServer
                    End If
            End Select
            
            End If
      
      
      
      Case Else
      
         GetOSType = osUnknown
   End Select
   
   End With

End Function


