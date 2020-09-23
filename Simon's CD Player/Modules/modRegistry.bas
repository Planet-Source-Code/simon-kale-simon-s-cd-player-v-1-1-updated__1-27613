Attribute VB_Name = "modRegistry"
Option Explicit
'get the return messages from the Registry APIs
Dim lngReturnValues As Long
'the variable that receives the handle of the opened or created key
Dim lngResult As Long
'Points to a variable that receives one of the following disposition values:
'REG_CREATED_NEW_KEY The key did not exist and was created.
'REG_OPENED_EXISTING_KEY The key existed and was simply opened without being changed.
Dim lngDisposition As Long
'handle of the currently opened registry key
Dim hKey As Long
'path of the Registry key to be opened
Dim strKeyPath As String


Public Sub CreateRegistrySettings()
    '* Purpose: Create the program's settings key on the registry
    
    strKeyPath = "Software\SimonSoft\Simon's CD Player\Settings"
    'creates the path 'HKEY_CURRENT_USER\Software\SimonSoft\Simon's CD Player\Settings'
    lngReturnValues = RegCreateKeyEx(HKEY_CURRENT_USER, strKeyPath, 0, _
                            "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                            ByVal 0&, lngResult, lngDisposition)
    
    strKeyPath = "Software\SimonSoft\Simon's CD Player\Dimension"
    'creates the path 'Software\SimonSoft\Simon's CD Player\Dimension'
    lngReturnValues = RegCreateKeyEx(HKEY_CURRENT_USER, strKeyPath, 0, _
                            "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                            ByVal 0&, lngResult, lngDisposition)
    
        'opens the path "Software\SimonSoft\Simon's CD Player\Settings"
        If OpenRegistrySettings = True Then
            'create these keys on the registry
            Dim strStopCDValue As String
            'set "strStopCDValue" to False
            'stores data in the value field of
            ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
            lngReturnValues = RegSetValueEx(hKey, "StopPlayOnExit", 0, REG_SZ, _
                                            ByVal strStopCDValue, Len(strStopCDValue))
            
            Dim strSaveSettings As String
            'set "strSaveSettings" True
            strSaveSettings = "True"
            'stores data in the value field of
            ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
            lngReturnValues = RegSetValueEx(hKey, "SaveSettingsOnExit", 0, REG_SZ, _
                                        ByVal strSaveSettings, Len(strSaveSettings))
        
            Dim strShowToolTips As String
            'set "Show ToolTips" to True
            strShowToolTips = "True"
            'stores data in the value field of
            ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
            lngReturnValues = RegSetValueEx(hKey, "ShowToolTips", 0, REG_SZ, _
                                        ByVal strShowToolTips, Len(strShowToolTips))
            
            Dim strIntroPlayLength As String
            'set the value of "Intro Play Length"
            strIntroPlayLength = "10"
            'stores data in the value field of
            ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
            lngReturnValues = RegSetValueEx(hKey, "IntroPlayLength", 0, REG_SZ, _
                                    ByVal strIntroPlayLength, Len(strIntroPlayLength))
            Call RegSmallFont
            Call RegLargeFont
            Call RegRandomOrder
            Call RegContinuousPlay
            Call RegIntroPlay
        End If
    
        'opens the path "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Dimonsion"
        If OpenRegistryDimension = True Then
            'Create these keys
            Call RegFormLeftAndTop
        End If
        'close the registry key
        Call CloseRegistry
End Sub

Public Function OpenRegistrySettings() As Boolean
    '* Purpose: Opens the program registry key specified by "strKeyPath"
    
    'opens the path specified by "strKeyPath"
    strKeyPath = "Software\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegOpenKeyEx(HKEY_CURRENT_USER, strKeyPath, 0, _
                                    KEY_ALL_ACCESS, hKey)
        'check if error occured
        If lngReturnValues = ERROR_SUCCESS Then
            'if not then
            OpenRegistrySettings = True
        Else
            'otherwise
            OpenRegistrySettings = False
        End If
End Function

Public Function OpenRegistryDimension() As Boolean
    '* Purpose: Opens the program registry key specified by "strKeyPath"
    
    'opens the path specified by "strKeyPath"
    strKeyPath = "Software\SimonSoft\Simon's CD Player\Dimension"
    lngReturnValues = RegOpenKeyEx(HKEY_CURRENT_USER, strKeyPath, 0, _
                            KEY_ALL_ACCESS, hKey)
        'check if error occured
        If lngReturnValues = ERROR_SUCCESS Then
            'if not then
            OpenRegistryDimension = True
        Else
            'otherwise
            OpenRegistryDimension = False
        End If
End Function

Public Sub RegPlayCDOnExit()
    '* Purpose: Creates a null termination string "StopPlayCDOnExit"
    '*          value in the registry

    Dim strStopCDValue As String
    'check if "Stop Playing CD On Exit" is checked in "Preferences" dialog box
    If frmPreferences.chkStopCDPlay.Value = Checked Then
        'set "strStopCDValue" to True
        strStopCDValue = "True"
    Else
        'otherwise set it to False
        strStopCDValue = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "StopPlayOnExit", 0, REG_SZ, _
                                ByVal strStopCDValue, Len(strStopCDValue))
End Sub

Public Sub RegSaveSettingsOnExit()
    '* Purpose: Creates a null termination string "SaveSettingsOnExit"
    '*          value in the registry

    Dim strSaveSettings As String
    'check if "Save Settings On Exit" is checked in "Preferences" dialog box
    If frmPreferences.chkSaveSettings.Value = Checked Then
        'set "strSaveSettings" True
        strSaveSettings = "True"
    Else
        'otherwise set it to False
        strSaveSettings = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "SaveSettingsOnExit", 0, REG_SZ, _
                                    ByVal strSaveSettings, Len(strSaveSettings))
End Sub

Public Sub RegShowToolTips()
    '* Purpose: Creates a null termination string "ShowToolTips"
    '*          value in the registry

    Dim strShowToolTips As String
    'check if "Show ToolTips" is checked in "Preferences" dialog box
    If frmPreferences.chkShowToolTips.Value = Checked Then
        'set "Show ToolTips" to True
        strShowToolTips = "True"
    Else
        'otherwise set it to False
        strShowToolTips = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "ShowToolTips", 0, REG_SZ, _
                                    ByVal strShowToolTips, Len(strShowToolTips))
End Sub

Public Sub RegIntroPlayLength()
    '* Purpose: Creates a null termination string "IntroPlayLength"
    '*          value in the registry

    Dim strIntroPlayLength As String
    'get the value of "Intro Play Length" from "Preferences" dialog box
    strIntroPlayLength = Trim$(frmPreferences.txtIntroPlay.Text)
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "IntroPlayLength", 0, REG_SZ, _
                                    ByVal strIntroPlayLength, Len(strIntroPlayLength))
End Sub

Public Sub RegSmallFont()
    '* Purpose: Creates a null termination string "SmallFont"
    '*          value in the registry

    Dim strSmallFont As String
    'check if "Small Font" is checked in "Preferences" dialog box
    If frmPreferences.optSmallFont.Value = True Then
        'set "strSmallFont" to True
        strSmallFont = "True"
    Else
        'otherwise set it to False
        strSmallFont = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "ShowSmallFont", 0, REG_SZ, _
                                    ByVal strSmallFont, Len(strSmallFont))
End Sub

Public Sub RegLargeFont()
    '* Purpose: Creates a null termination string "LargeFont"
    '*          value in the registry

    Dim strLargeFont As String
    'check if "Large Font" is checked in "Preferences" dialog box
    If frmPreferences.optLargeFont.Value = True Then
        'set "strLargeFont" to True
        strLargeFont = "True"
    Else
        'otherwise set it to False
        strLargeFont = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "ShowLargeFont", 0, REG_SZ, _
                                    ByVal strLargeFont, Len(strLargeFont))
End Sub

Public Sub CloseRegistry()
    '* Purpose: Closes the registry key
    
    'close the registry key
    lngReturnValues = RegCloseKey(hKey)
End Sub

Public Sub RegFormLeftAndTop()
    '* Purpose: Creates null termination strings "Left" and "Top" to store
    '*          main window Left and Top coordinates

    'first open "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    If OpenRegistryDimension = True Then

        Dim strLeft As String
        strLeft = Str(frmCdplay.Left)
        'stores data in the value field of
        ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Dimension"
        lngReturnValues = RegSetValueEx(hKey, "Left", 0, REG_SZ, _
                                        ByVal strLeft, Len(strLeft))
            
        Dim strTop As String
        strTop = Str(frmCdplay.Top)
        'stores data in the value field of
        ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Dimension"
        lngReturnValues = RegSetValueEx(hKey, "Top", 0, REG_SZ, _
                                        ByVal strTop, Len(strTop))
                                    
        'close registry
        Call CloseRegistry
    End If
End Sub

Public Sub SaveToRegistry()
    '* Purpose: Saves the current data from "Preferences" dialog box and updates
    '*          registry

    Call OpenRegistrySettings
    Call RegPlayCDOnExit
    Call RegSaveSettingsOnExit
    Call RegShowToolTips
    Call RegIntroPlayLength
    Call RegSmallFont
    Call RegLargeFont
    Call CloseRegistry
End Sub

Public Sub GetFromRegistry()
    '* Purpose: Gets the current data from registry and updates the "Preferences"
    '*          dialog box

    Dim strBuffer As String * 40
    Dim lngBufferSize As Long
    'get the states(i.e whether they are checked or not) of check boxes from registry
    'first open "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    Call OpenRegistrySettings
    lngBufferSize = Len(strBuffer)
    'retrieve state of "StopPlayOnExit" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "StopPlayOnExit", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of combo box to checked
            frmPreferences.chkStopCDPlay.Value = Checked
        End If
        
    'retrieve state of "SaveSettingsOnExit" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "SaveSettingsOnExit", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of combo box to checked
            frmPreferences.chkSaveSettings.Value = Checked
        End If
        
    'retrieve state of "ShowToolTips" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "ShowToolTips", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of combo box to checked
            frmPreferences.chkShowToolTips.Value = Checked
        End If
        
    'retrieve state of "ShowSmallFont" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "ShowSmallFont", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of check box to True
            frmPreferences.optSmallFont.Value = True
        End If
    
    'retrieve state of "ShowLargeFont" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "ShowLargeFont", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of check box to True
            frmPreferences.optLargeFont.Value = True
        End If
    
    'retrieve value of "IntroPlayLength" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "IntroPlayLength", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
    
    'set the value of "Intro Play Length" to "Preferences" dialog box
    frmPreferences.txtIntroPlay.Text = Trim$(strBuffer)
    
    'close the registry
    Call CloseRegistry
End Sub

Public Sub RegRandomOrder()
    '* Purpose: Creates a null termination string "RandomOrder"
    '*          and sets its value in the registry

    Dim strRandomOrder As String
    'check if menuitem "Random Order" is checked
    If frmCdplay.mnuOptionsRandomOrder.Checked = True Then
        'set "strRandomOrder" to True
        strRandomOrder = "True"
    Else
        'otherwise set it to False
        strRandomOrder = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "RandomOrder", 0, REG_SZ, _
                                    ByVal strRandomOrder, Len(strRandomOrder))
End Sub

Public Sub RegContinuousPlay()
    '* Purpose: Creates a null termination string "ContinuousPlay"
    '*          and sets its value in the registry

    Dim strContinuousPlay As String
    'check if menuitem "Continuous Play" is checked
    If frmCdplay.mnuOptionsContinuousPlay.Checked = True Then
        'set "strContinuousPlay" to True
        strContinuousPlay = "True"
    Else
        'otherwise set it to False
        strContinuousPlay = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "ContinuousPlay", 0, REG_SZ, _
                                    ByVal strContinuousPlay, Len(strContinuousPlay))
End Sub

Public Sub RegIntroPlay()
    '* Purpose: Creates a null termination string "IntroPlay"
    '*          and sets its value in the registry

    Dim strIntroPlay As String
    'check if menuitem "Intro Play" is checked
    If frmCdplay.mnuOptionsIntroPlay.Checked = True Then
        'set "strIntroPlay" to True
        strIntroPlay = "True"
    Else
        'otherwise set it to False
        strIntroPlay = "False"
    End If
    'stores data in the value field of
    ' "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    lngReturnValues = RegSetValueEx(hKey, "IntroPlay", 0, REG_SZ, _
                                    ByVal strIntroPlay, Len(strIntroPlay))
End Sub

Public Sub GetMenuStateFromRegistry()
    '* Purpose: Gets the menuitems states form Registry and updates the menuitems
    
    Dim strBuffer As String * 40
    Dim lngBufferSize As Long
    'get the states(i.e whether they are checked or not) of menuitems from registry
    'first open "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Settings"
    Call OpenRegistrySettings
    lngBufferSize = Len(strBuffer)
    'retrieve state of "RandomOrder" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "RandomOrder", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of menuitem to checked
            frmCdplay.mnuOptionsRandomOrder.Checked = True
        End If
    
    'retrieve state of "ContinuousPlay" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "ContinuousPlay", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of menuitem to checked
            frmCdplay.mnuOptionsContinuousPlay.Checked = True
        End If
    
    'retrieve state of "IntroPlay" from currently opened registry key
    lngReturnValues = RegQueryValueEx(hKey, "IntroPlay", 0, REG_SZ, _
                                    ByVal strBuffer, lngBufferSize)
        'check if "strBuffer" is True
        If Mid$(strBuffer, 1, lngBufferSize - 1) = "True" Then
            'save the state of menuitem to checked
            frmCdplay.mnuOptionsIntroPlay.Checked = True
        End If
    
    'close the registry
    Call CloseRegistry
End Sub

Public Sub RegGetFormDimension()
    '* Purpose: Gets the main windows Left And Top position which was saved before
    '*          it was closed last time
    
    Dim strBuffer As String * 40
    Dim lngBufferSize As Long
    
    'first open "HKEY_CURRENT_USER\SimonSoft\Simon's CD Player\Dimension"
    If OpenRegistryDimension = True Then
        lngBufferSize = Len(strBuffer)
    
        'retrieve the value of "Left" from currently opened registry key
        lngReturnValues = RegQueryValueEx(hKey, "Left", 0, REG_SZ, _
                                        ByVal strBuffer, lngBufferSize)
        'set the main window Left coordinate
        frmCdplay.Left = Int(strBuffer)
    
        'retrieve the value of "Top" from currently opened registry key
        lngReturnValues = RegQueryValueEx(hKey, "Top", 0, REG_SZ, _
                                        ByVal strBuffer, lngBufferSize)
        'set the main window Top coordinate
        frmCdplay.Top = Int(strBuffer)
                                        
        'close registry
        Call CloseRegistry
    End If
End Sub
