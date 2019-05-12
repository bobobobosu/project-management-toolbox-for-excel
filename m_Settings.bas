Attribute VB_Name = "m_Settings"
Option Explicit

Public Const gsRegKey As String = "EC_ListSearch"  'The registry entries' name
Public Const gsKeyName As String = "Direction"
Public gsDirection As String
Public gsMenu As String
Public gsOpen As String


Sub GetDirection()
    'Get the shortcutkey, use "n" if anything goes wrong
    gsDirection = GetSetting(gsRegKey, "Settings", "Direction", 0)
End Sub

Sub SaveDirection(lDir As Long)
    'Save the shortcutkey to the registry
    SaveSetting gsRegKey, "Settings", "Direction", lDir
End Sub

Sub GetShowMenu()
    'Get the shortcutkey, use "n" if anything goes wrong
    gsMenu = GetSetting(gsRegKey, "Settings", "ShowMenu", False)
End Sub

Sub SaveShowMenu(bMenu As Boolean)
    'Save the shortcutkey to the registry
    SaveSetting gsRegKey, "Settings", "ShowMenu", bMenu
End Sub

Sub GetOpenOnSelectionChange()
    'Get the shortcutkey, use "n" if anything goes wrong
    gsOpen = GetSetting(gsRegKey, "Settings", "OpenOnSelectionChange", False)
End Sub

Sub SaveOpenOnSelectionChange(bOpen As Boolean)
    'Save the shortcutkey to the registry
    SaveSetting gsRegKey, "Settings", "OpenOnSelectionChange", bOpen
End Sub


Sub Deletesettings()
    'Remove all registry entries belonging to this application
    DeleteSetting gsRegKey
End Sub


