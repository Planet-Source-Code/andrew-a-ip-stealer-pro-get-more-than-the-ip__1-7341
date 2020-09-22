Attribute VB_Name = "Module1"
Public intTempFix As Integer
Public ScriptPath As String
Public ErrorText As String
Public IPData(0 To 99) As typIPData

Type typIPData
FreeSocket As Boolean
End Type

Sub AddNewIP(strIP As String, strOS As Variant, strBrowser As Variant)
With frmMain
  .lvIPs.ListItems.Add = strIP
  .lvIPs.ListItems.Item(.lvIPs.ListItems.Count).SubItems(1) = Time 'NOW = Date + Time (VB Command)
  .lvIPs.ListItems.Item(.lvIPs.ListItems.Count).SubItems(2) = strBrowser
  .lvIPs.ListItems.Item(.lvIPs.ListItems.Count).SubItems(3) = strOS
End With
End Sub

