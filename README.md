<div align="center">

## Convert INI Files to XML


</div>

### Description

Take your Desktop apps into this decade by using XML files for INI Files. This Module will translate or convert your exsisting INI files into XML. If you are interested in the source code to access elements in your XML Files in place of INI Values, let me know
 
### More Info
 
This code uses MSXML Parser Version 3 (XML DOM). You can get it from microsoft.com or change MSXML2.Properties to MSXML.Properties and set a reference in your code to Microsoft XML


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Bender](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-bender.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-bender-convert-ini-files-to-xml__1-15076/archive/master.zip)

### API Declarations

```
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA"
```


### Source Code

```
Option Explicit
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function INI_to_XML(INIFile_IN As String, XMLFile_Out As String) As Boolean
  Dim iFile As Integer
  Dim oXMLDocument As New MSXML2.DOMDocument
  Dim oXMLBlock As MSXML2.IXMLDOMNode
  Dim oXMLSectionListBlock As MSXML2.IXMLDOMNode
  Dim oXMLSectionBlock As MSXML2.IXMLDOMNode
  Dim oXMLKeyListBlock As MSXML2.IXMLDOMNode
  Dim oXMLKeyBlock As MSXML2.IXMLDOMNode
  Dim oNode As MSXML2.IXMLDOMNode
  '-- Create Initial Blocks
  Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "INISchema", "")
  Set oXMLBlock = oXMLDocument.appendChild(oNode)
  Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "SectionList", "")
  Set oXMLSectionListBlock = oXMLBlock.appendChild(oNode)
  '-- Write a SectionList count tag and fill it in later
  Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Count", "")
  oXMLSectionListBlock.appendChild oNode
  '-- Loop through each line and find sections
  iFile = FreeFile
  Dim sWorking As String
  Dim iCount As Integer
  Open Trim(INIFile_IN) For Input As iFile
  Do Until EOF(iFile)
    Line Input #iFile, sWorking
    sWorking = Trim(sWorking)
    If Left$(sWorking, 1) = "[" And Right$(sWorking, 1) = "]" Then
      '-- Section Found Add to XML Document
      Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Section", "")
      Set oXMLSectionBlock = oXMLSectionListBlock.appendChild(oNode)
      Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Name", "")
      oNode.Text = Mid$(sWorking, 2, Len(sWorking) - 2)
      oXMLSectionBlock.appendChild oNode
      '-- Get keys from current Section
      Dim iRetCode As Integer
      Dim sBuf As String
      Dim sSize As String
      Dim sKeys As String
      sBuf = Space$(1024)
      sSize = Len(sBuf)
      iRetCode = GetPrivateProfileSection(oNode.Text, sBuf, sSize, INIFile_IN)
      If (sSize > 0) Then
        sKeys = Left$(sBuf, iRetCode)
        Dim arKeys() As String
        Dim sKey As String
        Dim sValue As String
        arKeys = Split(sKeys, vbNullChar)
        If Not isArrayEmpty(arKeys) Then
          '-- We have at least one Key so Build a KeyList Block
          Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "KeyList", "")
          Set oXMLKeyListBlock = oXMLSectionBlock.appendChild(oNode)
          '-- Write a count tag and fill it in later
          Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Count", "")
          oXMLKeyListBlock.appendChild oNode
          For iCount = LBound(arKeys) To UBound(arKeys)
            If arKeys(iCount) <> "" Then
              If InStr(1, arKeys(iCount), "=") <> 0 Then
                sKey = Left$(arKeys(iCount), InStr(1, arKeys(iCount), "=") - 1)
                sValue = Right$(arKeys(iCount), Len(arKeys(iCount)) - InStr(1, arKeys(iCount), "="))
              Else
                sKey = arKeys(iCount)
                sValue = ""
              End If
              Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Key", "")
              Set oXMLKeyBlock = oXMLKeyListBlock.appendChild(oNode)
              Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Name", "")
              oNode.Text = sKey
              oXMLKeyBlock.appendChild oNode
              Set oNode = oXMLDocument.createNode(NODE_ELEMENT, "Value", "")
              oNode.Text = sValue
              oXMLKeyBlock.appendChild oNode
            End If
          Next
          '-- Add the KeyList Count
          oXMLKeyListBlock.childNodes(0).Text = oXMLKeyListBlock.childNodes.length - 1
        End If
      Else
        sKeys = ""
      End If
    End If
  Loop
  '-- Add the SectionList Count
  oXMLSectionListBlock.childNodes(0).Text = oXMLSectionListBlock.childNodes.length - 1
  Close iFile
  oXMLDocument.save XMLFile_Out
Cleanup:
  Set oXMLDocument = Nothing
  Exit Function
Err_Handler:
  INI_to_XML = False
  GoTo Cleanup
End Function
Private Function isArrayEmpty(arr As Variant) As Boolean
 Dim i
 isArrayEmpty = True
 On Error Resume Next
 i = UBound(arr) ' cause an error if array is empty
 If Err.Number > 0 Then Exit Function
 isArrayEmpty = False
End Function
```

