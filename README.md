<div align="center">

## Download Binary \(graphics/executables\) over http without Inet Controls


</div>

### Description

Download Binary (graphics/executables) over http without Inet Controls or Api's. This program uses xmlhttp to grab it and then writes it to a file. This is very helpful for an app that is distibuted on multiple platforms because the inet controls are so buggy when trying to install on user that don't have it.
 
### More Info
 
Url

The user needs to have XMLDOM on the computer

Binary


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bradley Dick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bradley-dick.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bradley-dick-download-binary-graphics-executables-over-http-without-inet-controls__1-26910/archive/master.zip)





### Source Code

```
' Code By: Bradley Dick (bradley@bdick.com)
'
'
'
' Use this code as much as you like it really saved me a lot of headache. All I ask is that you reply and vote. Thanks All this is my first post.......
'
'
' Explantion:
' strUrl is the path to the file you want ie.(http://www.abc.com/logo.gif)
' strFile is the path where you want to save the file and the name. ie.(C:\logo.gif)
'That's it... Just make sure that you include the Microsoft XML 3.0 or if you use a different on then just change the dim and the set statement..
Private Sub GetFile(strURL As String, strFile As String)
    Dim xml As MSXML2.XMLHTTP30
    Dim X() As Byte
    Set xml = New MSXML2.XMLHTTP30
      xml.Open "GET", strURL, False
      xml.send
      X = xml.responseBody
  Open strFile For Binary Access Write As #1
      Put #1, , X()
      Close #1
End sub
```

