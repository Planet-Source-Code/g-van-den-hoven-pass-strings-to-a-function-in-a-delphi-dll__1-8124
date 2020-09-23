<div align="center">

## Pass Strings To A Function In A Delphi DLL


</div>

### Description

The code was made by me because i did not find any information about how to pass a VB string to a delphi dll function (accepting strings) hope this code helps.<BR><BR>This code shows how to call a delphi DLL and how to code the Delphi dll to accept strings from VB. The code (2 demo projects, 1 in VB and 1 in Delphi) are pretty simple, but you get the idea.

FEEL FREE TO VOTE FOR ME!
 
### More Info
 
The code can be easily adjusted to implement full DELPHI applications (e.g. if you want some things done in Delphi, from within a VB application)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[G\. van den Hoven](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/g-van-den-hoven.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/g-van-den-hoven-pass-strings-to-a-function-in-a-delphi-dll__1-8124/archive/master.zip)





### Source Code

```
-------------------
VB CODE:
(START A NEW PROJECT, REPLACE THE EXISTING CODE WITH THIS:)
-------------------
Option Explicit
Private Declare Function CountLetters Lib "..\delphi\project1.dll" (ByVal Str As String) As Long
Private Sub Form_Load()
 Call CountLetters("This is a teststring, passed to a function in a delphi DLL")
End Sub
-------------------
DELPHI CODE
(START A NEW LIBRARY, REPLACE THE EXISTING CODE WITH THIS:)
-------------------
library Project1;
uses
 Windows,
 SysUtils;
function CountLetters(pData : PChar) : Cardinal; export; stdcall;
var
 Handle : Integer;
 tMsg : String;
begin
 tMsg := 'The string passed by you; "' + pData + '" is counting ' + IntToStr(Length(pData)) + ' letters.';
 MessageBoxA(Handle, pChar(tMsg), 'Delphi DLL', MB_OK);
end;
exports
 CountLetters name 'CountLetters' resident;
begin
end.
```

