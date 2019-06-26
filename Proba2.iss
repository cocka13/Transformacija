; InnoScript Version 2.2.2
; Randem Systems, Inc.
; Copyright 2003, 2004
; website:  http://www.randem.com
; email:  innoscript@randem.com

; Date: veljaèa 23, 2005

;         Visual Basic 6 Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript\InnoScript 2.2\VB6 Runtime\
;            Visual Basic Project File (.vbp):   C:\VB_Project\Transform\Transform.vbp
;        Inno Setup Script Output File (.iss):   C:\VB_Project\Transform\Proba2.iss
;Visual Basic Project Application File (.exe):   C:\VB_Project\Transform\Transform.exe

; ------------------------
; Visual Basic References
; ------------------------

; OLE Automation


; --------------------------
; Visual Basic Components
; --------------------------

; Microsoft Common Dialog Control 6.0


[Setup]
AppName=Transform
AppVerName=Transform 1.0.0.0
DefaultGroupName=Transformacija
AppPublisher= Mladi destruktivci
;AppPublisherURL=http://www.yourwebsite.com
AppVersion=1.0.0.0
;AppSupportURL=http://www.yourwebsite.com
;AppUpdatesURL=http://www.yourwebsite.com
AllowNoIcons=yes
;InfoBeforeFile=Setup.txt
;InfoAfterFile=ReadMe.txt
;WizardImageFile=yourlogo.bmp
AppCopyright=c2005 Sva prava pridržana
PrivilegesRequired=admin
OutputBaseFilename=Transform1000
DefaultDirName={pf}\Transform

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
Name: "quicklaunchicon"; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\olepro32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\comcat.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\stdole2.tlb"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  uninsneveruninstall restartreplace sharedfile regtypelib
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\asycfilt.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\oleaut32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\msvbvm60.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\vb6stkit.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  restartreplace sharedfile
Source: "C:\WINNT\system32\comdlg32.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "c:\vb_project\transform\transform.exe"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion


[Icons]
Name: "{group}\Transform"; Filename: "{app}\Transform.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall Transform"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Transform"; Filename: "{app}\Transform.exe"; Tasks: desktopicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\Transform.exe"; Description: "Launch Transform"; Flags: nowait postinstall skipifsilent; WorkingDir: "{app}"

[UninstallDelete]
Type: files; Name: "{app}\Transform.url"
