; -------------------------------------------------------------------------------------------------------------------------------------------------
; ������ ��� ��������� ��������� MS OFFICE
; -------------------------------------------------------------------------------------------------------------------------------------------------
; �������� ���������� �����������
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define AppName      "DomesticTransports"                                            ; �������� ����������
#define AppVersion   "3.12.2.0"                                               ; ������ ���������
#define AppPublisher "Micro-Solution LLC"
#define AppURL       "https://micro-solution.ru"
#define AppGUI       "63EFF963-1E2F-41A0-8AD8-0D23FB6B4087"

#define ProjectPath  "C:\Users\aevse\Source\Repos\DomesticTransport\"
#define SetupPath    ProjectPath + "Setup\"                        
#define AppIco       SetupPath + "icon-application.ico"                       ; ���� � �������

#define FilesPath    ProjectPath + "DomesticTransport\bin\Release\"                  ; ����� � �������, ������� ���������� ���������
#define ReleasePath  SetupPath + "Release\"                                   ; �������� �����

#define TypeAddIn    "Excel"                                                   ; Word or Excel

; -------------------------------------------------------------------------------------------------------------------------------------------------
; ��������� NetFramework 
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define NeedNetFramework 1                                                   ; 0/1
#define NetFrameworkVerName "4.5"
;�������� ����� ����������� ������ ������ NetFramework. ������ ������ � SetupPath
#define NetFrameworkFileSetup "dotNetFx45_Full_setup.exe"                         ; 4.5
;#define NetFrameworkSetup "NDP472-KB4054530-x86-x64-AllOS-ENU.exe"           ; 4.7.2  Full

; -------------------------------------------------------------------------------------------------------------------------------------------------
; ������������ ���������
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define SignTool    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
#define SingNameSSL AppPublisher ; ��� �����������


[Setup]
;������������ ����
SignTool=byparam {#SignTool} sign /a /n $q{#SingNameSSL}$q /t http://timestamp.comodoca.com/authenticode /d $q{#AppName}$q $f

;������������ ������������� VS GUI
AppId            = {{{#AppGUI}}
AppName          = {#AppName}
AppVersion       = {#AppVersion}
AppPublisher     = {#AppPublisher}
AppPublisherURL  = {#AppURL}

;AppSupportURL    = {#AppURL}
;AppUpdatesURL    = {#AppURL}

DefaultDirName       = {autopf}\Micro-Solution\{#AppName}
DefaultGroupName     = Micro-Solution\{#AppName}
UninstallDisplayIcon ={#AppIco}
UninstallDisplayName ={#AppName}
AllowNoIcons         = yes

;���� ������������� ���������� ��� �������������
;LicenseFile = {#FilesPath}License.rtf

PrivilegesRequired=none

; ��������� ���������� �����������
OutputDir            = {#ReleasePath}
OutputBaseFilename   = Setup {#AppName}
SetupIconFile        = {#AppIco}
Compression          = lzma
SolidCompression     = yes
WizardStyle          = modern
WizardImageFile      = {#SetupPath}WizardImage.bmp 
WizardSmallImageFile = {#SetupPath}WizardSmallImage.bmp
DisableWelcomePage   = no

[Languages]
;Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Messages]
WelcomeLabel1=��� ������������ ������ ��������� ��������� [name]
WelcomeLabel2=��������� ��������� [name/ver] �� ��� ���������.%n%n����������, �������� ��� ����� {#TypeAddIn} ����� ���, ��� ����������.
ReadyLabel1=��� ��������� ��������� � ����� ���������� � ��������� [name] �� ��� ���������.
FinishedLabel=��������� [name] ����������� �� ��� ���������. ��������� ����������� ������ � ���������� Microsoft {#TypeAddIn}.

[Files]
Source: "{#FilesPath}{#AppName}.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}Microsoft.Office.Tools.Excel.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}DomesticTransport.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AppIco}"; DestDir: "{app}"; Flags: ignoreversion

; .NET Framework 4.5
Source: "{#SetupPath}{#NetFrameworkFileSetup}"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not IsDotNetDetected

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#AppName}}"; Filename: "{#AppURL}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

[Registry]
;Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}";  Flags: uninsdeletekey
;Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
;Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
;Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\{#AppName}.vsto|vstolocal"; Flags: uninsdeletekey

[Code]


// ��������� ������ ������ ���������� � �������
function GetFrameworkVer(const AppName: String): cardinal;
  begin
    Result := 0;
    case AppName of
      '4.5'   :Result := 378389;
      '4.5.1'	:Result := 378675;
      '4.5.2'	:Result := 379893;
      '4.6'   :Result := 393295;
      '4.6.1' :Result := 394254;
      '4.6.2' :Result := 394802;
      '4.7'	  :Result := 460798;
      '4.7.1'	:Result := 461308;
      '4.7.2'	:Result := 461808;
      '4.8'   :Result := 528040;	
    end;
  end;

function IsDotNetDetected(): boolean;
  var 
    reg_key: string; // ��������������� ��������� ���������� �������
    full_key: string;
    success: boolean; // ���� ������� ������������� ������ .NET
    release_number: cardinal; // ����� ������ ��� ������ 4.5.x
    sub_key: string;
  begin
    success := false;
    reg_key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\';
    
    // ������ 4.5 � ����
    sub_key := 'v4\Full';
    full_key := reg_key + sub_key;
    success := RegQueryDWordValue(HKLM, full_key, 'Release', release_number);
    success := success and (release_number >= GetFrameworkVer('{#NetFrameworkVerName}'));
    result := success;
  end;


// ����� ����������� ����������
function FindApp(const AppName: String): Boolean;
  var
    WMIService:    Variant;
    WbemLocator:   Variant;
    WbemObjectSet: Variant;
  begin
    WbemLocator   := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService    := WbemLocator.ConnectServer('localhost', 'root\CIMV2');
    WbemObjectSet :=
      WMIService.ExecQuery('SELECT * FROM Win32_Process Where Name="' + AppName + '"');
    if not VarIsNull(WbemObjectSet) and (WbemObjectSet.Count > 0) then
    begin
      Log(AppName + ' is up and running');
      Result := True
    end;
  end;

function GetNameApp(const TypeAddIn: String): String;
  begin
    case TypeAddIn of
      'Excel' :Result := 'excel.exe';
      'Word'	:Result := 'winword.exe';	
    end;
  end;


 //Callback-�������, ���������� ��� ������������� ���������
procedure InitializeWizard();
  begin
      // �������� ����� ����������
  end;


// ����� ������� ������ �����
function NextButtonClick(CurPageID: Integer): Boolean;
  begin
    Result := True;

    // ����� �����������
    //case CurPageID of wpWelcome:
    //  if (FindApp(GetNameApp('{#TypeAddIn}'))) then
    //  begin
    //    MsgBox('����������, �������� ��� ����� {#TypeAddIn} ����� ���������� ���������!', mbError, MB_OK);
    //    Result := False;
    //  end;
    //end;

  end;

// ����� ������� �������������
function  InitializeUninstall(): Boolean;
  begin
    Result := True;
    //if (FindApp(GetNameApp('{#TypeAddIn}'))) then
    //begin
    //  MsgBox('����������, �������� ��� ����� {#TypeAddIn} ����� ��������� ���������!', mbError, MB_OK);
    //  Result := False;
    //end;
    
  end;
[Run]
Filename: {tmp}\{#NetFrameworkFileSetup}; Parameters: "/q:a /c:""install /l /q"""; Check: not IsDotNetDetected; StatusMsg: Microsoft Framework 4.5 is installed. Please wait...