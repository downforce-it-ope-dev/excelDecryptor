[Setup]
AppId={{B8113C7E-CFB9-420E-A2A7-188359CA5C09}
AppName=Excel Decryptor
AppVersion=1.0.0
AppPublisher=downforceITkkt
DefaultDirName={autopf}\ExcelDecryptor
DefaultGroupName=Excel Decryptor
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=ExcelDecryptorSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 아이콘 만들기"; GroupDescription: "추가 작업:"; Flags: unchecked

[Files]
Source: "dist\ExcelDecryptor.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Excel Decryptor"; Filename: "{app}\ExcelDecryptor.exe"
Name: "{autodesktop}\Excel Decryptor"; Filename: "{app}\ExcelDecryptor.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ExcelDecryptor.exe"; Description: "Excel Decryptor 실행"; Flags: nowait postinstall skipifsilent
