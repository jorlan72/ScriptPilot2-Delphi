unit intercode;

interface

uses
  Registry, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.ValEdit, System.NetEncoding,
  LbClass, LbCipher, LbUtils, ShellAPI, System.Generics.Collections, VCL.Forms, mymessage;

type
// CLASS FOR HOLDING VALUE LIST DATA STRUCTURES IN MEMORY
  TMyValueList = class
  private
    FList: TDictionary<string, string>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddOrUpdate(Key, Value: string);
    function GetValue(Key: string): string;
    procedure DeleteKey(Key: string);
    function KeyExists(Key: string): Boolean;
    procedure Clear;
    function GetAllKeys: TArray<string>;
    function Count: Integer;
  end;

// A RECORD HOLDING DATA TO BE USED IN A MEMORY DATA STRUCTURE
type
  TDataInfo = record
    DataIndex: Integer;
    DataName: string;
    DataDescription: string;
    DataStringParameters: array[0..5] of string;
    DataIntegerParameters: array[0..5] of Integer;
  end;

type
// CLASS FOR HOLDING STRING DATA STRUCTURES IN MEMORY
  TMyDataList = class
  private
    FList: TDictionary<string, TDataInfo>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure DeleteKey(Key: string);
    procedure AddOrUpdate(Key: string; Data: TDataInfo);
    function KeyExists(Key: string): Boolean;
    procedure Clear;
    function GetAllKeys: TArray<string>;
    function Count: Integer;
  end;

type
// CLASS FOR GENERIC THREAD GENERATION OF PROCEDURES
// EXECUTE IS THE ACTION TO MULTITHREAD AND FINISH IS WHAT IS QUEUED TO THE MAIN THREAD WHEN IT IS DONE
  TGenericThread = class(TThread)
  private
    FProc: TProc;
    FOnFinish: TProc;
  protected
    procedure Execute; override;
    procedure DoOnFinish;
  public
    constructor Create(const AProc: TProc; const AOnFinish: TProc = nil);
    property OnFinish: TProc read FOnFinish write FOnFinish;
  end;

const
  RegistryKey = 'Software\CodeRed\ScriptPilot\';
  InternalSeed = 'Potus#45lolNexus72jXtc569n932Stamina@72j';

var
  Variables : TMyValueList;
  Runners : TMyDataList;
  Connections : TMyDataList;
  EditVariables : TMyValueList;
  EditTags : TMyValueList;
  CurrentCommandCounter : Integer;
  EchoReply : String;
  OKButtonPressed : Boolean;
  AbortButtonClicked : Boolean;
  AreRunnerReady : Boolean;
  CurrentRunner : String;
  MaxNumberOfRunners, CurrentNumberOfRunners : Integer;
  FileRunnerExe : String;
  DidScriptComplete : Boolean;
  WarningHappened : Boolean;
  ScriptRunning : Boolean;

Function ReadRegistryString(const RootKey: HKEY; const Key, ValueName: string; const DefaultValue: string = ''): string; // READ VALUES FROM REGISTRY - REQUIRES UNIT REGISTRY
Function WriteRegistryString(const RootKey: HKEY; const Key, ValueName, Value: string): Boolean; // WRITE VALUES TO REGISTRY - REQUIRES UNIT REGISTRY
function GetApplicationVersion: string; // GET THE VERSION BUILD
procedure SendWinMessage(WindowsCaptionName: String; MessageToSend: String; MessageId: Cardinal); // USED TO SEND MESSAGES BETWEEN RUNNING APPLICATIONS BASED ON TITLE CAPTION OF FORM
function HexToString(const Hex: string): string; // CONVERTS STRING TO HEX. USED BY CRYPTO PROCEDURES
function StringToHex(const Buffer: string): string; // CONVERTS HEX TO STRING. USED BY CRYPTO PROCEDURES
function EncryptString(const APlainText, APassword: string): string; // GENERATE A RANDOM KEY TO PREFIX ENCRYPTION PASSWORDS FOR CREATING UNIQUE ENCRYPTED STRINGS
function DecryptString(const ACipherText, APassword: string): string; // ENCRYPT A STRING USING THE PROVIDED PASSWORD
function GenerateRandomKey(Length: Integer): string; // DECRYPT A STRING USING THE PROVIDED PASSWORD
procedure OpenRegeditToKey(const KeyPath: string); // OPEN THE REGISTRY EDITOR AND SET THE ACTIVE KEY TO CODERED
Procedure Log(logtext : string); // LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - NORMAL LOG - CODE 2
Procedure Debug(logtext : string); // LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - DEBUG LOG - CODE 3
Procedure Error(logtext : string); // LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - ERROR LOG - CODE 7
Procedure Warning(logtext : string); // LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - WARNING LOG - CODE 8
Procedure Echo(receiver, sender : String); // SEND ECHO TO A WINDOWS FORM
procedure ValueListToString(MyValueList: TMyValueList; var ResultString: string); // CONVERT MEMORY VALUELIST TO STRING
procedure StringToValueList(const InputString: string; MyValueList: TMyValueList); // CONVERT STRING TO MEMORY VALUELIST
function VariantType(const V: Variant): Integer; // CHECK WHAT TYPE OF VARIANT WE ARE DEALING WITH
procedure LaunchProcess(const AppPath, Params: string; CreateNoWindow, WaitForProcessToEnd: Boolean); // TO LAUNCH A RUNNER OR ANY APPLICATION WITH OPTIONS

implementation

constructor TGenericThread.Create(const AProc: TProc; const AOnFinish: TProc = nil);
// CLASS FOR GENERIC THREAD GENERATION OF PROCEDURES
// Example : TGenericThread.Create(MultiThreaded-Procedure, Talk-to-GUI-Procedure);
begin
  inherited Create(True);
  FProc := AProc;
  FOnFinish := AOnFinish;
  FreeOnTerminate := True; // Automatically free thread after execution
  Start;
end;

procedure TGenericThread.Execute;
// CLASS FOR GENERIC THREAD GENERATION OF PROCEDURES
begin
  if Assigned(FProc) then FProc;
  if Assigned(FOnFinish) then Queue(DoOnFinish);
end;

procedure TGenericThread.DoOnFinish;
// CLASS FOR GENERIC THREAD GENERATION OF PROCEDURES
begin
  if Assigned(FOnFinish) then FOnFinish;
end;

constructor TMyValueList.Create;
// PART OF CLASS TMyValueList
begin
  inherited;
  FList := TDictionary<string, string>.Create;
end;

constructor TMyDataList.Create;
// PART OF CLASS TMyDataList
begin
  inherited;
  FList := TDictionary<string, TDataInfo>.Create;
end;

destructor TMyValueList.Destroy;
// PART OF CLASS TMyValueList
begin
  FList.Free;
  inherited;
end;

destructor TMyDataList.Destroy;
// PART OF CLASS TMyDataList
begin
  FList.Free;
  inherited;
end;

procedure TMyValueList.AddOrUpdate(Key, Value: string);
// PART OF CLASS TMyValueList
begin
  FList.AddOrSetValue(Key, Value);
end;

procedure TMyDataList.AddOrUpdate(Key: string; Data: TDataInfo);
begin
  FList.AddOrSetValue(Key, Data);
end;

function TMyValueList.GetValue(Key: string): string;
// PART OF CLASS TMyValueList
begin
  if not FList.TryGetValue(Key, Result) then
    Result := '';
end;

procedure TMyValueList.DeleteKey(Key: string);
// PART OF CLASS TMyValueList
begin
  FList.Remove(Key);
end;

procedure TMyDataList.DeleteKey(Key: string);
// PART OF CLASS TMyDataList
begin
  FList.Remove(Key);
end;

function TMyValueList.KeyExists(Key: string): Boolean;
// PART OF CLASS TMyValueList
begin
  Result := FList.ContainsKey(Key);
end;

function TMyDataList.KeyExists(Key: string): Boolean;
// PART OF CLASS TMyDataList
begin
  Result := FList.ContainsKey(Key);
end;

procedure TMyValueList.Clear;
// PART OF CLASS TMyValueList
begin
  FList.Clear;
end;

procedure TMyDataList.Clear;
// PART OF CLASS TMyDataList
begin
  FList.Clear;
end;

function TMyValueList.GetAllKeys: TArray<string>;
// PART OF CLASS TMyValueList
begin
  Result := FList.Keys.ToArray;
end;

function TMyDataList.GetAllKeys: TArray<string>;
// PART OF CLASS TMyDataList
begin
  Result := FList.Keys.ToArray;
end;

function TMyValueList.Count: Integer;
// PART OF CLASS TMyValueList
begin
  Result := FList.Count;
end;

function TMyDataList.Count: Integer;
// PART OF CLASS TMyDataList
begin
  Result := FList.Count;
end;

Function ReadRegistryString(const RootKey: HKEY; const Key, ValueName: string; const DefaultValue: string = ''): string;
// READ VALUES FROM REGISTRY - REQUIRES UNIT REGISTRY
var
  Reg: TRegistry;
begin
  Result := DefaultValue;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := RootKey;
    if Reg.OpenKeyReadOnly(Key) then
      begin
        if Reg.ValueExists(ValueName) then Result := Reg.ReadString(ValueName)
        else Result := DefaultValue;
        Reg.CloseKey;
      end;
  finally
    Reg.Free;
  end;
end;

Function WriteRegistryString(const RootKey: HKEY; const Key, ValueName, Value: string): Boolean;
// WRITE VALUES TO REGISTRY - REQUIRES UNIT REGISTRY
var
  Reg: TRegistry;
begin
  Result := False;
  Reg := TRegistry.Create;
  Reg.RootKey := RootKey;
  if Reg.OpenKey(Key, True) then
    begin
      Reg.WriteString(ValueName, Value);
      Result := True;
    end;
    Reg.Free;
end;

function GetApplicationVersion: string;
// GET THE VERSION BUILD
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then exit;
  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then exit;
  if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then exit;
  Result := Format('%d.%d.%d.%d', [LongRec(FixedPtr.dwFileVersionMS).Hi, LongRec(FixedPtr.dwFileVersionMS).Lo, LongRec(FixedPtr.dwFileVersionLS).Hi, LongRec(FixedPtr.dwFileVersionLS).Lo])
end;

procedure SendWinMessage(WindowsCaptionName: String; MessageToSend: String; MessageId: Cardinal);
// USED TO SEND MESSAGES BETWEEN RUNNING APPLICATIONS BASED ON TITLE CAPTION OF FORM
var
  WindowID: HWND;
  CopyDataStruct: TCopyDataStruct;
begin
  WindowID := FindWindow(nil, PChar(WindowsCaptionName));
  if WindowID > 0 then
  begin
    CopyDataStruct.dwData := MessageId;
    CopyDataStruct.cbData := (Length(MessageToSend) + 1) * SizeOf(Char);
    CopyDataStruct.lpData := PChar(MessageToSend);
    SendMessage(WindowID, WM_COPYDATA, WPARAM(0), LPARAM(@CopyDataStruct));
  end;
end;

function StringToHex(const Buffer: string): string;
// CONVERTS STRING TO HEX. USED BY CRYPTO PROCEDURES
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(Buffer) do
    Result := Result + IntToHex(Ord(Buffer[i]), 2);
end;

function HexToString(const Hex: string): string;
// CONVERTS HEX TO STRING. USED BY CRYPTO PROCEDURES
var
  i: Integer;
  ByteValue: string;
begin
  Result := '';
  for i := 1 to Length(Hex) div 2 do
  begin
    ByteValue := '$' + Copy(Hex, (i-1)*2+1, 2);
    Result := Result + Char(StrToInt(ByteValue));
  end;
end;

function GenerateRandomKey(Length: Integer): string;
// GENERATE A RANDOM KEY TO PREFIX ENCRYPTION PASSWORDS FOR CREATING UNIQUE ENCRYPTED STRINGS
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length do
    Result := Result + Chr(Random(256));
end;

function EncryptString(const APlainText, APassword: string): string;
// ENCRYPT A STRING USING THE PROVIDED PASSWORD
var
  Blowfish: TLbBlowfish;
  Encrypted, RandomKey, KeyWithPrefix: string;
begin
  Result := '';
  try
    Randomize;
    RandomKey := GenerateRandomKey(8);
    KeyWithPrefix := RandomKey + APassword;
    Blowfish := TLbBlowfish.Create(nil);
    try
      Blowfish.GenerateKey(KeyWithPrefix);
      Encrypted := Blowfish.EncryptString(APlainText);
      Result := 'Encrypted#' + StringToHex(RandomKey) + StringToHex(Encrypted);
    finally
      Blowfish.Free;
    end;
  except
    on E: Exception do Result := '';
  end;
end;

function DecryptString(const ACipherText, APassword: string): string;
// DECRYPT A STRING USING THE PROVIDED PASSWORD
var
  Blowfish: TLbBlowfish;
  Decrypted, RandomKey, KeyWithPrefix, EncryptedWithoutPrefix, CipherTextWithoutPrefix: string;
begin
  Result := '';
  try
    if Pos('Encrypted#', ACipherText) = 1 then CipherTextWithoutPrefix := Copy(ACipherText, Length('Encrypted#') + 1, MaxInt) else Exit('');
    RandomKey := HexToString(Copy(CipherTextWithoutPrefix, 1, 16));
    EncryptedWithoutPrefix := HexToString(Copy(CipherTextWithoutPrefix, 17, MaxInt));
    KeyWithPrefix := RandomKey + APassword;
    Blowfish := TLbBlowfish.Create(nil);
    try
      Blowfish.GenerateKey(KeyWithPrefix);
      Decrypted := Blowfish.DecryptString(EncryptedWithoutPrefix);
      Result := Decrypted;
    finally
      Blowfish.Free;
    end;
  except
    on E: Exception do Result := '';
  end;
end;

procedure OpenRegeditToKey(const KeyPath: string);
// OPEN THE REGISTRY EDITOR AND SET THE ACTIVE KEY TO CODERED
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit', True) then
    begin
      Reg.WriteString('LastKey', KeyPath);
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
  ShellExecute(0, 'open', 'regedit.exe', nil, nil, SW_NORMAL);
end;

Procedure Log(logtext : string);
// LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - NORMAL LOG - CODE 2
begin
 SendWinMessage('ScriptPilot 2',logtext, 2);
end;

Procedure Debug(logtext : string);
// LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - DEBUG LOG - CODE 3
begin
 SendWinMessage('ScriptPilot 2',logtext, 3);
end;

Procedure Error(logtext : string);
// LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - ERROR LOG - CODE 7
begin
 SendWinMessage('ScriptPilot 2',logtext, 7);
end;

Procedure Warning(logtext : string);
// LOG VIA SENDMESSAGE TO SCRIPTPILOT MAIN FORM - WARNING LOG - CODE 8
begin
 SendWinMessage('ScriptPilot 2',logtext, 8);
end;

Procedure Echo(receiver, sender : String);
// SEND WINMESSAGE PING
begin
  SendWinMessage(receiver, sender, 0);
end;

procedure ValueListToString(MyValueList: TMyValueList; var ResultString: string);
var
  Key: string;
begin
  ResultString := '';
  for Key in MyValueList.GetAllKeys do
  begin
    ResultString := ResultString + Key + '=' + MyValueList.GetValue(Key) + #13#10;
  end;
end;

procedure StringToValueList(const InputString: string; MyValueList: TMyValueList);
var
  StringList: TStringList;
  i: Integer;
begin
  StringList := TStringList.Create;
  try
    StringList.Text := InputString;
    MyValueList.Clear;
    for i := 0 to StringList.Count - 1 do
    begin
      if Trim(StringList.Names[i]) <> '' then MyValueList.AddOrUpdate(StringList.Names[i], StringList.ValueFromIndex[i]);
    end;
  finally
    StringList.Free;
  end;
end;

function VariantType(const V: Variant): Integer;
// CHECK WHAT TYPE OF VARIANT WE ARE DEALING WITH
begin
  Result := VarType(V);
end;

procedure LaunchProcess(const AppPath, Params: string; CreateNoWindow, WaitForProcessToEnd: Boolean);
var
  StartInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CmdLine: string;
  CreationFlags: DWORD;
begin
  FillChar(StartInfo, SizeOf(StartInfo), 0);
  StartInfo.cb := SizeOf(StartInfo);
  FillChar(ProcInfo, SizeOf(ProcInfo), 0);
  CmdLine := '"' + AppPath + '" ' + Params;
  CreationFlags := 0;
  if CreateNoWindow then
    CreationFlags := CreationFlags or CREATE_NO_WINDOW;
  if CreateProcess(nil, PChar(CmdLine), nil, nil, False, CreationFlags, nil, nil, StartInfo, ProcInfo) then
  begin
    if WaitForProcessToEnd then
    begin
      WaitForSingleObject(ProcInfo.hProcess, INFINITE);
    end;
    CloseHandle(ProcInfo.hProcess);
    CloseHandle(ProcInfo.hThread);
  end
  else
  begin
    RaiseLastOSError;
  end;
end;

end.
