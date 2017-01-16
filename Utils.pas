unit Utils;

interface

uses
  SysUtils, DateUtils, Windows, Messages, tlhelp32, Classes, ShlObj,
  ActiveX, ComObj, Variants, Clipbrd, Richedit,
  Forms, Vcl.Controls, Vcl.StdCtrls;

type
  StringOLE = string[255];

function DateTimeAssigned(Date: TDateTime): Boolean;
function DateTimeAsNull(): TDateTime;
function DateTimeIsNull(DateTime: TDateTime): Boolean;
procedure RaiseLabeledException(LabelText, ExceptionText: string);
function GetApplicationParameters(): string;
function GetForegroundWindowWithCaret(): HWnd;
function GetSelectedTextFromForegroundWindow(): string;
function GetSystemComputerName: string;
function GetSystemUserName: string;
function GetExeDateTime: TDateTime;
function GetExePath: string;
function GetManagerCaption(): string;
function GetOSVersion: string;
procedure KillAllAnotherInstancesOfThisExecutable();
function KeyboardLayoutName(): string;
procedure KeyDown(vkKey: Integer; DelayToWaitAfter: Cardinal = 0);
procedure KeyUp(vkKey: Integer; DelayToWaitAfter: Cardinal = 0);
procedure RedrawingLock(h: HWND);
procedure RedrawingUnlock(h: HWND);
procedure SetEnglishKeyboardLayout();
function StringToStringOLE(s: string): StringOLE;
function ShiftKeyIsPressed(): Boolean;
procedure ShowMessageMemo(const Msg: string);
function IntInArray(const X: Integer; const A: array of Integer): Boolean;
function StrInArray(const X: string; const A: array of string): Boolean;
function TextExtentSize(DC: HDC; const Text: string): TSize;
procedure TryChangeCharToDecimalSeparator(var C: Char);
function VarToDateTime_custom(Val: Variant; DefaultValue: TDateTime = Default(TDateTime)): TDateTime;
function VarToInt(X: Variant; DefaultValue: Integer = 0): Integer;
function UserIsAdministrator: Boolean;
function WindowsMyDocumentsDirectory(): string;
function WindowsTempDirectory: string;


implementation

function DateTimeAssigned(Date: TDateTime): Boolean;
begin
  Result := FloatToStr(DateTimeToUnix(Date)) <> '-2209161600';
end;

function DateTimeAsNull(): TDateTime;
begin
  Result := Default(TDateTime);
end;

function DateTimeIsNull(DateTime: TDateTime): Boolean;
begin
  Result := SameDateTime(DateTime, DateTimeAsNull);
end;

procedure RaiseLabeledException(LabelText, ExceptionText: string);
begin
  if LabelText <> '' then
    ExceptionText := LabelText + ' Exception: ' + ExceptionText;

  raise Exception.Create(ExceptionText);
end;

function GetExePath(): string;
begin
  SetLength(Result, MAX_PATH);
  SetLength(Result, GetModuleFileName(0, PChar(Result), MAX_PATH));
end;

function GetExeDateTime(): TDateTime;
begin
  FileAge(GetExePath(), Result);
end;

function GetManagerCaption(): string;
begin
  Result := 'Manager - версия ' + DateTimeToStr(GetExeDateTime()) + ' (XE)';
end;

procedure ShowMessageMemo(const Msg: string);
var
  F: TForm;
begin
  F := TForm.Create(nil);
  F.Position := poScreenCenter;

  with TMemo.Create(F) do
  begin
    Parent := F;
    Align := alClient;
    Text := Msg;
  end;

  F.ShowModal;
  FreeAndNil(F);
end;

procedure TryChangeCharToDecimalSeparator(var C: Char);
begin
  if CharInSet(C, ['.',',']) then
    C := FormatSettings.decimalseparator;
end;

function IntInArray(const X: Integer; const A: array of Integer): Boolean;
var
  tmpX: Integer;
begin
  for tmpX in A do
    if tmpX = X then
      Exit(True);

  Result := False;
end;

function StrInArray(const X: string; const A: array of string): Boolean;
var
  tmpX: string;
begin
  for tmpX in A do
    if tmpX = X then
      Exit(True);

  Result := False;
end;

function GetOSVersion: string;

  function GetWMIObject(const objectName: string): IDispatch;
  var
    chEaten: Integer;
    BindCtx: IBindCtx;
    Moniker: IMoniker;
  begin
    OleCheck(CreateBindCtx(0, BindCtx));
    OleCheck(MkParseDisplayName(BindCtx, PWideChar(objectName), chEaten, Moniker));
    OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));
  end;

  function VarToString(const Value: OleVariant): string;
  begin
    if VarIsStr(Value) then
      Result := Trim(Value)
    else
      Result := '';
  end;

  function FullVersionString(const Item: OleVariant): string;
  var
    Caption, ServicePack, Version, Architecture: string;
  begin
    Caption := VarToString(Item.Caption);
    ServicePack := VarToString(Item.CSDVersion);
    Version := VarToString(Item.Version);
    Architecture := VarToString(Item.OSArchitecture);
    Result := Caption;

    if ServicePack <> '' then
      Result := Result + ' ' + ServicePack;

    Result := Result + ', ' + Version + ', ' + Architecture;
  end;

var
  objWMIService: OleVariant;
  colItems: OleVariant;
  Item: OleVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
begin
  Result := 'Unknown';

  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems := objWMIService.ExecQuery('SELECT Caption, CSDVersion, Version, OSArchitecture FROM Win32_OperatingSystem', 'WQL', 0);
    oEnum := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, Item, iValue) = 0 then
    begin
      Result := FullVersionString(Item);
      Exit();
    end;
  except
    Result := TOSVersion.ToString();
  end;
end;

function TrySetPrivilege(SetOn: Boolean): Boolean;
var
  hToken: THandle;
  tkp: TOKEN_PRIVILEGES;
  ReturnLength: Cardinal;
begin
  Result := False;

  if not OpenProcessToken( GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES
  or TOKEN_QUERY, hToken )
  then
    Exit();

  tkp.PrivilegeCount:= 1;
  if SetOn then
    tkp.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED
  else
    tkp.Privileges[0].Attributes := 0;

  if LookupPrivilegeValue( nil, 'SeDebugPrivilege', tkp.Privileges[0].Luid ) then
  begin
    AdjustTokenPrivileges(hToken, False, tkp, SizeOf(tkp), tkp, ReturnLength);
    Result := (GetLastError() = ERROR_SUCCESS);
  end;

  CloseHandle(hToken);
end;

function ProcessTerminate(dwPID:Cardinal): Boolean;
var
  hProcess: THandle;
begin
  Result := False;
  TrySetPrivilege(True);
  hProcess := OpenProcess(PROCESS_TERMINATE, FALSE, dwPID);
  if hProcess <> 0 then
  begin
    Result := TerminateProcess(hProcess, Cardinal(-1));
    CloseHandle(hProcess);
  end;
  TrySetPrivilege(False);
end;

procedure KillAllAnotherInstancesOfThisExecutable();
var
  Snapshot: THandle;
  ProcessEntry: TProcessEntry32;
  ThisExeName: string;
  ProcessName: string;
begin
  ThisExeName := ExtractFileName(GetExePath());

  Snapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  ProcessEntry.dwSize := SizeOf(ProcessEntry);

  if Process32First(Snapshot, ProcessEntry) then
  repeat
    if ProcessEntry.th32ProcessID = GetCurrentProcessId then
      Continue;

    ProcessName := ExtractFileName(ProcessEntry.szExeFile);
    if SameText(ThisExeName, ProcessName) then
      ProcessTerminate(ProcessEntry.th32ProcessID);
  until not Process32Next(Snapshot, ProcessEntry);

  CloseHandle(Snapshot);
end;

function KeyboardLayoutName(): string;
var
  LayoutName: array [0 .. KL_NAMELENGTH + 1] of Char;
  LangName: array [0 .. 1024] of Char;
begin
  Result := '??';

  if not Windows.GetKeyboardLayoutName(@LayoutName) then
    Exit;

  if GetLocaleInfo(
    StrToInt('$' + StrPas(LayoutName)),
    LOCALE_SABBREVLANGNAME,
    @LangName,
    SizeOf(LangName) - 1
  ) = 0 then
    Exit;

  Result := StrPas(LangName);
end;

function GetApplicationParameters(): string;
var
  i:integer;
begin
  Result := '';
  if ParamCount > 0 then
    for i := 1 to ParamCount do
      Result := Result + ParamStr(i) + ' ';
end;

function GetForegroundWindowWithCaret(): HWnd;
var
  ThreadId, ProcessId: Cardinal;
  GUIThreadInfo: TGUIThreadInfo;
begin
  Result := GetForegroundWindow();
  if Result = 0 then
    Exit();

  ThreadId := GetWindowThreadProcessId(Result, ProcessId);

  GUIThreadInfo.cbSize := SizeOf(TGUIThreadInfo);
  if not GetGUIThreadInfo(ThreadId, GUIThreadInfo) then
    Exit();

  Result := GUIThreadInfo.hwndCaret;
end;

procedure WaitForClipboardIsOpenable(TimeoutInMs: Cardinal = 2000);
const
  TimeStepInMs = 100;
var
  WhenWaitingMustBeFinished: Cardinal;
begin
  WhenWaitingMustBeFinished := GetTickCount() + TimeoutInMs;

  repeat
    if GetOpenClipboardWindow() = 0 then
      Exit();
    Sleep(TimeStepInMs);
  until (GetTickCount() > WhenWaitingMustBeFinished);

  raise Exception.Create('Clipboard waiting time out.');
end;

function GetSelectedTextFromForegroundWindow(): string;
var
  OldTextVal: string;
  Wnd: HWnd;
begin
  Result := '';
  try
    WaitForClipboardIsOpenable();
    Clipboard.Open();
    OldTextVal := Clipboard.AsText;
    Clipboard.AsText := '';
    Clipboard.Close();

    WaitForClipboardIsOpenable();
    Wnd := GetForegroundWindowWithCaret();
    OpenClipboard(Wnd);
    SendMessage(Wnd, WM_COPY, 0, 0);
    CloseClipboard();

    WaitForClipboardIsOpenable();
    Clipboard.Open();
    Result := Clipboard.AsText;
    Clipboard.AsText := OldTextVal;
  finally
    Clipboard.Close();
  end;
end;

function GetSystemComputerName: string;
var
  ComputerName: array[0..255] of Char;
  ComputerNameSize: DWORD;
begin
  ComputerNameSize := 255;
  if Windows.GetComputerName(@ComputerName, ComputerNameSize) then
    Result := string(ComputerName)
  else
    Result := '';
end;

function GetSystemUserName: string;
var
  UserName: array[0..255] of Char;
  UserNameSize: DWORD;
begin
  UserNameSize := 255;
  if Windows.GetUserName(@UserName, UserNameSize) then
    Result := string(UserName)
  else
    Result := '';
end;

procedure RedrawingLock(h: HWND);
begin
  SendMessage(h, WM_SETREDRAW, 0, 0);
end;

procedure RedrawingUnlock(h: HWND);
begin
  SendMessage(h, WM_SETREDRAW, 1, 0);
end;

procedure SetEnglishKeyboardLayout();
begin
  LoadKeyboardLayout('00000409', KLF_ACTIVATE);
end;

function StringToStringOLE(s: string): StringOLE;
begin
  s := s.Substring(0, 256);
  Result := StringOLE(s);
end;

function ShiftKeyIsPressed(): Boolean;
begin
  Result := Windows.GetKeyState(VK_SHIFT) < 0;
end;

function TextExtentSize(DC: HDC; const Text: string): TSize;
var
  Rect: TRect;
begin
  FillChar(Rect, SizeOf(TRect), 0);
  DrawTextEx(DC, PChar(Text), Length(Text), Rect, DT_CALCRECT, nil);
  Result.cx := Rect.Width;
  Result.cy := Rect.Height;
end;

function VarToDateTime_custom(Val: Variant; DefaultValue: TDateTime): TDateTime;
begin
  if Val = Null then
    Result := DefaultValue
  else
    Result := Variants.VarToDateTime(Val);
end;

function VarToInt(X: Variant; DefaultValue: Integer): Integer;
begin
  if not TryStrToInt(VarToStr(X), Result) then
    Result := DefaultValue;
end;

function UserIsAdministrator: Boolean;
begin
  Result := (GetSystemComputerName='VM-AVK')
    or (GetSystemComputerName='WS-IT-A15')
    or (GetSystemComputerName='VM-MIHA')
    or (GetSystemComputerName='VM-MANAGER')
    or (GetSystemComputerName='PRGM4');
end;

function WindowsTempDirectory: string;
begin
  Result := GetEnvironmentVariable('TEMP');
end;

function GetSpecialPath(CSIDL: word): string;
var
  s: string;
begin
  SetLength(s, MAX_PATH);
  if not SHGetSpecialFolderPath(0, PChar(s), CSIDL, true) then
    s := '';
end;

function WindowsMyDocumentsDirectory(): string;
begin
  Result := GetSpecialPath(5);
end;

function IsKeyExtended(Key: Word): Byte;
begin
  {расширенные клавиши - правый ALT и СTRL на основном разделе клавиатуры; INS,
   DEL, HOME, END, PAGE UP, PAGE DOWN, клавиши курсора слева от цифровой
   клавиатуры, наклонная черта вправо (/) и клавиша ENTER на цифровой клавиатуре}
  if IntInArray(Key, [37, 38, 39, 40, 45, 36, 35, 33, 34, 191]) then
    Result := 1
  else
    Result := 0;
end;

procedure KeyDown(vkKey: Integer; DelayToWaitAfter: Cardinal = 0);
begin
  Windows.keybd_event(vkKey, Windows.MapVirtualKey(vkKey, 0),
    0 + IsKeyExtended(vkKey), 0);
  sleep(DelayToWaitAfter);
end;

procedure KeyUp(vkKey: Integer; DelayToWaitAfter: Cardinal = 0);
begin
  Windows.keybd_event(vkKey, Windows.MapVirtualKey(vkKey, 0),
    2 + IsKeyExtended(vkKey), 0);
  sleep(DelayToWaitAfter);
end;

end.

