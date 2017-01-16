unit uControlHelpers;

interface

uses
  Windows, SysUtils, System.Classes, Vcl.Dialogs, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  VCL.DBCtrls, Data.DB, System.Variants, Vcl.ValEdit;

type
  TWinControlHelper = class helper for TWinControl
  public type
    TChildControlProc = reference to procedure(Control: TControl);
  public
    function FindFirstControl<T: TControl>(): T;
    procedure ForEachChildDo<T: TControl>(Proc: TChildControlProc);
  end;

  TCategoryPanelGroupHelper = class helper for TCategoryPanelGroup
  public
    procedure AdjustPanelHeights();
    procedure DeleteAllPanels;
    function Panel(Caption: string): TCategoryPanel;
  end;

  TCategoryPanelHelper = class helper for TCategoryPanel
  public
    procedure AdjustHeight();
    function FindFirstControl<T: TControl>(): T;
    function Surface: TCategoryPanelSurface;
  end;

  TStatusBarHelper = class helper for TStatusBar
  public
    procedure AdjustPanelsSize();
    function PanelFromUnderMouse(): TStatusPanel;
    function PanelIsUnderMouse(Index: Integer): Boolean;
  end;

  TSplitterHelper = class helper for TSplitter
  private
    procedure SplitterHelperCanResize(Sender: TObject; var NewSize: Integer;
      var Accept: Boolean);
  public
    procedure ReplaceCanResizeEventToPreventFullMinimizing();
  end;

  TDBLookupComboBoxHelper = class helper for TDBLookupComboBox
  public
    procedure InitWithDataset(ds: TDataset);
  end;

  TValueListEditorHelper = class helper for TValueListEditor
    procedure TryDeleteByKey(Key: string);
  end;

implementation

uses
  Utils;

{ TWinControlHelper }

function TWinControlHelper.FindFirstControl<T>: T;
var
  I: Integer;
begin
  for I := 0 to Self.ControlCount - 1 do
    if Self.Controls[I] is T then
      Exit(Self.Controls[I] as T);

  Result := nil;
end;

procedure TWinControlHelper.ForEachChildDo<T>(Proc: TChildControlProc);
var
  I: Integer;
begin
  if not Assigned(Proc) then
    Exit();

  for I := 0 to Self.ControlCount - 1 do
    if Self.Controls[I] is T then
      Proc(Self.Controls[I]);
end;

{ TCategoryPanelHelper }

procedure TCategoryPanelHelper.AdjustHeight;
var
  I: Integer;
  BottomOfMostBottomControl: Integer;
  BottomOfCurrentControl: Integer;
  tmpSurface: TCategoryPanelSurface;
begin
  tmpSurface := Self.Surface;
  if not Assigned(tmpSurface) then
    Exit();

  BottomOfMostBottomControl := 0;
  for I := 0 to tmpSurface.ControlCount - 1 do
  begin
    BottomOfCurrentControl := tmpSurface.Controls[I].BoundsRect.Bottom;
    if BottomOfCurrentControl > BottomOfMostBottomControl then
      BottomOfMostBottomControl := BottomOfCurrentControl;
  end;

  Self.ClientHeight := BottomOfMostBottomControl + 1;
end;

function TCategoryPanelHelper.FindFirstControl<T>: T;
begin
  if Surface = nil then
    Result := nil
  else
    Result := Surface.FindFirstControl<T>();
end;

function TCategoryPanelHelper.Surface: TCategoryPanelSurface;
begin
  Result := inherited FindFirstControl<TCategoryPanelSurface>();
end;

{ TStatusBarHelper }

procedure TStatusBarHelper.AdjustPanelsSize();
const
  WidthGap = 6;
var
  I: Integer;
  Panel: TStatusPanel;
begin
  Canvas.Font := Self.Font;
  for I := 0 to Panels.Count - 1 do
  begin
    Panel := Panels.Items[I];
    if Panel.Text <> '' then
      Panel.Width := TextExtentSize(Canvas.Handle, Panel.Text).cx + WidthGap;
  end;
end;

function TStatusBarHelper.PanelFromUnderMouse: TStatusPanel;
var
  CurPos: TPoint;
  PanelRight: Integer;
  IndexOfLastPanel: Integer;
  I: Integer;
begin
  Result := nil;
  if not MouseInClient
  or not GetCursorPos(CurPos)
  then
    Exit();

  CurPos := ScreenToClient(CurPos);

  PanelRight := 0;
  IndexOfLastPanel := Panels.Count - 1;
  for I := 0 to IndexOfLastPanel - 1 do
  begin
    PanelRight := PanelRight + Panels[I].Width;

    if PanelRight > CurPos.X then
      Exit(Panels[I]);
  end;

  Result := Panels[IndexOfLastPanel];
end;

function TStatusBarHelper.PanelIsUnderMouse(Index: Integer): Boolean;
var
  Panel: TStatusPanel;
begin
  Result := False;

  Panel := PanelFromUnderMouse();
  if Assigned(Panel) then
    if Panel.Index = Index then
      Result := True;
end;

{ TCategoryPanelGroupHelper }

procedure TCategoryPanelGroupHelper.AdjustPanelHeights;
var
  PPanel: Pointer;
begin
  for PPanel in Self.Panels do
    TCategoryPanel(PPanel).AdjustHeight();
end;

function TCategoryPanelGroupHelper.Panel(Caption: string): TCategoryPanel;
var
  PanelPointer: Pointer;
begin
  Result := nil;
  for PanelPointer in Self.Panels do
    if TCategoryPanel(PanelPointer).Caption = Caption then
      Exit(PanelPointer);
end;

procedure TCategoryPanelGroupHelper.DeleteAllPanels();
var
  PanelPointer: Pointer;
begin
  for PanelPointer in Self.Panels do
    TCategoryPanel(PanelPointer).Free;
  Self.Panels.Clear();
end;

{ TSplitterHelper }

procedure TSplitterHelper.ReplaceCanResizeEventToPreventFullMinimizing;
begin
  Self.OnCanResize := SplitterHelperCanResize;
end;

procedure TSplitterHelper.SplitterHelperCanResize(Sender: TObject; var NewSize: Integer;
  var Accept: Boolean);
begin
  if NewSize <= MinSize then
    Accept := False;
end;

{ TDBLookupComboBoxHelper }

procedure TDBLookupComboBoxHelper.InitWithDataset(ds: TDataset);
begin
  if not Assigned(ListSource) then
    ListSource := TDatasource.Create(Self);

  if Assigned(ListSource.DataSet) then
  begin
    ListSource.DataSet.Cancel();
    ListSource.DataSet.Free;
  end;

  ListSource.DataSet := ds;
  ds.Open();
  ds.Last();
  ds.First();
end;

{ TValueListEditorHelper }

procedure TValueListEditorHelper.TryDeleteByKey(Key: string);
var
  IndexOfRowToDelete: Integer;
begin
  IndexOfRowToDelete := Strings.IndexOfName(VarToStr(Key));
  if IndexOfRowToDelete > -1 then
    Strings.Delete(IndexOfRowToDelete);
end;

end.
