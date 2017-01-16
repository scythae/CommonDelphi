unit uDynamicForm;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, Vcl.Forms, Vcl.Controls,
  Vcl.StdCtrls, Vcl.Dialogs, Vcl.ComCtrls, Vcl.DBCtrls, System.Generics.Collections,
  uIBDB;

type
  TControlAligner = class;

  TDynamicFormBuilderBase = class
  private
    class var
      FForm: TForm;
      FAligner: TControlAligner;
  private
    class procedure CreateControl(ControlClass: TControlClass; var Reference);
    class procedure SetAlign(Control: TControl);
    class procedure RaiseIfBuilding();
    class procedure RaiseIfNotBuilding();
  protected
    class procedure DoException(CustomMessage: string);
  public
    class procedure BuildForm(FormOwner: TComponent); virtual;
    class function GetForm(): TForm; virtual;
    class procedure OnExceptionDuringBuild();
  end;

  TControlAligner = class
  private
    FForm: TForm;
  public
    constructor Create(Form: TForm);
    procedure AlignControl(Control: TControl);
    procedure AlignForm();
  end;

  TDynamicFormBuilder = class(TDynamicFormBuilderBase)
  private
    class var
      FDBController: TIBDBConnection;
  private
    class function GetDBController: TIBDBConnection; static;
  public
    class procedure BuildForm(FormOwner: TComponent); override;
    class function GetForm(): TForm; override;
    class function CreateLabel(Text: string = ''): TLabel;
    class function CreateMemo(): TMemo;
    class function CreateButton(): TButton; overload;
    class function CreateButton(Caption: string; ModalResult: Integer = mrNone): TButton; overload;
    class function CreateDBLookupComboBox(): TDBLookupComboBox; overload;
    class function CreateDBLookupComboBox(SQLQuery: string;
      KeyField, ListField: string): TDBLookupComboBox; overload;
    class function CreateDBLookupComboBox(Dataset: TDataset; KeyField,
      ListField: string): TDBLookupComboBox; overload;
    class property DBController: TIBDBConnection read GetDBController write FDBController;
  end;

  TEventContainer = class(TComponent)
  private type
    TNotifyEventProc = TProc<TObject>;
    TCloseEventProc = reference to procedure(Sender: TObject; var Action: TCloseAction);
  private
    EventProc: Pointer;
  public
    constructor Create(Owner: TComponent; const Proc: Pointer); reintroduce;
    destructor Destroy(); override;
    procedure NotifyEvent(Sender: TObject);
    procedure CloseEvent(Sender: TObject; var Action: TCloseAction);
  end;

implementation

uses
  uControlHelpers;

class procedure TDynamicFormBuilderBase.BuildForm(FormOwner: TComponent);
begin
  RaiseIfBuilding();
  FForm := TForm.Create(FormOwner);
  FAligner := TControlAligner.Create(FForm);
end;

class function TDynamicFormBuilderBase.GetForm(): TForm;
begin
  RaiseIfNotBuilding();
  FAligner.AlignForm();
  Result := FForm;

  FreeAndNil(FAligner);
  FForm := nil;
end;

class procedure TDynamicFormBuilderBase.CreateControl(ControlClass: TControlClass; var Reference);
var
  Control: TControl;
begin
  RaiseIfNotBuilding();

  Control := ControlClass.Create(FForm);
  try
    Control.Parent := FForm;
    SetAlign(Control);
  except
    FreeAndNil(Control);
  end;

  TControl(Reference) := Control;
end;

class procedure TDynamicFormBuilderBase.SetAlign(Control: TControl);
begin
  FAligner.AlignControl(Control);
end;

class procedure TDynamicFormBuilderBase.RaiseIfBuilding();
begin
  if Assigned(FForm) then
    DoException('Form building is already started.');
end;

class procedure TDynamicFormBuilderBase.RaiseIfNotBuilding();
begin
  if not Assigned(FForm) then
    DoException('Form building is not started.');
end;

class procedure TDynamicFormBuilderBase.DoException(CustomMessage: string);
begin
  raise Exception.Create(Self.ClassName + ': ' + CustomMessage);
end;

class procedure TDynamicFormBuilderBase.OnExceptionDuringBuild();
begin
  GetForm().Free();
end;

{ TControlAligner }

constructor TControlAligner.Create(Form: TForm);
begin
  inherited Create();
  FForm := Form;
end;

procedure TControlAligner.AlignControl(Control: TControl);
begin
  Control.Align := alTop;
  Control.Top := FForm.ClientHeight;
end;

procedure TControlAligner.AlignForm;
begin
  FForm.AutoSize := True;
  FForm.Position := poOwnerFormCenter;
end;

{TDynamicFormBuilder}

class procedure TDynamicFormBuilder.BuildForm(FormOwner: TComponent);
begin
  inherited BuildForm(FormOwner);
  FDBController := nil;
end;

class function TDynamicFormBuilder.GetForm: TForm;
begin
  Result := inherited GetForm();
  FDBController := nil;
end;

class function TDynamicFormBuilder.GetDBController: TIBDBConnection;
begin
  Result := FDBController;
  if not Assigned(Result) then
    DoException('DBController is not assigned.');
end;

class function TDynamicFormBuilder.CreateButton(): TButton;
begin
  CreateControl(TButton, Result);
end;

class function TDynamicFormBuilder.CreateButton(Caption: string; ModalResult: Integer): TButton;
begin
  Result := CreateButton();
  Result.Caption := Caption;
  Result.ModalResult := ModalResult;
end;

class function TDynamicFormBuilder.CreateDBLookupComboBox(): TDBLookupComboBox;
begin
  CreateControl(TDBLookupComboBox, Result);
end;

class function TDynamicFormBuilder.CreateDBLookupComboBox(SQLQuery: string;
  KeyField, ListField: string): TDBLookupComboBox;
begin
  Result := CreateDBLookupComboBox();

  Result.InitWithDataset(
    DBController.CreateReadQuery(FForm, SQLQuery)
  );
  Result.KeyField := KeyField;
  Result.ListField := ListField;
end;

class function TDynamicFormBuilder.CreateDBLookupComboBox(Dataset: TDataset;
  KeyField, ListField: string): TDBLookupComboBox;
begin
  Result := CreateDBLookupComboBox();
  Result.InitWithDataset(Dataset);
  Result.KeyField := KeyField;
  Result.ListField := ListField;
end;

class function TDynamicFormBuilder.CreateLabel(Text: string): TLabel;
begin
  CreateControl(TLabel, Result);
  Result.Caption := Text;
end;

class function TDynamicFormBuilder.CreateMemo(): TMemo;
begin
  CreateControl(TMemo, Result);
end;

{ TEventContainer }

constructor TEventContainer.Create(Owner: TComponent;
  const Proc: Pointer);
begin
  inherited Create(Owner);
  EventProc := Proc;
  IInterface(EventProc)._AddRef();
end;

destructor TEventContainer.Destroy;
begin
  IInterface(EventProc)._Release();
  inherited;
end;

procedure TEventContainer.NotifyEvent(Sender: TObject);
begin
  TNotifyEventProc(EventProc)(Sender);
end;

procedure TEventContainer.CloseEvent(Sender: TObject; var Action: TCloseAction);
begin
  TCloseEventProc(EventProc)(Sender, Action);
end;

end.

