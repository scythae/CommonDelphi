unit uIBDB;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, IBX.IBDatabase, IBX.IB, Data.DB,
  IBX.IBCustomDataSet, IBX.IBQuery, DBLogdlg, uHashTable, contnrs;

type
  THashTable = uHashTable.THashTable;
  TRecord = THashTable;
  THashTableList = uHashTable.THashTableList;
  TRecordList = THashTableList;
  TDataSet = Data.DB.TDataSet;
  TIBQuery = IBX.IBQuery.TIBQuery;
  EIBInterBaseError = IBX.IB.EIBInterBaseError;
  TField = Data.DB.TField;
  TBlobField = Data.DB.TBlobField;
  function ftBlob(): TFieldType;

type
  TIBDBConnection = class
  private
    type
      TVariantArray = array of Variant;
  private
    Database: TIBDatabase;
    FReadTransaction: TIBTransaction;
    FUserWantsWriteModeForNextSelect: Boolean;
    function CreateReadTransaction(): TIBTransaction;
    function CreateRecordListFromQuery(Query: TIBQuery): TRecordList;
    function GetReadTransaction: TIBTransaction;
    function CreateWriteTransaction(AOwner: TComponent): TIBTransaction;
    procedure CheckQuery(var Query: TIBQuery);
    property ReadTransaction: TIBTransaction read GetReadTransaction;
    function CreateQueryForNextSelect(SqlQueryText: string): TIBQuery;
    procedure InitializeQueryWithParameters(Query: TIBQuery; Params: TVariantArray);
  public
    constructor Create(DatabaseName, UserName, Password: string;
      Role: string = '');
    destructor Destroy(); override;
    procedure SetWriteModeForNextSelect();
    function SelectField(SqlQueryText: string; Params: TVariantArray = []): Variant;
    function SelectRecord(SqlQueryText: string; Params: TVariantArray = []): TRecord;
    function SelectRecordList(SqlQueryText: string; Params: TVariantArray = []): TRecordList;
    procedure ExecuteWriteQuery(SqlQueryText: string; Params: TVariantArray = []);

    function CreateReadQuery(SqlQueryText: string; Params: TVariantArray = []): TIBQuery; overload;
    function CreateReadQuery(AOwner: TComponent;
      SqlQueryText: string; Params: TVariantArray = []): TIBQuery; overload;
    function CreateWriteQuery(SqlQueryText: string; Params: TVariantArray = []): TIBQuery; overload;
    function CreateWriteQuery(AOwner: TComponent;
      SqlQueryText: string; Params: TVariantArray = []): TIBQuery; overload;

    function CreateDatasource(AOwner: TComponent = nil): TDataSource;
    function HasAccess(AccessId: Integer; MarkId: Integer): Boolean;
    function GetServerDateTime(): TDateTime;
  end;

implementation

function ftBlob(): TFieldType;
begin
  Result := TFieldType.ftBlob;
end;

constructor TIBDBConnection.Create(DatabaseName, UserName, Password: string;
  Role: string);
begin
  inherited Create();

  Database := TIBDatabase.Create(nil);
  Database.DatabaseName := DatabaseName;
  Database.LoginPrompt := (Password = '');
  if not Database.LoginPrompt then
  begin
    Database.Params.Values['user_name'] := UserName;
    Database.Params.Values['password'] := Password;
    Database.Params.Values['sql_role_name'] := Role;
  end;

  try
    Database.Connected := True;
  except
    FreeAndNil(Database);
  end;
end;

destructor TIBDBConnection.Destroy;
begin
  FreeAndNil(FReadTransaction);
  FreeAndNil(Database);

  inherited Destroy();
end;

function TIBDBConnection.GetReadTransaction: TIBTransaction;
begin
  if not Assigned(FReadTransaction) then
    FReadTransaction := CreateReadTransaction();

  Result := FReadTransaction;
end;

function TIBDBConnection.GetServerDateTime(): TDateTime;
begin
  try
    Result := SelectField('Select current_timestamp ts from rdb$database');
  except
    Result := 0;
  end;
end;

function TIBDBConnection.HasAccess(AccessId, MarkId: Integer): Boolean;
var
  MarkId_Var: Variant;
begin
  if MarkId = 0 then
    MarkId_Var := Null
  else
    MarkId_Var := MarkId;

  Result := SelectField(
    'select result from check_user_rights(null, :access_id, :mark_id)',
    [AccessId, MarkId_Var]
  );
end;

function TIBDBConnection.SelectField(SqlQueryText:
  string; Params: TVariantArray): Variant;
var
  SqlRecord: TRecord;
begin
  SqlRecord := SelectRecord(SqlQueryText, Params);

  try
    if SqlRecord.Count() = 0 then
      Result := Null
    else if SqlRecord.Count() = 1 then
      Result := SqlRecord.Values[0]
    else
      raise Exception.Create('Too many fields have been selected.');
  finally
    FreeAndNil(SqlRecord);
  end;
end;

function TIBDBConnection.SelectRecord(SqlQueryText: string; Params: TVariantArray): TRecord;
var
  RecordList: TRecordList;
begin
  RecordList := SelectRecordList(SqlQueryText, Params);

  if RecordList.Count > 0 then
    Result := TRecord(RecordList.Last).Clone()
  else
    Result := TRecord.Create();

  FreeAndNil(RecordList);
end;

function TIBDBConnection.SelectRecordList(SqlQueryText: string; Params: TVariantArray): TRecordList;
var
  Query: TIBQuery;
begin
  Query := CreateQueryForNextSelect(SqlQueryText);

  try
    InitializeQueryWithParameters(Query, Params);
    Query.Open;
  except
    FreeAndNil(Query);
    raise;
  end;

  Result := CreateRecordListFromQuery(Query);

  FreeAndNil(Query);
end;

function TIBDBConnection.CreateQueryForNextSelect(SqlQueryText: string): TIBQuery;
begin
  if FUserWantsWriteModeForNextSelect then
  begin
    Result := CreateWriteQuery(SqlQueryText);
    FUserWantsWriteModeForNextSelect := False;
  end
  else
    Result := CreateReadQuery(SqlQueryText);
end;

procedure TIBDBConnection.SetWriteModeForNextSelect();
begin
  FUserWantsWriteModeForNextSelect := True;
end;

function TIBDBConnection.CreateReadQuery(SqlQueryText: string; Params: TVariantArray): TIBQuery;
begin
  Result :=  CreateReadQuery(nil, SqlQueryText, Params);
end;

function TIBDBConnection.CreateReadQuery(AOwner: TComponent;
  SqlQueryText: string; Params: TVariantArray): TIBQuery;
begin
  Result:= TIBQuery.Create(AOwner);
  Result.Transaction := ReadTransaction;
  Result.Database := Database;
  Result.ParamCheck := True;
  Result.SQL.Text := SqlQueryText;

  CheckQuery(Result);
  InitializeQueryWithParameters(Result, Params);
end;

procedure TIBDBConnection.CheckQuery(var Query: TIBQuery);
begin
  try
    Query.Prepare();
  except
    FreeAndNil(Query);
    raise;
  end;
end;

function TIBDBConnection.CreateRecordListFromQuery(Query: TIBQuery): TRecordList;
var
  Field: TField;
  SqlRecord: TRecord;
begin
  Result := TRecordList.Create();

  Query.First();
  while not Query.Eof do
  begin
    SqlRecord := TRecord.Create();
    for Field in Query.Fields do
      SqlRecord.Value[Field.FieldName] := Field.Value;
    Result.Add(SqlRecord);

    Query.Next();
  end;
end;

function TIBDBConnection.CreateReadTransaction(): TIBTransaction;
begin
  Result := TIBTransaction.Create(nil);
  Result.DefaultDatabase := Database;
  Result.Params.Add('read');
  Result.Params.Add('read_committed');
  Result.Params.Add('rec_version');
  Result.Params.Add('nowait');
end;

procedure TIBDBConnection.InitializeQueryWithParameters(Query: TIBQuery;
  Params: TVariantArray);
var
  I: Integer;
begin
  if Length(Params) = 0 then
    Exit();

  if Query.ParamCount <> Length(Params) then
    raise Exception.Create('Incorrect number of parameters.');

  for I := 0 to Query.ParamCount - 1 do
    Query.Params.Items[I].Value := Params[I];
end;

procedure TIBDBConnection.ExecuteWriteQuery(SqlQueryText: string; Params: TVariantArray);
var
  Query: TIBQuery;
begin
  Query := CreateWriteQuery(SqlQueryText);
  try
    InitializeQueryWithParameters(Query, Params);
    Query.ExecSQL();
  finally
    FreeAndNil(Query);
  end;
end;

function TIBDBConnection.CreateWriteQuery(SqlQueryText: string; Params: TVariantArray): TIBQuery;
begin
  Result := CreateWriteQuery(nil, SqlQueryText, Params);
end;

function TIBDBConnection.CreateWriteQuery(AOwner: TComponent;
  SqlQueryText: string; Params: TVariantArray): TIBQuery;
begin
  Result:= TIBQuery.Create(AOwner);
  Result.Transaction := CreateWriteTransaction(Result);
  Result.Database := Database;
  Result.ParamCheck := True;
  Result.SQL.Text := SqlQueryText;

  CheckQuery(Result);
  InitializeQueryWithParameters(Result, Params);
end;

function TIBDBConnection.CreateWriteTransaction(AOwner: TComponent): TIBTransaction;
begin
  Result := TIBTransaction.Create(AOwner);
  Result.DefaultDatabase := Database;
  Result.Params.Add('write');
  Result.Params.Add('read_committed');
  Result.Params.Add('rec_version');
  Result.Params.Add('nowait');
end;

function TIBDBConnection.CreateDatasource(AOwner: TComponent): TDataSource;
begin
  Result := TDataSource.Create(AOwner);
end;



end.

