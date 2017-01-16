unit uHashTable;

interface

uses
  Variants;

type
  TVariantArray = array of Variant;

  THashTable = class
  private
    FKeys: TVariantArray;
    FValues: TVariantArray;
    function IndexOfKey(Key: Variant): Integer;
    function Get(Key: Variant): Variant;
    procedure Put(Key: Variant; const Value: Variant);
    procedure AddNewPair(const Key, Value: Variant);
  public
    destructor Destroy(); override;
    function Clone(): THashTable;
    procedure Clear();
    function Count(): Integer;
    property Value[Key: Variant]: Variant read Get write Put; default;
    property Keys: TVariantArray read FKeys;
    property Values: TVariantArray read FValues;
  end;

  THashTableArray = array of THashTable;
  THashTableList = class
  private
    FItems: THashTableArray;
    function GetCount: Integer;
    function GetFirst: THashTable;
    function GetLast: THashTable;
  public
    destructor Destroy(); override;
    procedure Clear();
    procedure Add(HashTable: THashTable);
    function Clone(): THashTableList;
    property Items: THashTableArray read FItems;
    property Count: Integer read GetCount;
    property First: THashTable read GetFirst;
    property Last: THashTable read GetLast;
  end;

function Null: Variant;  
  
implementation

function Null: Variant;
begin
  Result := Variants.Null;
end;

destructor THashTable.Destroy();
begin
  Clear();
  inherited;
end;

function THashTable.Get(Key: Variant): Variant;
var
  I: Integer;
begin
  I := IndexOfKey(Key);

  if I in [Low(FKeys)..High(FKeys)] then
    Result := FValues[I]
  else
    Result := Null;
end;

function THashTable.IndexOfKey(Key: Variant): Integer;
begin
  for Result := Low(FKeys) to High(FKeys) do
    if FKeys[Result] = Key then
      Exit;

  Result := Low(FKeys) - 1;
end;

procedure THashTable.Put(Key: Variant; const Value: Variant);
var
  I: Integer;
begin
  I := IndexOfKey(Key);

  if (I >= Low(FKeys)) and (I <= High(FKeys)) then
    FValues[I] := Value
  else
    AddNewPair(Key, Value)
end;

procedure THashTable.AddNewPair(const Key, Value: Variant);
begin
  SetLength(FKeys, Length(FKeys) + 1);
  FKeys[High(FKeys)] := Key;
  SetLength(FValues, Length(FValues) + 1);
  FValues[High(FValues)] := Value;
end;

function THashTable.Count(): Integer;
begin
  Result := Length(FKeys);
end;

procedure THashTable.Clear();
begin
  SetLength(FKeys, 0);
  SetLength(FValues, 0);
end;

function THashTable.Clone(): THashTable;
var
  Key: Variant;
begin
  Result := THashTable.Create;

  for Key in Self.Keys do
    Result[Key] := Self[Key];
end;

{ THashTableList }

procedure THashTableList.Add(HashTable: THashTable);
begin
  SetLength(FItems, Length(FItems) + 1);
  FItems[High(FItems)] := HashTable;
end;

procedure THashTableList.Clear;
var
  tmpRecord: THashTable;
begin
  for tmpRecord in FItems do
    tmpRecord.Free;

  SetLength(FItems, 0);
end;

function THashTableList.Clone: THashTableList;
var
  tmpHashTable: THashTable;
begin
  Result := THashTableList.Create();
  for tmpHashTable in Self.Items do
    Result.Add(tmpHashTable.Clone());
end;

destructor THashTableList.Destroy;
begin
  Clear();

  inherited;
end;

function THashTableList.GetCount: Integer;
begin
  Result := Length(FItems);
end;

function THashTableList.GetFirst: THashTable;
begin
  Result := FItems[Low(FItems)];
end;

function THashTableList.GetLast: THashTable;
begin
  Result := FItems[High(FItems)];
end;

end.
