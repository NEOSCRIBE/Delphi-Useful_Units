{ ************************************************************

  --------------------Модуль uArray---------------------------
  -----------------------ver.1.0------------------------------

  Разработчик: scribe
  Создан: 12.10.2015
  Модифицирован: 15.10.2015
  Связь: justscribe@yahoo.com

  Описание:
  2 класса для работы с динамическими массивами
  строкового типа.
  Создано специально для Юнита uExcel.

  ************************************************************ }
{$DEFINE Debug}
unit uArray;

interface

uses
  SysUtils, Classes;

type
  TDoubleStrArray = class
  strict private
    FHeader: Pointer;
    FDArray: array of array of string;
    FWidth: integer;
    FHeight: integer;
    FOnError: TNotifyEvent;
    FDefValue: string;
    FLastError: string;
    function dsaGetOnError: TNotifyEvent;
    procedure dsaSetOnError(const Value: TNotifyEvent);
    function dsaGetDefValue: string;
    procedure dsaSetDefValue(const Value: string);
  public
    constructor Create(const aWidth, aHeight: integer;
      const aDefValue: string = '');
    destructor Destroy; override;
    procedure SetSize(const aWidth, aHeight: integer;
      const aClear: boolean = false);
    procedure SetValue(const aX, aY: integer; const aValue: string);
    function GetValue(const aX, aY: integer): string;
    procedure Clear;
    property Header: Pointer read FHeader; // Работа через указатель не безопасна, нет проверки на выход
    property Widht: integer read FWidth;
    property Height: integer read FHeight;
    property DefValue: string read dsaGetDefValue write dsaSetDefValue;
    property LastError: string read FLastError;
    property OnError: TNotifyEvent read dsaGetOnError write dsaSetOnError;
  end;

  TSingleStrArray = class
  strict private
    FHeader: Pointer;
    FSArray: array of string;
    FWidth: integer;
    FOnError: TNotifyEvent;
    FDefValue: string;
    FLastError: string;
    function ssaGetOnError: TNotifyEvent;
    procedure ssaSetOnError(const Value: TNotifyEvent);
    function sdaGetDefValue: string;
    procedure sdaSetDefValue(const Value: string);
  public
    constructor Create(const aWidth: integer; const aDefValue: string = '');
    destructor Destroy; override;
    procedure SetSize(const aWidth: integer);
    procedure SetValue(const aX: integer; const aValue: string);
    function GetValue(const aX: integer): string;
    function AsText: string;
    procedure Clear;
    property Header: Pointer read FHeader; // Работа через указатель не безопасна, нет проверки на выход
    property Widht: integer read FWidth;
    property DefValue: string read sdaGetDefValue write sdaSetDefValue;
    property LastError: string read FLastError;
    property OnError: TNotifyEvent read ssaGetOnError write ssaSetOnError;
  end;

implementation

{ TDoubleStrArray }
{ TDoubleStrArray.Clear

  Очистка массива путем заполнения значением по умолчанию }
procedure TDoubleStrArray.Clear;
var
  x, y: integer;
begin
  try
    for x := 0 to High(FDArray) do
      for y := 0 to High(FDArray[x]) do
        FDArray[x, y] := FDefValue; // Очистка значением по умолчанию
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|Clear: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TDoubleStrArray.Create

  Первоначальная инициализация и создание массива }
constructor TDoubleStrArray.Create(const aWidth, aHeight: integer;
  const aDefValue: string);
var
  x, y: integer;
begin
  inherited Create;
  try
    if (aWidth < 0) or (aHeight < 0) then
      raise Exception.Create('Unallowable size');
    SetLength(FDArray, aHeight);
    for y := 0 to High(FDArray) do
      SetLength(FDArray[y], aWidth);
    FDefValue := aDefValue;
    for x := 0 to High(FDArray) do
      for y := 0 to High(FDArray) do
        FDArray[x, y] := FDefValue;
    FHeader := @FDArray[ Low(FDArray)];
    FWidth := aWidth;
    FHeight := aHeight;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|Create: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TDoubleStrArray.Destroy

  Перед уничтожением, очистка всего }
destructor TDoubleStrArray.Destroy;
var
  i: integer;
begin
  FHeader := nil;
  for i := 0 to High(FDArray) do
    SetLength(FDArray[i], 0);
  SetLength(FDArray, 0);
  inherited;
end;

{ TDoubleStrArray.dsaGetOnError

  Возвращает объект возникающих ошибок }
function TDoubleStrArray.dsaGetOnError: TNotifyEvent;
begin
  Result := FOnError;
end;

{ TDoubleStrArray.dsaSetOnError

  Установка обработчика ошибок }
procedure TDoubleStrArray.dsaSetOnError(const Value: TNotifyEvent);
begin
  FOnError := Value;
end;

{ TDoubleStrArray.GetValue

  Возвращает значение массива по его индексам }
function TDoubleStrArray.GetValue(const aX, aY: integer): string;
begin
  try
    if Assigned(FDArray) and (Length(FDArray) > 0) then
    begin
      if (aX < 0) or (aX > FWidth) or (aY < 0) or (aY > FHeight) then
        raise Exception.Create('Unallowable value position');
      Result := FDArray[aX, aY];
    end;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|GetValue: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TDoubleStrArray.dsaGetDefValue

  Возвращает назначенное стандартное значение }
function TDoubleStrArray.dsaGetDefValue: string;
begin
  Result := FDefValue;
end;

procedure TDoubleStrArray.dsaSetDefValue(const Value: string);
begin
  FDefValue := Value;
end;

procedure TDoubleStrArray.SetSize(const aWidth, aHeight: integer;
  const aClear: boolean);
var
  y: integer;
begin
  try
    if (aWidth < 0) or (aHeight < 0) then
      raise Exception.Create('Unallowable size');
    SetLength(FDArray, aHeight);
    for y := 0 to High(FDArray) do
      SetLength(FDArray[y], aWidth);
    if aClear then
      Clear;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|SetSize: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TDoubleStrArray.SetValue

  Установка значения массива по его индексу }
procedure TDoubleStrArray.SetValue(const aX, aY: integer; const aValue: string);
begin
  try
    if (aX < 0) or (aX > FWidth) or (aY < 0) or (aY > FHeight) then
      raise Exception.Create('Unallowable value position');
    FDArray[aX, aY] := aValue;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|SetValue: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray }

{ TSingleStrArray.AsText

  Сериализация массива }
function TSingleStrArray.AsText: string;
var
  i: integer;
begin
  try
    if Length(FSArray) > 0 then
      for i := 0 to High(FSArray) do
        Result := Result + FSArray[i] + #13#10;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|Clear: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.Clear

  Очистка массива }
procedure TSingleStrArray.Clear;
var
  x: integer;
begin
  try
    for x := 0 to High(FSArray) do
      FSArray[x] := FDefValue;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|Clear: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.Create

  Создание и инициализация массива }
constructor TSingleStrArray.Create(const aWidth: integer;
  const aDefValue: string);
var
  x: integer;
begin
  inherited Create;
  try
    if aWidth < 0 then
      raise Exception.Create('Unallowable size');
    FDefValue := aDefValue;
    SetLength(FSArray, aWidth);
    for x := 0 to High(FSArray) do
      FSArray[x] := FDefValue;
    FHeader := @FSArray[ Low(FSArray)];
    FWidth := aWidth;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|Create: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.Destroy

  Очистка и уничтожение }
destructor TSingleStrArray.Destroy;
begin
  FHeader := nil;
  SetLength(FSArray, 0);
  inherited;
end;

{ TSingleStrArray.GetValue

  Возвращает значение массива по его индексу }
function TSingleStrArray.GetValue(const aX: integer): string;
begin
  try
    if (aX < 0) or (aX > FWidth) then
      raise Exception.Create('Unallowable value position');
    Result := FSArray[aX];
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|GetValue: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.sdaGetDefValue

  Возвращает назначенное стандартное значение }
function TSingleStrArray.sdaGetDefValue: string;
begin
  Result := FDefValue;
end;

{ TSingleStrArray.sdaSetDefValue

  Устанавливает стандартное значение ячейки массива }
procedure TSingleStrArray.sdaSetDefValue(const Value: string);
begin
  FDefValue := Value;
end;

{ TSingleStrArray.SetSize

  Установка размера массива }
procedure TSingleStrArray.SetSize(const aWidth: integer);
begin
  try
    if aWidth < 0 then
      raise Exception.Create('Unallowable size');
    SetLength(FSArray, aWidth);
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|SetSize: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.SetValue

  Установка значения массива }
procedure TSingleStrArray.SetValue(const aX: integer; const aValue: string);
begin
  try
    if (aX < 0) or (aX > High(FSArray)) then
      raise Exception.Create('Unallowable position');
    FSArray[aX] := aValue;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FLastError := DateTimeToStr(now) + '|SetValue: ' + E.Message;
{$ELSE}
      FLastError := E.Message;
{$ENDIF}
      FOnError(Self);
    end;
  end;
end;

{ TSingleStrArray.ssaGetOnError

  Возвращает объект возникающих ошибок }
function TSingleStrArray.ssaGetOnError: TNotifyEvent;
begin
  Result := FOnError;
end;

{ TSingleStrArray.ssaSetOnError

  Установка обработчика ошибок }
procedure TSingleStrArray.ssaSetOnError(const Value: TNotifyEvent);
begin
  FOnError := Value;
end;

end.
