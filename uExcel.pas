{

  Юнит для работы с файлами Эксель
  ver.1.0

  author: scribe
  date: 08.10.2015
  link: justscribe@yahoo.com

  Описание:
  Возможно, упрощение работы с экселем. Применим для тривиальных задач.
  В качестве массива передается ссылка на его первый элемент (@MyHeader[Low(MyHeader)])),
  или свойство Header класса TDoubleStrArray или TSingleStrArray (они в uArray).
}
{$DEFINE Debug}
unit uExcel;

interface

uses
  SysUtils, Windows, Graphics, Classes, ComObj, // модули Делфи
  ClipBrd, // Буфер обмена
  Variants, ActiveX; // Для работы Эксель

// Excel consts (MSDN)
const
  ExcelApp = 'Excel.Application';
  // xlUnderlineStyle
  xlUnderlineStyleNone = -4142;
  xlUnderlineStyleSingle = 2;
  // xlLineStyles
  xlLineStyleNone = -4142;
  xlContinuous = 1;
  xlDash = -4115;
  xlDashDot = 4;
  xlDot = -4118;
  xlDouble = -4119;
  // xlHorizontalAlignment
  xlGeneral = 1;
  xlLeft = 2;
  xlCenter = 3;
  xlRight = 4;
  // xlWindowState
  xlMaximized = -4137;
  xlMinimized = -4140;
  xlNormal = -4143;

type
  TFStyle = set of (exBold, exItalic, exStrikeThrough, exSingleUnderline);
  TFBorder = (exNone, exSingle, exDot, exDouble, exDash, exDashDot);
  TFAlignment = (exNormal, exLeft, exCenter, exRight);

  TExcelFont = class
  strict private
    FSize: integer;
    FColor: integer;
    FBdColor: integer;
    FName: string;
    FStyle: TFStyle;
    FBorder: TFBorder;
    FAlignment: TFAlignment;
    FNumberFormat: string;
  public
    constructor Create(const aSize, aColor, aBdColor: integer;
      const aName: string; aStyle: TFStyle; aBorder: TFBorder;
      aAlignment: TFAlignment; const aNumberFormat: string = 'General');
    procedure SetFont(aSel: Variant);
    // property NumberFormat:
  end;

  TStrAArray = array of array of string;
  TStrArray = array of string;

  TExcelObj = class
  strict private
    FExcelFile: OleVariant;
    FInstalled: boolean;
    FOk: boolean;
    FWBAdded: boolean;
    FSheetsCount: integer;
    FUseClipboard: boolean;
    FErrorList: TStringList;
    FOnError: TNotifyEvent;
    class function exInstalled: boolean;
    function exGetVisible: boolean;
    procedure exSetVisible(const aVisible: boolean);
    function exGetSheetCount: integer;
    function exGetBookCount: integer;
    function exGetLastError: string;
    procedure exSetDefaultHeader(aSel: Variant);
    procedure exSetCustomHeader(aSel: Variant; aFont: TExcelFont);
    function exGetOnError: TNotifyEvent;
    procedure exSetOnError(aEvent: TNotifyEvent);
    function exGetUseClipboard: boolean;
    procedure exSetUseClipboard(const aUseClipboard: boolean);
    procedure exInsertArrayByCell(const aX, aY: integer;
      var aArray: TStrAArray; aFont: TExcelFont);
    procedure exInsertArrayByClipboard(const aX, aY: integer;
      var aArray: TStrAArray; aFont: TExcelFont);
    procedure exInsertHeaderByCell(const aX, aY: integer; aArray: Pointer;
      const aDefFormat: boolean; aFont: TExcelFont);
    procedure exInsertHeaderByClipboard(const aX, aY: integer; aArray: Pointer;
      const aDefFormat: boolean; aFont: TExcelFont);
  public
    constructor Create(const aDisableAlerts: boolean = true;
      const aVisible: boolean = false);
    destructor Destroy; override;
    procedure AddWorkBook;
    procedure AddWorkSheet;
    procedure ChangeWorkSheetName(const aIndex: integer; const aName: string;
      const aActivate: boolean = false);
    procedure SetBookActive(aIndex: integer);
    procedure SetSheetActive(aIndex: integer);
    procedure InsertArray(const aX, aY: integer; var aArray: TStrAArray;
      aFont: TExcelFont = nil);
    procedure InsertArrayFromClipboard(const aX, aY: integer;
      aFont: TExcelFont = nil);
    procedure InsertHeader(const aX, aY: integer; aArray: Pointer;
      const aDefFormat: boolean = true; aFont: TExcelFont = nil);
    procedure ChangeFont(const aRect: TRect; aFont: TExcelFont);
    function AllErrors: string;
    property SheetCount: integer read exGetSheetCount;
    property BookCount: integer read exGetBookCount;
    property Visible: boolean read exGetVisible write exSetVisible;
    property Installed: boolean read FInstalled;
    property Available: boolean read FOk;
    property UseClipboard: boolean read exGetUseClipboard write
      exSetUseClipboard;
    property LastError: string read exGetLastError;
    property OnError: TNotifyEvent read exGetOnError write exSetOnError;
  end;

implementation

{ TExcelObj }

// Добавляем книгу
procedure TExcelObj.AddWorkBook;
var
  i: integer;
begin
  FWBAdded := false;
  if FOk then
    try
      FExcelFile.WorkBooks.Add;
      FWBAdded := true;
      for i := 1 to FExcelFile.WorkSheets.Count - 1 do
        FExcelFile.WorkSheets[i].Delete;
      FExcelFile.WorkSheets[1].Name := 'Sheet_1';
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|AddWorkBook: ' + E.Message);
        FWBAdded := false;
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FWBAdded := false;
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Добавляем лист
procedure TExcelObj.AddWorkSheet;
begin
  if FOk then
    try
      if not FWBAdded or (BookCount = 0) then
        raise Exception.Create('WorkBook not added, total count: ' + inttostr
            (BookCount));
      FExcelFile.WorkSheets.Add;
      FExcelFile.WorkSheets[1].Name := 'Sheet_' + inttostr
        (FExcelFile.WorkSheets.Count);
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|AddWorkSheet: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Изменение шрифта для выделенной области
procedure TExcelObj.ChangeFont(const aRect: TRect; aFont: TExcelFont);
begin
  if FOk then
    try
      if (aRect.Left < 0) or (aRect.Top < 0) or (aRect.Right < 0) or
        (aRect.Bottom < 0) then
        raise Exception.Create('Unallowable range: ' + inttostr(aRect.Left)
            + ' - ' + inttostr(aRect.Top) + ' | ' + inttostr(aRect.Right)
            + ' - ' + inttostr(aRect.Bottom));
      FExcelFile.ActiveSheet.Range[FExcelFile.ActiveSheet.Cells[aRect.Top + 1,
        aRect.Left + 1], FExcelFile.ActiveSheet.Cells[aRect.Bottom + 1,
        aRect.Right + 1]].Select;
      aFont.SetFont(FExcelFile.Selection);
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|ChangeFont: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Установка имени листа по индексу
procedure TExcelObj.ChangeWorkSheetName(const aIndex: integer;
  const aName: string; const aActivate: boolean = false);
begin
  if FOk then
    try
      if (aIndex < 0) or (aIndex > SheetCount - 1) then
        raise Exception.Create('Unallowable index: ' + inttostr(aIndex));
      FExcelFile.WorkSheets[aIndex + 1].Name := aName;
      if aActivate then
        FExcelFile.WorkSheets[aIndex].Activate;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add
          (DateTimeToStr(now) + '|ChangeWorkSheetName: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Создаем объект
constructor TExcelObj.Create(const aDisableAlerts: boolean = true;
  const aVisible: boolean = false);
begin
  inherited Create;
  FOk := false;
  FInstalled := exInstalled;
  if FInstalled then
    try
      FExcelFile := CreateOleObject(ExcelApp);
      FErrorList := TStringList.Create;
      FExcelFile.Application.EnableEvents := aDisableAlerts;
      FExcelFile.Visible := aVisible;
      FExcelFile.WindowState := xlMaximized;
      FWBAdded := false;
      FUseClipboard := true;
      // по умолчанию используем буфер
      FOk := true;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|Create: ' + E.Message);
        FOk := false;
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOk := false;
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Убиваем Эксель
destructor TExcelObj.Destroy;
var
  i: integer;
begin
  if FExcelFile.Visible then
    FExcelFile.Visible := false;
  for i := 1 to FExcelFile.WorkBooks.Count do
    FExcelFile.WorkBooks[i].Saved := true;
  FExcelFile.Quit;
  FExcelFile := Unassigned;
  FErrorList.Free;
  inherited Destroy;
end;

procedure TExcelObj.exInsertArrayByCell(const aX, aY: integer;
  var aArray: TStrAArray; aFont: TExcelFont);
var
  x, y: integer;
begin
  try
    try
      FExcelFile.ScreenUpdating := false;
      if (aX < 0) or (aY < 0) then
        raise Exception.Create('Unallowable position: ' + inttostr(aX)
            + ' - ' + inttostr(aY));
      for x := 0 to High(aArray) do
        for y := 0 to High(aArray[x]) do
          FExcelFile.ActiveSheet.Cells[y + 1, x + 1].Value := aArray[x, y];
      FExcelFile.ActiveSheet.Range[FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1]
        , FExcelFile.ActiveSheet.Cells[aY + 1, aX + x]].Select;
      if aFont <> nil then
        aFont.SetFont(FExcelFile.ActiveSheet.Selection);
    finally
      FExcelFile.ScreenUpdating := true;
    end;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FErrorList.Add(DateTimeToStr(now) + '|exInsertArrayByCell: ' + E.Message);
      FOnError(Self);
{$ELSE}
      FErrorList.Add(E.Message);
      FOnError(Self);
{$ENDIF}
    end;
  end;
end;

procedure TExcelObj.exInsertArrayByClipboard(const aX, aY: integer;
  var aArray: TStrAArray; aFont: TExcelFont);
var
  outStr, sep: string;
  x, y: integer;
begin
  try
    try
      FExcelFile.ScreenUpdating := false;
      if (aX < 0) or (aY < 0) then
        raise Exception.Create('Unallowable position: ' + inttostr(aX)
            + ' - ' + inttostr(aY));
      sep := '';
      outStr := '';
      for x := 0 to High(aArray) do
      begin
        sep := '';
        for y := 0 to High(aArray[x]) do
        begin
          outStr := outStr + sep + aArray[x, y];
          sep := #9;
        end;
        outStr := outStr + #13#10;
      end;
      Clipboard.AsText := outStr;
      FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1].Select;
      FExcelFile.ActiveSheet.Paste;
      Clipboard.AsText := '';
      FExcelFile.ActiveSheet.Range[FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1]
        , FExcelFile.ActiveSheet.Cells[aY + 1, aX + x]].Select;
      if aFont <> nil then
        aFont.SetFont(FExcelFile.ActiveSheet.Selection);
    finally
      FExcelFile.ScreenUpdating := true;
    end;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FErrorList.Add(DateTimeToStr(now) + '|exInsertArrayByClipboard: ' +
          E.Message);
      FOnError(Self);
{$ELSE}
      FErrorList.Add(E.Message);
      FOnError(Self);
{$ENDIF}
    end;
  end;
end;

procedure TExcelObj.exInsertHeaderByCell(const aX, aY: integer;
  aArray: Pointer; const aDefFormat: boolean; aFont: TExcelFont);
var
  x: integer;
begin
  try
    try
      if (aX < 0) or (aY < 0) then
        raise Exception.Create('Unallowable position:' + inttostr(aX)
            + ' - ' + inttostr(aY));
      FExcelFile.ScreenUpdating := false;
      for x := 0 to High(TStrArray(aArray)) do
        FExcelFile.ActiveSheet.Cells[aY + 1, aX + x + 1].Value := TStrArray
          (aArray)[x];
      FExcelFile.ActiveSheet.Range[FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1]
        , FExcelFile.ActiveSheet.Cells[aY + 1, aX + x]].Select;
      if aDefFormat then
        exSetDefaultHeader(FExcelFile.Selection)
      else if aFont <> nil then
        exSetCustomHeader(FExcelFile.Selection, aFont)
      else
        raise Exception.Create('Using custom font without ExcelFont object');
      FExcelFile.ActiveSheet.Cells[aY + 2, aX + 1].Select;
      FExcelFile.ActiveWindow.FreezePanes := true;
    finally
      FExcelFile.ScreenUpdating := true;
    end;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FErrorList.Add(DateTimeToStr(now)
          + '|exInsertHeaderByCell(: ' + E.Message);
      FOnError(Self);
{$ELSE}
      FErrorList.Add(E.Message);
      FOnError(Self);
{$ENDIF}
    end;
  end;
end;

procedure TExcelObj.exInsertHeaderByClipboard(const aX, aY: integer;
  aArray: Pointer; const aDefFormat: boolean; aFont: TExcelFont);
var
  outStr, sep: string;
  x: integer;
  headerSel: Variant;
begin
  try
    try
      if (aX < 0) or (aY < 0) then
        raise Exception.Create('Unallowable position:' + inttostr(aX)
            + ' - ' + inttostr(aY));
      sep := '';
      outStr := '';
      for x := 0 to High(TStrArray(aArray)) do
      begin
        outStr := outStr + sep + TStrArray(aArray)[x];
        sep := #9;
      end;
      Clipboard.AsText := outStr;
      FExcelFile.ScreenUpdating := false;
      FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1].Select;
      FExcelFile.ActiveSheet.Paste;
      Clipboard.AsText := '';
      FExcelFile.ActiveSheet.Range[FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1]
        , FExcelFile.ActiveSheet.Cells[aY + 1, aX + x]].Select;
      if aDefFormat then
        exSetDefaultHeader(FExcelFile.Selection)
      else if aFont <> nil then
        exSetCustomHeader(FExcelFile.Selection, aFont)
      else
        raise Exception.Create('Using custom font without ExcelFont object');
      FExcelFile.ActiveSheet.Cells[aY + 2, aX + 1].Select;
      FExcelFile.ActiveWindow.FreezePanes := true;
    finally
      FExcelFile.ScreenUpdating := true;
    end;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FErrorList.Add(DateTimeToStr(now) + '|exInsertHeaderByClipboard: ' +
          E.Message);
      FOnError(Self);
{$ELSE}
      FErrorList.Add(E.Message);
      FOnError(Self);
{$ENDIF}
    end;
  end;
end;

// Проверка, установлен ли Эксель
class function TExcelObj.exInstalled: boolean;
var
  ClassID: TCLSID;
  Rez: HRESULT;
begin
  Rez := CLSIDFromProgID(PWideChar(WideString(ExcelApp)), ClassID);
  if Rez = S_OK then // Объект найден
    Result := true
  else
    Result := false;
end;

// Установка базового стиля для шапки
procedure TExcelObj.exSetCustomHeader(aSel: Variant; aFont: TExcelFont);
begin
  if Assigned(aFont) and (aFont <> nil) then
    try
      aFont.SetFont(aSel);
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|exSetCustomHeader: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

procedure TExcelObj.exSetDefaultHeader(aSel: Variant);
begin
  try
    aSel.HorizontalAlignment := xlCenter;
    aSel.Borders.LineStyle := xlContinuous;
    aSel.Interior.Color := RGB(255, 192, 0);
    aSel.Font.Bold := true;
    aSel.Font.Name := 'Verdana';
    aSel.Font.Size := 10;
  except
    on E: Exception do
    begin
{$IFDEF Debug}
      FErrorList.Add(DateTimeToStr(now) + '|exSetDefaultHeader: ' + E.Message);
      FOnError(Self);
{$ELSE}
      FErrorList.Add(E.Message);
      FOnError(Self);
{$ENDIF}
    end;
  end;
end;

procedure TExcelObj.exSetOnError(aEvent: TNotifyEvent);
begin
  FOnError := aEvent;
end;

procedure TExcelObj.exSetUseClipboard(const aUseClipboard: boolean);
begin
  FUseClipboard := aUseClipboard;
end;

// Установка видимости
procedure TExcelObj.exSetVisible(const aVisible: boolean);
begin
  if FOk then
    try
      FExcelFile.Visible := aVisible;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|exSetVisible: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Вставка данных
procedure TExcelObj.InsertArray(const aX, aY: integer; var aArray: TStrAArray;
  aFont: TExcelFont);
begin
  if FOk then
    try
      if FUseClipboard then
        exInsertArrayByClipboard(aX, aY, aArray, aFont)
      else
        exInsertArrayByCell(aX, aY, aArray, aFont);
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|InsertArray: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Вставка данных из буфера в ячейку
procedure TExcelObj.InsertArrayFromClipboard(const aX, aY: integer;
  aFont: TExcelFont);
var
  x, y: integer;
begin
  if FOk then
    try
      try
        FExcelFile.ScreenUpdating := false;
        if (aX < 0) or (aY < 0) then
          raise Exception.Create('Unallowable position: ' + inttostr(aX)
              + ' - ' + inttostr(aY));
        FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1].Select;
        FExcelFile.ActiveSheet.Paste;
        if aFont <> nil then
        begin
          FExcelFile.ActiveSheet.Range
            [FExcelFile.ActiveSheet.Cells[aY + 1, aX + 1],
            FExcelFile.ActiveSheet.Cells[aY + 1, aX + x]].Select;
          aFont.SetFont(FExcelFile.ActiveSheet.Selection);
        end;
      finally
        FExcelFile.ScreenUpdating := true;
      end;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now)
            + '|InsertArrayFromClipboard: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Вставка шапки
procedure TExcelObj.InsertHeader(const aX, aY: integer; aArray: Pointer;
  const aDefFormat: boolean = true; aFont: TExcelFont = nil);
var
  headerSel: Variant;
begin
  if FOk then
    try
      if FUseClipboard then
        exInsertHeaderByClipboard(aX, aY, aArray, aDefFormat, aFont)
      else
        exInsertHeaderByCell(aX, aY, aArray, aDefFormat, aFont);
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|InsertHeader: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

function TExcelObj.AllErrors: string;
begin
  Result := FErrorList.Text;
end;

// Устанавливаем активной книгу по ее индексу
procedure TExcelObj.SetBookActive(aIndex: integer);
begin
  if FOk then
    try
      if (aIndex < 0) or (aIndex > BookCount - 1) then
        raise Exception.Create('Unallowable index: ' + inttostr(aIndex));
      FExcelFile.WorkBooks[aIndex + 1].Activate;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|SetBookActive: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Устанавливаем активным лист по его индексу
procedure TExcelObj.SetSheetActive(aIndex: integer);
begin
  if FOk then
    try
      if (aIndex < 0) or (aIndex > SheetCount - 1) then
        raise Exception.Create('Unallowable index: ' + inttostr(aIndex));
      FExcelFile.ActiveWorkbook.WorkSheets[aIndex + 1].Activate;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|SetSheetActive: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

// Узнаем количество книг
function TExcelObj.exGetBookCount: integer;
begin
  if FOk then
    try
      Result := FExcelFile.WorkBooks.Count;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|exGetBookCount: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

function TExcelObj.exGetLastError: string;
begin
  Result := FErrorList.Strings[FErrorList.Count - 1];
end;

function TExcelObj.exGetOnError: TNotifyEvent;
begin
  Result := FOnError;
end;

// Узнаем количество листов в активной книге
function TExcelObj.exGetSheetCount: integer;
begin
  if FOk then
    try
      Result := FExcelFile.ActiveWorkbook.WorkSheets.Count;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|exGetSheetCount: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

function TExcelObj.exGetUseClipboard: boolean;
begin
  Result := FUseClipboard;
end;

// Видимость окна
function TExcelObj.exGetVisible: boolean;
begin
  if FOk then
    try
      Result := FExcelFile.Visible;
    except
      on E: Exception do
      begin
{$IFDEF Debug}
        FErrorList.Add(DateTimeToStr(now) + '|exGetVisible: ' + E.Message);
        FOnError(Self);
{$ELSE}
        FErrorList.Add(E.Message);
        FOnError(Self);
{$ENDIF}
      end;
    end;
end;

{ TExcelFont }

constructor TExcelFont.Create(const aSize, aColor, aBdColor: integer;
  const aName: string; aStyle: TFStyle; aBorder: TFBorder;
  aAlignment: TFAlignment; const aNumberFormat: string);
begin
  inherited Create;
  FSize := aSize;
  FColor := aColor;
  FBdColor := aBdColor;
  FName := aName;
  FStyle := aStyle;
  FBorder := aBorder;
  FAlignment := aAlignment;
  FNumberFormat := aNumberFormat;
end;

procedure TExcelFont.SetFont(aSel: Variant);
begin
  aSel.Font.Size := FSize;
  aSel.Font.Color := FColor;
  aSel.Font.Name := FName;
  aSel.Interior.Color := FBdColor;
  aSel.NumberFormat := FNumberFormat;
  // --HEADER FONT STYLES
  if exBold in FStyle then
    aSel.Font.Bold := true
  else
    aSel.Font.Bold := false;
  // -----------------------------
  if exItalic in FStyle then
    aSel.Font.Italic := true
  else
    aSel.Font.Italic := false;
  // -----------------------------
  if exStrikeThrough in FStyle then
    aSel.Font.StrikeThrough := true
  else
    aSel.Font.StrikeThrough := false;
  // -----------------------------
  if exSingleUnderline in FStyle then
    aSel.Font.Underline := xlUnderlineStyleSingle
  else
    aSel.Font.Underline := xlUnderlineStyleNone;
  // --HEADER BORDER STYLES
  case FBorder of
    exNone:
      aSel.Borders.LineStyle := xlLineStyleNone;
    exSingle:
      aSel.Borders.LineStyle := xlContinuous;
    exDot:
      aSel.Borders.LineStyle := xlDot;
    exDouble:
      aSel.Borders.LineStyle := xlDouble;
    exDash:
      aSel.Borders.LineStyle := xlDash;
    exDashDot:
      aSel.Borders.LineStyle := xlDashDot;
  end;
  // --HEADER ALIGNMENT
  case FAlignment of
    exNormal:
      aSel.HorizontalAlignment := xlGeneral;
    exLeft:
      aSel.HorizontalAlignment := xlLeft;
    exCenter:
      aSel.HorizontalAlignment := xlCenter;
    exRight:
      aSel.HorizontalAlignment := xlRight;
  end;
end;

end.
