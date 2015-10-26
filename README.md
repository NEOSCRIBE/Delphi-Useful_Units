# Delphi-Useful_Units
Полезные дополнения написанные мной.

Пример использования:

uses uExcel, uArray;

var
  Excel: TExcelObj;
  mData: TDoubleStrArray;
  mHeader: TSingleStrArray;
begin
  try
    // Создаем объект для работы с Экселем
    Excel := TExcelObj.Create(true, true); // Отключили показ предупреждений и сделали видимым
    Excel.AddWorkBook; // Добавили книгу
    Excel.AddWorkSheet; // Добавили лист
    // Создаем объект для работы с массивом данных
    mData := TDoubleStrArray.Create(10, 10, 'DefVal'); // Создаем массив 10 на 10 со значениями по умолчанию 'DefVal'
    mData.SetValue(0, 0, 'scribe'); // Ячейке с индексами (0, 0) присвоили значение 'scribe'
    // Это будет наш заголовок
    mHeader := TSingleStrArray.Create(10, 'header'); // Создаем массив размером 10 ячеек со значением по умолчанию 'header'
    // Вывод данных в Эксель
    Excel.UseClipboard := true; // Для вставки используется буффер
    Excel.InsertArray(0, 1, mData.Header);
    // Вывод заголовка
    Excel.InsertHeader(0, 0, mHeader.Header);
    // Примечание: вывод осуществляется к текущему листу, т.е. чтобы вывести данные в другой, надо его сначала выбрать методами SetBookActive и/или SetSheetActive класса TExcelObj. Также есть обработка ошибок и их лог, по названиям методов должно быть все понятно.
  finally
    mData.Free; 
    mHeader.Free;
    // Excel.Free // Если его освободить то изменения будут утрачены (Эксель закроется), это следует делать после сохранения.
  end;
end;

Буду рад критике и пожеланиям по увеличению функционала.
Спасибо.
