//Функция, проверяющая, установлен ли MS Excel/OpenOffice или нет
function IsOLEObjectInstalled(Name: String): boolean;
var
 ClassID: TCLSID;
 Rez: HRESULT;
begin
 Rez:= CLSIDFromProgID(PWideChar(WideString(Name)), ClassID);
 If Rez = S_OK Then
  Result:=True
 Else
  Result:=False;
end;

//Процедура автоматического выравнивания ячеек по ширине

procedure AutoFit(Imya_Tablicy : TStringGrid);
var i, j, temp, max: integer; Stroka1, Stroka2: string;
{Stroka1, Stroka2 - переменные, в которых будут храниться названия таблиц
перед очисткой ячеек с названиями таблицы
Это нужно для того, чтобы выравнивание ячеек по ширине происходило без учета длинных заголовков}
Begin
 Stroka1:=Imya_Tablicy.Cells[0,0];
 Stroka2:=Imya_Tablicy.Cells[0,1];
 Imya_Tablicy.Cells[0,0]:='';
 Imya_Tablicy.Cells[0,1]:='';
 for i := 0 to Imya_Tablicy.Colcount - 1 do
 begin
  max := 0;
  for j := 0 to Imya_Tablicy.Rowcount - 1 do
  begin
   temp := Imya_Tablicy.canvas.textWidth(Imya_Tablicy.cells[i, j]);
   if temp > max then
    max := temp;
  end;
  Imya_Tablicy.colWidths[i] := max + Imya_Tablicy.gridLineWidth + 5;
 end;
 Imya_Tablicy.Cells[0,0]:=Stroka1;
 Imya_Tablicy.Cells[0,1]:=Stroka2;
 temp:=UnAssigned;
 max:=UnAssigned;
 Stroka1:=UnAssigned;
 Stroka2:=UnAssigned;
 i:=UnAssigned;
 j:=UnAssigned;
end;

//Функция суммирования ячеек по вертикали
function SummaV(Imya_Tablici : TStringGrid; stolbec1: Integer; stroka1: Integer; stolbec2: Integer; stroka2:Integer): string;
var Summ1:Float; schetchik:Integer;
Begin
 stolbec1:=stolbec2;
 Summ1:=0;
 schetchik:=0;
 for schetchik := stroka1 To stroka2 do
 begin
  Summ1 := Summ1 + StrToFloat(Imya_Tablici.Cells[stolbec1,schetchik]);
 end;
 SummaV:=FloatToStr(Summ1); //возвращаем полученное значение суммы
 Summ1:=UnAssigned; //удаляем из памяти ненужные переменные
 schetchik:=UnAssigned;
End;

//Функция поиска минимума
function FindMin(Imya_Tablici : TStringGrid; stolbec1: Integer; stroka1: Integer; stolbec2: Integer; stroka2:Integer): string;
var Min1:Float; schetchik:Integer;
Begin
 stolbec1:=stolbec2;
 Min1:=StrToFloat(Imya_Tablici.Cells[stolbec1,stroka1]);
 schetchik:=0;
 for schetchik := stroka1 To stroka2 do
 begin
  if (StrToFloat(Imya_Tablici.Cells[stolbec1,schetchik]) < Min1) then
   Min1 := StrToFloat(Imya_Tablici.Cells[stolbec1,schetchik]);
 end;
 FindMin:=FloatToStr(Min1); //возвращаем полученное значение суммы
 Min1:=UnAssigned; //удаляем из памяти ненужные переменные
 schetchik:=UnAssigned;
End;

//Функция поиска максимума
function FindMax(Imya_Tablici : TStringGrid; stolbec1: Integer; stroka1: Integer; stolbec2: Integer; stroka2:Integer): string;
var Max1:Float; schetchik:Integer;
Begin
 stolbec1:=stolbec2;
 Max1:=0;
 schetchik:=0;
 Max1 := StrToFloat(Imya_Tablici.Cells[stolbec1,stroka1]);
 for schetchik := stroka1 To stroka2 do
 begin
  if (StrToFloat(Imya_Tablici.Cells[stolbec1,schetchik]) > Max1) then
   Max1 := StrToFloat(Imya_Tablici.Cells[stolbec1,schetchik]);
 end;
 FindMax:=FloatToStr(Max1); //возвращаем полученное значение суммы
 Max1:=UnAssigned; //удаляем из памяти ненужные переменные
 schetchik:=UnAssigned;
End;

//Функция поиска среднего значения
function FindAverage(Imya_Tablici : TStringGrid; stolbec1: Integer; stroka1: Integer; stolbec2: Integer; stroka2:Integer): string;
Begin
 FindAverage := FloatToStr(StrToFloat(SummaV(Imya_Tablici, stolbec1, stroka1, stolbec2, stroka2)) / (stroka2-stroka1+1)); 
End;

{Функция нахождения ранга числа в Excel
nstl - номер столбца, nstr - номер строки, nstlb - начальный столбец, strb - начальная строка,
nstle - конечный столбец, nstre - конечная строка, sq1 это таблица типа StringGrid,
название которой указывается в параметре функции}
function FindRank(nstl, nstr, nstlb, nstrb, nstle, nstre: Integer; sg1 : TStringGrid): Integer;
var massiv: Array of Real;
    i,j, dlina_massiva: Integer;
    x, tmp, tmp1: Real;
begin
  dlina_massiva:=nstre-nstrb+1;
  SetLength(massiv,dlina_massiva+1);
//запишем набор данных из ячеек и столбцов из таблицы в массив
  for i:=1 to dlina_massiva do begin          //1 to 15
    massiv[i]:=StrToFloat(sg1.Cells[nstlb, i+nstrb-1]); end;
//упорядочим массив по убыванию
  for j := 1 to dlina_massiva do
   for i := 1 to dlina_massiva-1 do
     if massiv[i] > massiv[i + 1] then
	 {поменять знак, если нужно изменить тип упорядочивания.
	 < - по убыванию, > - по возрастанию}
        begin
          x:= massiv[i];
          massiv[i] := massiv[i + 1];
          massiv[i + 1] := x;
        end;
i:=0;
j:=0;
{найдем повторяющиеся элементы в массиве
и заменим нулями (кроме первого)}
for i:=1 to dlina_massiva-1 do
begin
 if massiv[i]=massiv[i+1] then
    begin
      tmp1:=massiv[i];
      for j:=i+1 to dlina_massiva do
           if massiv[j]=tmp1 then massiv[j]:=0;
    end
    else massiv[i]:=massiv[i];
end;
 tmp:=StrToFloat(sg1.Cells[nstl, nstr]);
 for i:=1 to dlina_massiva do
 begin
   if massiv[i]=tmp then
    begin
     FindRank:=i;
    end;
  end;
 end;

//Экспорт таблицы в Microsoft Excel
procedure ExportToExcel(Imya_Tablici : TStringGrid);
var
ExcelApp, ExcelSheet, ExcelCol, ExcelRow: variant;
Size: byte;
i, j, N: word; tempstring2:string;
float: extended; num:integer;
begin
//Формируем имя листа, которое будет отображаться в Microsoft Office Excel
tempstring:=''; tempstring2:='';
tempstring:='Table'+IntToStr(Blok1_Form.PageControl1.TabIndex+1);
tempstring2:=SysToUTF8(IniF.ReadString('Blok1_Zagolovki',tempstring,''));
//**********************************************************************//
ExcelApp := CreateOleObject('Excel.Application');
ExcelApp.Visible := True;
ExcelApp.Workbooks.Add(-4167);
ExcelApp.Workbooks[1].WorkSheets[1].Name := WideString(UTF8Decode(tempstring2));
ExcelCol := ExcelApp.Workbooks[1].WorkSheets[WideString(UTF8Decode(tempstring2))].Columns;

Size := Imya_Tablici.DefaultRowHeight;
N := Imya_Tablici.ColCount - 1;
for j := 0 to N do
   ExcelCol.Columns[j + 1].ColumnWidth := Size;
   ExcelRow := ExcelApp.Workbooks[1].WorkSheets[WideString(UTF8Decode(tempstring2))].Rows;
{выделяем заголовок жирным шрифтом}
ExcelRow.Rows[1].Font.Bold := True;
ExcelRow.Rows[2].Font.Bold := True;
{/выделяем заголовок жирным шрифтом}
ExcelSheet := ExcelApp.Workbooks[1].WorkSheets[WideString(UTF8Decode(tempstring2))];
//Обьединяем ячейки
ExcelSheet.Range['A1:Z1'].Merge;
ExcelSheet.Range['A2:Z2'].Merge;
i:=0;
for i := 0 to Imya_Tablici.RowCount - 1 do
begin
  ExcelSheet.Rows[i+1].Font.Name := 'Arial'; //во всей таблице устанавливаем шрифт Arial
end;
i:=0;
for i := 0 to Imya_Tablici.RowCount - 1 do
for j := 0 to Imya_Tablici.ColCount - 1 do
If TryStrToInt((Imya_Tablici.Cells[j, i]),num) then
Begin
DecimalSeparator:=',';
ExcelSheet.Cells[i + 1, j + 1] := StrToInt(Imya_Tablici.Cells[j, i]); //если ячейка целое число, то сохраняем её как есть
End
Else IF TryStrToFloat((Imya_Tablici.Cells[j, i]),float) then
Begin
DecimalSeparator:=',';
ExcelSheet.Cells[i + 1, j + 1] := StrToFloat(Imya_Tablici.Cells[j, i]);  //если ячейка дробное число, то сохраняем её как есть
End
//если ячейка является текстом, то перекодируем её из UTF8, для корректного отображения кириллицы
ELSE ExcelSheet.Cells[i + 1, j + 1] := WideString(Utf8Decode(Imya_Tablici.Cells[j, i]));
tempstring2:=UnAssigned; //удаляем из памяти ненужные переменные
end;

//Экспорт таблицы в OpenOffice / LibreOffice
procedure ExportToOpenOffice(Imya_Tablici : TStringGrid);
var
Size: byte;
i, j, n: word;
float: extended; num:integer;
tempstring2: string;
OO, Desktop: variant;
Doc, Sheet: variant;
Col, Cell: variant;
s: WideString;
const
BoldFont = 150;
begin
OO := CreateOleObject('com.sun.star.ServiceManager');
Desktop := OO.createInstance('com.sun.star.frame.Desktop');
Doc := Desktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, VarArrayCreate([0, -1], varVariant));
for i := Doc.getSheets.Count - 1 downto 1 do
begin
   s := WideString(Doc.getSheets.GetByIndex(i).getName);
   Doc.getSheets.RemoveByName(s);
end;
Size := Imya_Tablici.DefaultRowHeight;
Sheet := Doc.getSheets.GetByIndex(0);
//Формируем имя листа, которое будет отображаться в OpenOffice.Calc
tempstring:=''; tempstring2:='';
tempstring:='Table'+IntToStr(Blok1_Form.PageControl1.TabIndex+1);
tempstring2:=SysToUTF8(IniF.ReadString('Blok1_Zagolovki',tempstring,''));
//**********************************************************************//
//Sheet.Name := 'Report';
Sheet.Name := WideString(UTF8Decode(tempstring2));
n := Imya_Tablici.ColCount - 1;
for i := 0 to n do
begin
Col := Sheet.getColumns.GetByIndex(i);
Col.setPropertyValue('Width', 2 * 100 * Size); //В сотых долях миллиметра
end;
i:=0;
for i := 0 to Imya_Tablici.RowCount - 1 do
for j := 0 to Imya_Tablici.ColCount - 1 do
begin
Cell := Sheet.getCellByPosition(j, i);
Cell.charFontName := 'Arial'; //во всей таблице устанавливаем шрифт Arial
Cell.CharHeight:= 10; //размер шрифта 10 пикселей
 If TryStrToInt((Imya_Tablici.Cells[j, i]),num) then
 Begin
 DecimalSeparator:=',';
 Cell.SetValue(StrToInt(Imya_Tablici.Cells[j, i])); //если ячейка целое число, то сохраняем её как есть
 End
 Else IF TryStrToFloat((Imya_Tablici.Cells[j, i]),float) then
 Begin
 DecimalSeparator:=',';
 Cell.SetValue(StrToFloat(Imya_Tablici.Cells[j, i])); //если ячейка дробное число, то сохраняем её как есть
 End
 //если ячейка является текстом, то перекодируем её из UTF8, для корректного отображения кириллицы
 ELSE Cell.SetString(WideString(UTF8Decode(Imya_Tablici.Cells[j, i])));
{выделяем заголовок жирным шрифтом}
 if i < 2 then
 begin
  Sheet.getCellByPosition(j, i).getText.createTextCursor.CharWeight :=  BoldFont;
  Sheet.getCellByPosition(j, i).getText.createTextCursor.CharHeight:= 12; //размер шрифта заголовка таблицы - 12 пикселей
 end;
end;
//Обьединяем ячейки
Cell := Sheet.getCellRangeByName('A1:Z1').Merge(true);
Cell := Sheet.getCellRangeByName('A2:Z2').Merge(true);
tempstring2:=UnAssigned; //удаляем из памяти ненужные переменные
end;
//

//Экспорт таблицы в файл
procedure ExportToFile(Imya_Tablici : TStringGrid; SaveDialog1 : TSaveDialog; Tabsheet1: TTabsheet);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  i,j:integer;
  SIZE:integer;
  float: extended; num:integer;
  tempstring2: string;
  selectedBtn: Integer; //в этой переменной хранится нажатие выбранной пользователем кнопки
begin
//SaveDialog1.FileName:=TLabel(FindComponent('Label' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Caption; //формируем имя файла
SaveDialog1.FileName:='Таблица №'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_'+Tabsheet1.Caption; //формируем имя файла
If SaveDialog1.Execute then
Begin
  // сохраняем таблицу в файл формата Excel (.xls)
  // и открываем её соответствующей программой
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.SetDefaultFont('Arial', 10);
  //Формируем имя листа, которое будет отображаться при открытии файла
  tempstring:=''; tempstring2:='';
  tempstring:='Table'+IntToStr(Blok1_Form.PageControl1.TabIndex+1);
  tempstring2:=SysToUTF8(IniF.ReadString('Blok1_Zagolovki',tempstring,''));
  //
  MyWorksheet := MyWorkbook.AddWorksheet(tempstring2);
  For i:=0 to Imya_Tablici.ColCount-1 do
   For j:=0 to Imya_Tablici.RowCount-1 do
    BEGIN
    SIZE:=Imya_Tablici.DefaultRowHeight;
    MyWorksheet.WriteColWidth(j, SIZE); //стандартная ширина колонки
    MyWorksheet.WriteFont(0, 0, 'Arial', 12, [fssBold], scBlack);
    MyWorksheet.MergeCells(0, 0, 0, Imya_Tablici.ColCount-1);
    MyWorksheet.WriteFont(1, 0, 'Arial', 12, [fssBold], scBlack);
    MyWorksheet.MergeCells(1, 0, 1, Imya_Tablici.ColCount-1);
    MyWorksheet.WriteRowHeight(0, 1.5); //высота ячейки с названием таблицы (в линиях)
    MyWorksheet.WriteRowHeight(1, 1.5);
    //проверяем, является ли ячейка целым числом
     if TryStrToInt(Imya_Tablici.Cells[i,j], num) then
        begin
         MyWorksheet.WriteNumber(j, i, StrToInt(Imya_Tablici.Cells[i,j]));
        end
        //проверяем, является ли ячейка дробным числом
        else if TryStrToFloat(Imya_Tablici.Cells[i,j], float) then
        begin
         MyWorksheet.WriteNumber(j, i, StrToFloat(Imya_Tablici.Cells[i,j]));
        end
     else BEGIN
      //очищаем ячейки с названием таблицы
     MyWorksheet.WriteUTF8Text(0, 0, '');
     MyWorksheet.WriteUTF8Text(0, 1, '');
     MyWorksheet.WriteUTF8Text(j, i, Imya_Tablici.Cells[i,j]);
     //MyWorksheet.WriteUsedFormatting(j, i, [uffWordwrap]);
     end;
    end;
   MyWorksheet.WriteUTF8Text(0, 0, Imya_Tablici.Cells[0,0]);
   MyWorksheet.WriteUTF8Text(0, 1, Imya_Tablici.Cells[1,0]);
//Удаляем расширение у имени файла, если оно не EXCEL и прибавляем к имени .xls (на всякий случай)
If ExtractFileExt(SaveDialog1.FileName)<>'.xls' then SaveDialog1.FileName:=SaveDialog1.FileName+'.xls';
//открываем файл таблицы (.xls) с помощью MS Excel Viewer 97
selectedBtn:= MessageDlg('Таблица сохранена в файл ' + ExtractFileName(SaveDialog1.FileName) + '!' + #13 + #13 + 'Хотите открыть её прямо сейчас ?', mtConfirmation, [mbYes, mbNO] , 0);
 if selectedBtn = mrYes then
 begin
  MyWorkbook.WriteToFile(UTF8ToSys(SaveDialog1.FileName), sfExcel8, true); //сохраняем файл
  MyWorkbook.Free; //удаляем из памяти ненужные переменны
  tempstring2:=UnAssigned;
  SIZE:=UnAssigned;
  selectedBtn:=UnAssigned;
  ShellExecute(0, '', 'Viewer\VIEWER.EXE', PChar('"'+UTF8ToSys(SaveDialog1.FileName)+'"'), nil, SW_SHOW);
 end
 else if selectedBtn=mrNo then
 begin
  MyWorkbook.WriteToFile(UTF8ToSys(SaveDialog1.FileName), sfExcel8, true);
  MyWorkbook.Free; //удаляем из памяти ненужные переменны
  tempstring2:=UnAssigned;
  SIZE:=UnAssigned;
  selectedBtn:=UnAssigned;
  exit;
 end;
end;
end;

//Вывод сообщентя об ошибке
procedure NotFoundException(num: integer);
begin
 MessageDlg('МкАРС 1.5 :: Ошибка', 'ОШИБКА :: Файл таблицы №'+IntToStr(num)+' не найден !'+#13+#13+'Для продолжения работы нажмите ОК',mtError, [mbOk], 0);
end;
