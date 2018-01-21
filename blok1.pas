unit blok1;
{$mode objfpc}{$H+}

interface

uses ActiveX, Classes, SysUtils, FileUtil, Forms, Controls, Graphics,
     Dialogs, StdCtrls, Grids, ComCtrls, ExtCtrls, ComObj, Variants, Math,
     Windows, fpspreadsheet, xlsbiff8, types;
//компонент Math используется для округления дробных переменных

type

{ TBlok1_Form }

TBlok1_Form = class(TForm)
 BackBtn1: TButton;
 BackBtn2: TButton;
 BackBtn3: TButton;
 CalculateBtn1: TButton;
 CalculateBtn2: TButton;
 CalculateBtn3: TButton;
 CancelBtn1: TButton;
 CancelBtn2: TButton;
 CancelBtn3: TButton;
 ExportToExcelBtn1: TButton;
 ExportToExcelBtn2: TButton;
 ExportToExcelBtn3: TButton;
 ExportToOOCalcBtn1: TButton;
 ExportToOOCalcBtn2: TButton;
 ExportToOOCalcBtn3: TButton;
 ExportToXLSBtn1: TButton;
 ExportToXLSBtn2: TButton;
 ExportToXLSBtn3: TButton;
 Label1: TLabel;
 Label2: TLabel;
 Label3: TLabel;
 PageControl1: TPageControl;
 SaveChangesBtn1: TButton;
 SaveDialog1: TSaveDialog;
 sGrid1: TStringGrid;
 sGrid2: TStringGrid;
 sGrid3: TStringGrid;
 TabSheet1: TTabSheet;
 TabSheet2: TTabSheet;
 TabSheet3: TTabSheet;
procedure CalculateBtn1Click(Sender: TObject);
procedure CalculateBtn2Click(Sender: TObject);
procedure CalculateBtn3Click(Sender: TObject);
procedure ExportToExcelBtn1Click(Sender: TObject);
procedure BackBtn1Click(Sender: TObject);
procedure FormClose(Sender: TObject);
procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
procedure FormResize(Sender: TObject);
procedure PageControl1Change(Sender: TObject);
procedure CancelBtn1Click(Sender: TObject);
procedure ExportToXLSBtn1Click(Sender: TObject);
procedure ExportToOOCalcBtn1Click(Sender: TObject);
procedure FormCreate(Sender: TObject);
procedure SaveChangesBtn1Click(Sender: TObject);
procedure sGrid1DrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
procedure sGrid1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
procedure sGrid1KeyPress(Sender: TObject; var Key: char);
procedure initValues;
procedure SaveAndRefresh(num: Integer);
private
{ private declarations }
public
{ public declarations }
end;

type Znachenie = record
 Col, Row: Integer;
 Value: String;
end;

var
 Blok1_Form: TBlok1_Form;
 i,j,x: integer;
 Tabl1 : array[0..128] of Znachenie;
implementation

uses welcome;

{$R *.lfm}

{ TBlok1_Form }
{Подключаем модуль со всеми функциями, который называется functions.pas}
{$INCLUDE functions}

procedure TBlok1_Form.CalculateBtn1Click(Sender: TObject);
//Кнопка "Рассчитать". Для каждой таблицы - свой обработчик!!!
begin
//0очищаем таблицу
 with TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))) do
 for i:=FixedCols to ColCount-1 do
 for j:=FixedRows to RowCount-1 do
  Cells[i, j]:='';
  if FileExists('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml') then  //проверяем существует ли файл с таблицей
  begin
   TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).LoadFromFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml'); //считываем данные из файла Tablica1.xml
   AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));  //выравниваем ячейки по ширине, в зависимости от длины текста
  end
  else begin   //если файл не существует, выводим сообщение об ошибке
   NotFoundException(Blok1_Form.PageControl1.TabIndex+1);
   Exit; //выходим из процедуры считывания данных
  end;

 //денежные потоки от текущих операций
 for i:=15 to 28 do
 begin
  sGrid1.Cells[12,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])-StrToFloat(sGrid1.Cells[11,i]));
  sGrid1.Cells[13,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[7,i]));
  sGrid1.Cells[14,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[11,i]));

  //обработка деления на ноль, чтобы программа не выдавала ошибку
  if (StrToFloat(sGrid1.Cells[11,i])=0) then
  begin
   sGrid1.Cells[15,i]:=FloatToStr(0);
   sGrid1.Cells[17,i]:=FloatToStr(0);
  end
  else
  begin
   sGrid1.Cells[15,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])/StrToFloat(sGrid1.Cells[11,i]));
   sGrid1.Cells[17,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[11,i]));
  end;

  if (StrToFloat(sGrid1.Cells[7,i])=0) then sGrid1.Cells[16,i]:=FloatToStr(0)
  else sGrid1.Cells[16,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[7,i]));

  //округление
  sGrid1.Cells[12,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[12,i]),-4));
  sGrid1.Cells[13,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[13,i]),-4));
  sGrid1.Cells[14,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[14,i]),-4));
  sGrid1.Cells[15,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[15,i]),-4));
  sGrid1.Cells[16,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[16,i]),-4));
  sGrid1.Cells[17,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[17,i]),-4));
 end;

 //денежные потоки от инвестиционных операций
 for i:=35 to 47 do
 begin
  sGrid1.Cells[12,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])-StrToFloat(sGrid1.Cells[11,i]));
  sGrid1.Cells[13,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[7,i]));
  sGrid1.Cells[14,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[11,i]));

  //обработка деления на ноль, чтобы программа не выдавала ошибку
  if (StrToFloat(sGrid1.Cells[11,i])=0) then
  begin
   sGrid1.Cells[15,i]:=FloatToStr(0);
   sGrid1.Cells[17,i]:=FloatToStr(0);
  end
  else
  begin
   sGrid1.Cells[15,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])/StrToFloat(sGrid1.Cells[11,i]));
   sGrid1.Cells[17,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[11,i]));
  end;

  if (StrToFloat(sGrid1.Cells[7,i])=0) then sGrid1.Cells[16,i]:=FloatToStr(0)
  else sGrid1.Cells[16,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[7,i]));

  //округление
  sGrid1.Cells[12,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[12,i]),-4));
  sGrid1.Cells[13,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[13,i]),-4));
  sGrid1.Cells[14,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[14,i]),-4));
  sGrid1.Cells[15,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[15,i]),-4));
  sGrid1.Cells[16,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[16,i]),-4));
  sGrid1.Cells[17,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[17,i]),-4));
 end;

 //денежные потоки от финансовых операций
 for i:=49 to 54 do
 begin
  sGrid1.Cells[12,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])-StrToFloat(sGrid1.Cells[11,i]));
  sGrid1.Cells[13,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[7,i]));
  sGrid1.Cells[14,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[11,i]));

  //обработка деления на ноль, чтобы программа не выдавала ошибку
  if (StrToFloat(sGrid1.Cells[11,i])=0) then
  begin
   sGrid1.Cells[15,i]:=FloatToStr(0);
   sGrid1.Cells[17,i]:=FloatToStr(0);
  end
  else
  begin
   sGrid1.Cells[15,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])/StrToFloat(sGrid1.Cells[11,i]));
   sGrid1.Cells[17,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[11,i]));
  end;

  if (StrToFloat(sGrid1.Cells[7,i])=0) then sGrid1.Cells[16,i]:=FloatToStr(0)
  else sGrid1.Cells[16,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[7,i]));

  //округление
  sGrid1.Cells[12,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[12,i]),-4));
  sGrid1.Cells[13,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[13,i]),-4));
  sGrid1.Cells[14,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[14,i]),-4));
  sGrid1.Cells[15,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[15,i]),-4));
  sGrid1.Cells[16,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[16,i]),-4));
  sGrid1.Cells[17,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[17,i]),-4));
 end;

 //платежи, сальдо денежных потоков от финансовых операций и остаток денежных средств
 for i:=59 to 68 do
 begin
  sGrid1.Cells[12,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])-StrToFloat(sGrid1.Cells[11,i]));
  sGrid1.Cells[13,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[7,i]));
  sGrid1.Cells[14,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])-StrToFloat(sGrid1.Cells[11,i]));

  //обработка деления на ноль, чтобы программа не выдавала ошибку
  if (StrToFloat(sGrid1.Cells[11,i])=0) then
  begin
   sGrid1.Cells[15,i]:=FloatToStr(0);
   sGrid1.Cells[17,i]:=FloatToStr(0);
  end
  else
  begin
   sGrid1.Cells[15,i]:= FloatToStr(StrToFloat(sGrid1.Cells[7,i])/StrToFloat(sGrid1.Cells[11,i]));
   sGrid1.Cells[17,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[11,i]));
  end;

  if (StrToFloat(sGrid1.Cells[7,i])=0) then sGrid1.Cells[16,i]:=FloatToStr(0)
  else sGrid1.Cells[16,i]:= FloatToStr(StrToFloat(sGrid1.Cells[6,i])/StrToFloat(sGrid1.Cells[7,i]));

  //округление
  sGrid1.Cells[12,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[12,i]),-4));
  sGrid1.Cells[13,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[13,i]),-4));
  sGrid1.Cells[14,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[14,i]),-4));
  sGrid1.Cells[15,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[15,i]),-4));
  sGrid1.Cells[16,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[16,i]),-4));
  sGrid1.Cells[17,i]:= FloatToStr(RoundTo(StrToFloat(sGrid1.Cells[17,i]),-4));
 end;
 AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));  //выравниваем ячейки по ширине, в зависимости от длины текста
end;

procedure TBlok1_Form.initValues;
begin
 //Данная функция записывает цифры в массив - номер столбца, номер строки и значение.
 //Цифры из этого массива используются при сохранении изменений в таблице и
 //последующем перерасчете.
Tabl1[0].Col:=6; Tabl1[0].Row:=15; Tabl1[0].Value:=sGrid1.Cells[6,15];
Tabl1[1].Col:=6; Tabl1[1].Row:=16; Tabl1[1].Value:=sGrid1.Cells[6,16];
Tabl1[2].Col:=6; Tabl1[2].Row:=17; Tabl1[2].Value:=sGrid1.Cells[6,17];
Tabl1[3].Col:=6; Tabl1[3].Row:=18; Tabl1[3].Value:=sGrid1.Cells[6,18];
Tabl1[4].Col:=6; Tabl1[4].Row:=19; Tabl1[4].Value:=sGrid1.Cells[6,19];
Tabl1[5].Col:=6; Tabl1[5].Row:=20; Tabl1[5].Value:=sGrid1.Cells[6,20];
Tabl1[6].Col:=6; Tabl1[6].Row:=21; Tabl1[6].Value:=sGrid1.Cells[6,21];
Tabl1[7].Col:=6; Tabl1[7].Row:=22; Tabl1[7].Value:=sGrid1.Cells[6,22];
Tabl1[8].Col:=6; Tabl1[8].Row:=23; Tabl1[8].Value:=sGrid1.Cells[6,23];
Tabl1[9].Col:=6; Tabl1[9].Row:=24; Tabl1[9].Value:=sGrid1.Cells[6,24];
Tabl1[10].Col:=6; Tabl1[10].Row:=25; Tabl1[10].Value:=sGrid1.Cells[6,25];
Tabl1[11].Col:=6; Tabl1[11].Row:=26; Tabl1[11].Value:=sGrid1.Cells[6,26];
Tabl1[12].Col:=6; Tabl1[12].Row:=27; Tabl1[12].Value:=sGrid1.Cells[6,27];
Tabl1[13].Col:=6; Tabl1[13].Row:=28; Tabl1[13].Value:=sGrid1.Cells[6,28];
Tabl1[14].Col:=6; Tabl1[14].Row:=35; Tabl1[14].Value:=sGrid1.Cells[6,35];
Tabl1[15].Col:=6; Tabl1[15].Row:=36; Tabl1[15].Value:=sGrid1.Cells[6,36];
Tabl1[16].Col:=6; Tabl1[16].Row:=37; Tabl1[16].Value:=sGrid1.Cells[6,37];
Tabl1[17].Col:=6; Tabl1[17].Row:=38; Tabl1[17].Value:=sGrid1.Cells[6,38];
Tabl1[18].Col:=6; Tabl1[18].Row:=39; Tabl1[18].Value:=sGrid1.Cells[6,39];
Tabl1[19].Col:=6; Tabl1[19].Row:=40; Tabl1[19].Value:=sGrid1.Cells[6,40];
Tabl1[20].Col:=6; Tabl1[20].Row:=41; Tabl1[20].Value:=sGrid1.Cells[6,41];
Tabl1[21].Col:=6; Tabl1[21].Row:=42; Tabl1[21].Value:=sGrid1.Cells[6,42];
Tabl1[22].Col:=6; Tabl1[22].Row:=43; Tabl1[22].Value:=sGrid1.Cells[6,43];
Tabl1[23].Col:=6; Tabl1[23].Row:=44; Tabl1[23].Value:=sGrid1.Cells[6,44];
Tabl1[24].Col:=6; Tabl1[24].Row:=45; Tabl1[24].Value:=sGrid1.Cells[6,45];
Tabl1[25].Col:=6; Tabl1[25].Row:=46; Tabl1[25].Value:=sGrid1.Cells[6,46];
Tabl1[26].Col:=6; Tabl1[26].Row:=47; Tabl1[26].Value:=sGrid1.Cells[6,47];
Tabl1[27].Col:=6; Tabl1[27].Row:=49; Tabl1[27].Value:=sGrid1.Cells[6,49];
Tabl1[28].Col:=6; Tabl1[28].Row:=50; Tabl1[28].Value:=sGrid1.Cells[6,50];
Tabl1[29].Col:=6; Tabl1[29].Row:=51; Tabl1[29].Value:=sGrid1.Cells[6,51];
Tabl1[30].Col:=6; Tabl1[30].Row:=52; Tabl1[30].Value:=sGrid1.Cells[6,52];
Tabl1[31].Col:=6; Tabl1[31].Row:=53; Tabl1[31].Value:=sGrid1.Cells[6,53];
Tabl1[32].Col:=6; Tabl1[32].Row:=54; Tabl1[32].Value:=sGrid1.Cells[6,54];
Tabl1[33].Col:=6; Tabl1[33].Row:=59; Tabl1[33].Value:=sGrid1.Cells[6,59];
Tabl1[34].Col:=6; Tabl1[34].Row:=60; Tabl1[34].Value:=sGrid1.Cells[6,60];
Tabl1[35].Col:=6; Tabl1[35].Row:=61; Tabl1[35].Value:=sGrid1.Cells[6,61];
Tabl1[36].Col:=6; Tabl1[36].Row:=62; Tabl1[36].Value:=sGrid1.Cells[6,62];
Tabl1[37].Col:=6; Tabl1[37].Row:=63; Tabl1[37].Value:=sGrid1.Cells[6,63];
Tabl1[38].Col:=6; Tabl1[38].Row:=64; Tabl1[38].Value:=sGrid1.Cells[6,64];
Tabl1[39].Col:=6; Tabl1[39].Row:=65; Tabl1[39].Value:=sGrid1.Cells[6,65];
Tabl1[40].Col:=6; Tabl1[40].Row:=66; Tabl1[40].Value:=sGrid1.Cells[6,66];
Tabl1[41].Col:=6; Tabl1[41].Row:=67; Tabl1[41].Value:=sGrid1.Cells[6,67];
Tabl1[42].Col:=6; Tabl1[42].Row:=68; Tabl1[42].Value:=sGrid1.Cells[6,68];
Tabl1[43].Col:=7; Tabl1[43].Row:=15; Tabl1[43].Value:=sGrid1.Cells[7,15];
Tabl1[44].Col:=7; Tabl1[44].Row:=16; Tabl1[44].Value:=sGrid1.Cells[7,16];
Tabl1[45].Col:=7; Tabl1[45].Row:=17; Tabl1[45].Value:=sGrid1.Cells[7,17];
Tabl1[46].Col:=7; Tabl1[46].Row:=18; Tabl1[46].Value:=sGrid1.Cells[7,18];
Tabl1[47].Col:=7; Tabl1[47].Row:=19; Tabl1[47].Value:=sGrid1.Cells[7,19];
Tabl1[48].Col:=7; Tabl1[48].Row:=20; Tabl1[48].Value:=sGrid1.Cells[7,20];
Tabl1[49].Col:=7; Tabl1[49].Row:=21; Tabl1[49].Value:=sGrid1.Cells[7,21];
Tabl1[50].Col:=7; Tabl1[50].Row:=22; Tabl1[50].Value:=sGrid1.Cells[7,22];
Tabl1[51].Col:=7; Tabl1[51].Row:=23; Tabl1[51].Value:=sGrid1.Cells[7,23];
Tabl1[52].Col:=7; Tabl1[52].Row:=24; Tabl1[52].Value:=sGrid1.Cells[7,24];
Tabl1[53].Col:=7; Tabl1[53].Row:=25; Tabl1[53].Value:=sGrid1.Cells[7,25];
Tabl1[54].Col:=7; Tabl1[54].Row:=26; Tabl1[54].Value:=sGrid1.Cells[7,26];
Tabl1[55].Col:=7; Tabl1[55].Row:=27; Tabl1[55].Value:=sGrid1.Cells[7,27];
Tabl1[56].Col:=7; Tabl1[56].Row:=28; Tabl1[56].Value:=sGrid1.Cells[7,28];
Tabl1[57].Col:=7; Tabl1[57].Row:=35; Tabl1[57].Value:=sGrid1.Cells[7,35];
Tabl1[58].Col:=7; Tabl1[58].Row:=36; Tabl1[58].Value:=sGrid1.Cells[7,36];
Tabl1[59].Col:=7; Tabl1[59].Row:=37; Tabl1[59].Value:=sGrid1.Cells[7,37];
Tabl1[60].Col:=7; Tabl1[60].Row:=38; Tabl1[60].Value:=sGrid1.Cells[7,38];
Tabl1[61].Col:=7; Tabl1[61].Row:=39; Tabl1[61].Value:=sGrid1.Cells[7,39];
Tabl1[62].Col:=7; Tabl1[62].Row:=40; Tabl1[62].Value:=sGrid1.Cells[7,40];
Tabl1[63].Col:=7; Tabl1[63].Row:=41; Tabl1[63].Value:=sGrid1.Cells[7,41];
Tabl1[64].Col:=7; Tabl1[64].Row:=42; Tabl1[64].Value:=sGrid1.Cells[7,42];
Tabl1[65].Col:=7; Tabl1[65].Row:=43; Tabl1[65].Value:=sGrid1.Cells[7,43];
Tabl1[66].Col:=7; Tabl1[66].Row:=44; Tabl1[66].Value:=sGrid1.Cells[7,44];
Tabl1[67].Col:=7; Tabl1[67].Row:=45; Tabl1[67].Value:=sGrid1.Cells[7,45];
Tabl1[68].Col:=7; Tabl1[68].Row:=46; Tabl1[68].Value:=sGrid1.Cells[7,46];
Tabl1[69].Col:=7; Tabl1[69].Row:=47; Tabl1[69].Value:=sGrid1.Cells[7,47];
Tabl1[70].Col:=7; Tabl1[70].Row:=49; Tabl1[70].Value:=sGrid1.Cells[7,49];
Tabl1[71].Col:=7; Tabl1[71].Row:=50; Tabl1[71].Value:=sGrid1.Cells[7,50];
Tabl1[72].Col:=7; Tabl1[72].Row:=51; Tabl1[72].Value:=sGrid1.Cells[7,51];
Tabl1[73].Col:=7; Tabl1[73].Row:=52; Tabl1[73].Value:=sGrid1.Cells[7,52];
Tabl1[74].Col:=7; Tabl1[74].Row:=53; Tabl1[74].Value:=sGrid1.Cells[7,53];
Tabl1[75].Col:=7; Tabl1[75].Row:=54; Tabl1[75].Value:=sGrid1.Cells[7,54];
Tabl1[76].Col:=7; Tabl1[76].Row:=59; Tabl1[76].Value:=sGrid1.Cells[7,59];
Tabl1[77].Col:=7; Tabl1[77].Row:=60; Tabl1[77].Value:=sGrid1.Cells[7,60];
Tabl1[78].Col:=7; Tabl1[78].Row:=61; Tabl1[78].Value:=sGrid1.Cells[7,61];
Tabl1[79].Col:=7; Tabl1[79].Row:=62; Tabl1[79].Value:=sGrid1.Cells[7,62];
Tabl1[80].Col:=7; Tabl1[80].Row:=63; Tabl1[80].Value:=sGrid1.Cells[7,63];
Tabl1[81].Col:=7; Tabl1[81].Row:=64; Tabl1[81].Value:=sGrid1.Cells[7,64];
Tabl1[82].Col:=7; Tabl1[82].Row:=65; Tabl1[82].Value:=sGrid1.Cells[7,65];
Tabl1[83].Col:=7; Tabl1[83].Row:=66; Tabl1[83].Value:=sGrid1.Cells[7,66];
Tabl1[84].Col:=7; Tabl1[84].Row:=67; Tabl1[84].Value:=sGrid1.Cells[7,67];
Tabl1[85].Col:=7; Tabl1[85].Row:=68; Tabl1[85].Value:=sGrid1.Cells[7,68];
Tabl1[86].Col:=11; Tabl1[86].Row:=15; Tabl1[86].Value:=sGrid1.Cells[8,15];
Tabl1[87].Col:=11; Tabl1[87].Row:=16; Tabl1[87].Value:=sGrid1.Cells[11,16];
Tabl1[88].Col:=11; Tabl1[88].Row:=17; Tabl1[88].Value:=sGrid1.Cells[11,17];
Tabl1[89].Col:=11; Tabl1[89].Row:=18; Tabl1[89].Value:=sGrid1.Cells[11,18];
Tabl1[90].Col:=11; Tabl1[90].Row:=19; Tabl1[90].Value:=sGrid1.Cells[11,19];
Tabl1[91].Col:=11; Tabl1[91].Row:=20; Tabl1[91].Value:=sGrid1.Cells[11,20];
Tabl1[92].Col:=11; Tabl1[92].Row:=21; Tabl1[92].Value:=sGrid1.Cells[11,21];
Tabl1[93].Col:=11; Tabl1[93].Row:=22; Tabl1[93].Value:=sGrid1.Cells[11,22];
Tabl1[94].Col:=11; Tabl1[94].Row:=23; Tabl1[94].Value:=sGrid1.Cells[11,23];
Tabl1[95].Col:=11; Tabl1[95].Row:=24; Tabl1[95].Value:=sGrid1.Cells[11,24];
Tabl1[96].Col:=11; Tabl1[96].Row:=25; Tabl1[96].Value:=sGrid1.Cells[11,25];
Tabl1[97].Col:=11; Tabl1[97].Row:=26; Tabl1[97].Value:=sGrid1.Cells[11,26];
Tabl1[98].Col:=11; Tabl1[98].Row:=27; Tabl1[98].Value:=sGrid1.Cells[11,27];
Tabl1[99].Col:=11; Tabl1[99].Row:=28; Tabl1[99].Value:=sGrid1.Cells[11,28];
Tabl1[100].Col:=11; Tabl1[100].Row:=35; Tabl1[100].Value:=sGrid1.Cells[11,35];
Tabl1[101].Col:=11; Tabl1[101].Row:=36; Tabl1[101].Value:=sGrid1.Cells[11,36];
Tabl1[102].Col:=11; Tabl1[102].Row:=37; Tabl1[102].Value:=sGrid1.Cells[11,37];
Tabl1[103].Col:=11; Tabl1[103].Row:=38; Tabl1[103].Value:=sGrid1.Cells[11,38];
Tabl1[104].Col:=11; Tabl1[104].Row:=39; Tabl1[104].Value:=sGrid1.Cells[11,39];
Tabl1[105].Col:=11; Tabl1[105].Row:=40; Tabl1[105].Value:=sGrid1.Cells[11,40];
Tabl1[106].Col:=11; Tabl1[106].Row:=41; Tabl1[106].Value:=sGrid1.Cells[11,41];
Tabl1[107].Col:=11; Tabl1[107].Row:=42; Tabl1[107].Value:=sGrid1.Cells[11,42];
Tabl1[108].Col:=11; Tabl1[108].Row:=43; Tabl1[108].Value:=sGrid1.Cells[11,43];
Tabl1[109].Col:=11; Tabl1[109].Row:=44; Tabl1[109].Value:=sGrid1.Cells[11,44];
Tabl1[110].Col:=11; Tabl1[110].Row:=45; Tabl1[110].Value:=sGrid1.Cells[11,45];
Tabl1[111].Col:=11; Tabl1[111].Row:=46; Tabl1[111].Value:=sGrid1.Cells[11,46];
Tabl1[112].Col:=11; Tabl1[112].Row:=47; Tabl1[112].Value:=sGrid1.Cells[11,47];
Tabl1[113].Col:=11; Tabl1[113].Row:=49; Tabl1[113].Value:=sGrid1.Cells[11,49];
Tabl1[114].Col:=11; Tabl1[114].Row:=50; Tabl1[114].Value:=sGrid1.Cells[11,50];
Tabl1[115].Col:=11; Tabl1[115].Row:=51; Tabl1[115].Value:=sGrid1.Cells[11,51];
Tabl1[116].Col:=11; Tabl1[116].Row:=52; Tabl1[116].Value:=sGrid1.Cells[11,52];
Tabl1[117].Col:=11; Tabl1[117].Row:=53; Tabl1[117].Value:=sGrid1.Cells[11,53];
Tabl1[118].Col:=11; Tabl1[118].Row:=54; Tabl1[118].Value:=sGrid1.Cells[11,54];
Tabl1[119].Col:=11; Tabl1[119].Row:=59; Tabl1[119].Value:=sGrid1.Cells[11,59];
Tabl1[120].Col:=11; Tabl1[120].Row:=60; Tabl1[120].Value:=sGrid1.Cells[11,60];
Tabl1[121].Col:=11; Tabl1[121].Row:=61; Tabl1[121].Value:=sGrid1.Cells[11,61];
Tabl1[122].Col:=11; Tabl1[122].Row:=62; Tabl1[122].Value:=sGrid1.Cells[11,62];
Tabl1[123].Col:=11; Tabl1[123].Row:=63; Tabl1[123].Value:=sGrid1.Cells[11,63];
Tabl1[124].Col:=11; Tabl1[124].Row:=64; Tabl1[124].Value:=sGrid1.Cells[11,64];
Tabl1[125].Col:=11; Tabl1[125].Row:=65; Tabl1[125].Value:=sGrid1.Cells[11,65];
Tabl1[126].Col:=11; Tabl1[126].Row:=66; Tabl1[126].Value:=sGrid1.Cells[11,66];
Tabl1[127].Col:=11; Tabl1[127].Row:=67; Tabl1[127].Value:=sGrid1.Cells[11,67];
Tabl1[128].Col:=11; Tabl1[128].Row:=68; Tabl1[128].Value:=sGrid1.Cells[11,68];
end;

procedure TBlok1_Form.CalculateBtn2Click(Sender: TObject);
begin
 with TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))) do
 for i:=FixedCols to ColCount-1 do
 for j:=FixedRows to RowCount-1 do
  Cells[i, j]:='';
  if FileExists('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml') then  //проверяем существует ли файл с таблицей
  begin
   TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).LoadFromFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml'); //считываем данные из файла Tablica2.xml
   AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1)))); // выравниваем ячейки по ширине, в зависимости от длины текста
  end
  else begin   //если файл не существует, выводим сообщение об ошибке
    NotFoundException(Blok1_Form.PageControl1.TabIndex+1);
    Exit; //выходим из процедуры считывания данных
  end;
 AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));  //выравниваем ячейки по ширине, в зависимости от длины текста
end;

procedure TBlok1_Form.CalculateBtn3Click(Sender: TObject);
begin
 with TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))) do
 for i:=FixedCols to ColCount-1 do
 for j:=FixedRows to RowCount-1 do
  Cells[i, j]:='';
 if FileExists('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml') then  //проверяем существует ли файл с таблицей
 begin
  TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).LoadFromFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml'); //считываем данные из файла Tablica3.xml
  AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1)))); // выравниваем ячейки по ширине, в зависимости от длины текста
 end
 else begin   //если файл не существует, выводим сообщение об ошибке
  NotFoundException(Blok1_Form.PageControl1.TabIndex+1);
  Exit; //выходим из процедуры считывания данных
 end;
 AutoFit(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));  //выравниваем ячейки по ширине, в зависимости от длины текста
end;

procedure TBlok1_Form.BackBtn1Click(Sender: TObject);
begin
 WelcomeForm.WindowState:=wsNormal; //разворачиваем форму "Добро пожаловать"
 WelcomeForm.Visible:=True; //показываем форму "Добро пожаловать"
 Blok1_Form.Hide(); //скрываем форму с таблицей
 Application.Title:=WelcomeForm.Caption; //меняем заголовок программы на панели задач
end;

procedure TBlok1_Form.FormClose(Sender: TObject);
begin
 WelcomeForm.WindowState:=wsNormal; //разворачиваем форму "Добро пожаловать"
 WelcomeForm.Visible:=True; //показываем форму "Добро пожаловать"
 Application.Title:=WelcomeForm.Caption; //меняем заголовок программы на панели задач
end;

procedure TBlok1_Form.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
 if MessageDlg('Подтверждение', 'Хотите закрыть окно ?', mtConfirmation, [mbYes, mbNO] , 0) = mrYes then
 begin
  CanClose:=True;
  Blok1_Form.Close;
 end
 else
  CanClose:=False;
end;

procedure TBlok1_Form.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (key = VK_F1) then WelcomeForm.MenuItem5.Click //при нажатии клавиши F1 вызываем справку
 else if (key = VK_ESCAPE) then  //при нажатии клавиши ESC спрашиваем пользователя хочет ли он выйти
  begin
   if MessageDlg('Подтверждение', 'Хотите закрыть окно ?', mtConfirmation, [mbYes, mbNO] , 0) = mrYes then
    Blok1_Form.Close;
  end;
end;

procedure TBlok1_Form.FormResize(Sender: TObject);
begin //при изменении размеров окна выполняем повторное масштабирование элементов
 PageControl1.Width:=Width;
 PageControl1.Height:=Height;
 TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Height:=Height-105;
 TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Width:=Width-8;
 For i:=1 to s do
 Begin
  TLabel(FindComponent('Label' + IntToStr(i))).Width:=Blok1_Form.Width;
  TLabel(FindComponent('Label' + IntToStr(i))).Alignment:=taCenter; //разместим название таблицы в центре формы
 End;
end;

procedure TBlok1_Form.PageControl1Change(Sender: TObject);
begin //при переключении вкладок выполняем масштабирование элементов в соответствии с размером окна
 PageControl1.Width:=Width;
 PageControl1.Height:=Height;
 TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Height:=Height-105;
 TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Width:=Width-8;
 Blok1_Form.Caption:='МкАРС 1.5 :: '+'Таблица №'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+' '+TTabsheet(FindComponent('TabSheet' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Caption; //меняем заголовок формы при переключении вкладки
 Application.Title:=Blok1_Form.Caption; //меняем надпись на панели задач при переключении вкладки
end;

//Процедура, выполняющая выравнивание названия таблицы по левому краю
procedure TBlok1_Form.sGrid1DrawCell(Sender: TObject; aCol, aRow: Integer;
 aRect: TRect; aState: TGridDrawState);
var SG1:TStringGrid; Stroka1:string; Rect1,Rect2:TRect; xFilled:Boolean; a:integer;
begin
 SG1:=TStringGrid(Sender);
 with SG1 do
  xFilled:=True;
  SG1.Canvas.Brush.Color:=clBlack;
  if (ARow=0) and (ACol in [0..SG1.ColCount-1]) then Begin
   Rect1:=SG1.CellRect(0,0);
   Rect2:=SG1.CellRect(SG1.ColCount-1,0);
   aRect:=Classes.Rect(Rect1.Left,Rect1.Top,Rect2.Right,Rect2.Bottom);
   Stroka1:='           '+SG1.Cells[0,0];
   Stroka1:=UTF8ToSys(Stroka1);
  end
  else if (ARow = 1) and (ACol in [0..SG1.ColCount-1]) then Begin
   Rect1:=SG1.CellRect(0,1);
   Rect2:=SG1.CellRect(SG1.ColCount-1,1);
   aRect:=Classes.Rect(Rect1.Left,Rect1.Top,Rect2.Right,Rect2.Bottom);
   Stroka1:='           '+SG1.Cells[0,1];
   Stroka1:=UTF8ToSys(Stroka1);
  end
  Else Begin
  xFilled:=False;
  end;
  SG1.Canvas.FrameRect(Classes.Rect(aRect.Left-1,aRect.Top-1,aRect.Right+1,aRect.Bottom+1));
  if xFilled Then Begin
   SG1.Canvas.Brush.Color:=clWindow; //закрашиваем две верхних строки в такой же цвет, как и у ячейки
   SG1.Canvas.FillRect(aRect);
   SG1.Canvas.Font.Name:='Times New Roman'; //шрифт Times New Roman
   SG1.Canvas.Font.Bold:=True; //Жирный
  end
  else SG1.Canvas.Brush.Color:=clWhite;
  DrawText(SG1.Canvas.Handle,PChar(Stroka1), Length(Stroka1), aRect, DT_LEFT);
  Stroka1:='';

  for a:=0 to Length(Tabl1)-1 do begin
   If (SG1=sGrid1) AND (ACol = Tabl1[a].Col) AND (ARow = Tabl1[a].Row)  then
   Begin
    SG1.Canvas.Brush.Color:=$00DCF8FF; //телесный цвет
    SG1.Canvas.FillRect(aRect);
    SG1.Canvas.TextOut(aRect.Left,aRect.Top,SG1.Cells[Acol,Arow]);
   End;
  end;
end;

procedure TBlok1_Form.sGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
 myRect: TGridRect; SG1:TStringGrid;
begin
 SG1:=TStringGrid(Sender);     // при нажатии CTRL + A выделяем всю таблицу
 with SG1 do
  if (Key = Ord('A')) and (ssCtrl in Shift) then
   begin
    myRect.Left := 0;
    myRect.Top := 0;
    myRect.Right := SG1.ColCount-1;
    myRect.Bottom := SG1.rowcount-1;
    SG1.Selection := myRect;
   end;
end;

procedure TBlok1_Form.sGrid1KeyPress(Sender: TObject; var Key: char);
//Процедура, ограничивающая ввод данных в ячейку - допускаются только цифры, знаки +,- и дробный разделитель (. или ,)
const
 AllowedChars: string = '1234567890+-,';
var
 i: Integer; Ok: Boolean; SG1:TStringGrid;
begin
 SG1:=TStringGrid(Sender);
 with SG1 do
 i:= 0;
 Ok := False;
 if Key = #8 then Ok := True;
 repeat
  i := i + 1;
  if (Key = AllowedChars[i]) or (Key = #13) or (Key = #27) then Ok := True;
  //Если была нажата клавиша Enter, ESCAPE или были введены допустимые символы, то обрабатываем нажатие клавиши
  until (Ok) or (i = Length(AllowedChars));
  if not Ok then begin  //Иначе - блокируем ввод в ячейку
  Key := #0; end; end;

procedure TBlok1_Form.CancelBtn1Click(Sender: TObject);
begin
RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_1.xml', 'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_9.xml');
RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml' ,'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_1.xml');
RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_9.xml', 'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml');
TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Enabled:=FALSE;
x:=0;
x:=Blok1_Form.PageControl1.TabIndex; 
Blok1_Form.Hide;
//производим расчёты во всех 3 таблицах
Blok1_Form.PageControl1.TabIndex:=0; //первая таблица
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(1))).Click;
Blok1_Form.PageControl1.TabIndex:=1; //вторая таблица
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(2))).Click;
Blok1_Form.PageControl1.TabIndex:=2; //третья таблица
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(3))).Click;
Blok1_Form.PageControl1.TabIndex:=x;
Blok1_Form.Show;
TButton(FindComponent('CalculateBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Click();
end;

procedure TBlok1_Form.ExportToExcelBtn1Click(Sender: TObject);
begin
 //Экспорт таблицы в Excel
 ExportToExcel(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));
end;

procedure TBlok1_Form.ExportToOOCalcBtn1Click(Sender: TObject);
begin
 //Экспорт таблицы в OpenOffice
 ExportToOpenOffice(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));
end;

procedure TBlok1_Form.ExportToXLSBtn1Click(Sender: TObject);
begin
 //Экспорт таблицы в файл
 ExportToFile(TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))), SaveDialog1, TTabsheet(FindComponent('TabSheet' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))));
end;

procedure TBlok1_Form.FormCreate(Sender: TObject);
//Считываем из файла MKARS.ini в секции Blok1_Zagolovki заголовки для всех вкладок данной формы и присваиваем каждой вкладке свой заголовок
begin
initValues;
for i:=1 to s do
begin
 tempstring:='';
 tempstring:='Table'+IntToStr(i); //формируем имя вкладки
 TLabel(FindComponent('Label' + IntToStr(i))).Caption:=SysToUTF8(IniF.ReadString('Blok1',tempstring,''));
 TLabel(FindComponent('Label' + IntToStr(i))).Width:=Blok1_Form.Width;
 TLabel(FindComponent('Label' + IntToStr(i))).Alignment:=taCenter; //разместим название таблицы в центре формы
 TTabsheet(FindComponent('TabSheet' + IntToStr(i))).Caption:=SysToUTF8(IniF.ReadString('Blok1_Zagolovki',tempstring,'')); //имя вкладки
 TTabsheet(FindComponent('TabSheet' + IntToStr(i))).ShowHint:=False;
end;
DecimalSeparator := ',';
WindowState:=wsMaximized;
//Проверяем, установлен ли у пользователя Excel или OpenOffice
//Если MS Excel установлен, то включаем кнопку "Экспорт в Excel" во всех вкладках
if not IsOLEObjectInstalled('Excel.Application') then
begin
i:=0;
for i:=1 to s do
Begin
TButton(FindComponent('ExportToExcelBtn' + IntToStr(i))).Hint:='К сожалению, на вашем компьютере не установлен Microsoft Office, поэтому экспорт таблицы в Excel невозможен !';
End;
end
else
begin
i:=0;
for i:=1 to s do
Begin
TButton(FindComponent('ExportToExcelBtn' + IntToStr(i))).Enabled:=TRUE;
TButton(FindComponent('ExportToExcelBtn' + IntToStr(i))).ShowHint:=FALSE;
End;
end;
//Если OpenOffice установлен, то включаем кнопку "Экспорт в OpenOffice" во всех вкладках
if not IsOLEObjectInstalled('com.sun.star.ServiceManager') then
begin
i:=0;
for i:=1 to s do
Begin
TButton(FindComponent('ExportToOOCalcBtn' + IntToStr(i))).Hint:='К сожалению, на вашем компьютере не установлен OpenOffice, поэтому экспорт таблицы в OpenOffice Calc невозможен !';
End;
end
else
begin
i:=0;
for i:=1 to s do
Begin
TButton(FindComponent('ExportToOOCalcBtn' + IntToStr(i))).Enabled:=TRUE;
TButton(FindComponent('ExportToOOCalcBtn' + IntToStr(i))).ShowHint:=FALSE;
End;
end;
end;

procedure TBlok1_Form.SaveChangesBtn1Click(Sender: TObject);
begin
 for i:=0 to Length(Tabl1)-1 do begin
  Tabl1[i].Value := sGrid1.Cells[Tabl1[i].Col, Tabl1[i].Row]; //запишем сохраненные значения
 end;
 for i:=sGrid1.FixedCols to sGrid1.ColCount-1 do
 for j:=sGrid1.FixedRows to sGrid1.RowCount-1 do
  sGrid1.Cells[i, j]:='';
 TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).LoadFromFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml');
 for i:=0 to  Length(Tabl1)-1 do begin
  sGrid1.Cells[Tabl1[i].Col, Tabl1[i].Row] := Tabl1[i].Value;
 end;
 SaveAndRefresh(1);
end;

procedure TBlok1_Form.SaveAndRefresh(Num: Integer);
var TableNumbers: array [0..2] of integer; tbl:integer;
begin
 Blok1_Form.Hide;  //TableNumbers - массив с номерами таблиц, он нам нужен для корректного перерасчета, если таблицы связаны не по порядку
 //например 6 таблица использует формулы из 21, а значит нам надо сперва расчитать 21 таблицу, а затем 6
 //Пример: TableNumbers[0]:=6; TableNumbers[1]:=21;
 //Длина массива TableNumbers соответствует числу таблиц, при этом нумерация начинается с 0
 TableNumbers[0]:=1;
 TableNumbers[1]:=2;
 TableNumbers[2]:=3;
 try
  TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Enabled:=True;
  TStringGrid(FindComponent('sGrid' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).SaveToFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_1.xml');
  RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_1.xml', 'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_3.xml');
  RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml' ,'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_1.xml');
  RenameFile('Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'_3.xml', 'Data\Tablica'+IntToStr(Blok1_Form.PageControl1.TabIndex+1)+'.xml');
  TButton(FindComponent('CalculateBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Click();
  tbl:=0;
  for tbl:=0 to Length(TableNumbers)-1 do
   begin
    if TableNumbers[tbl] <> Num then
     begin
      Blok1_Form.PageControl1.TabIndex:=TableNumbers[tbl]-1;
      TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(TableNumbers[tbl]))).Click;
     end
    else Continue;
   end;
  Blok1_Form.Show;
  Blok1_Form.PageControl1.TabIndex:=Num-1;
except
 on EZeroDivide do
 begin
 MessageDlg('МкАРС 1.5 :: Ошибка при сохранении изменений', 'Произошла ошибка связанная с попыткой деления на ноль!'+#13#10+'Внесённые вами изменения будут отменены.',mtError, [mbOk], 0);
   Blok1_Form.PageControl1.TabIndex:=Num-1;
   If (TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Enabled=True) then
     TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Click();
 end;

 on EOverflow do
 begin
 MessageDlg('МкАРС 1.5 :: Ошибка при сохранении изменений', 'Произошла ошибка переполнения! Были введены слишком большие числа'+#13#10+'Внесённые вами изменения будут отменены.',mtError, [mbOk], 0);
   Blok1_Form.PageControl1.TabIndex:=Num-1;
   If (TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Enabled=True) then
     TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Click();
 end;
 on EConvertError do
 begin
 MessageDlg('МкАРС 1.5 :: Ошибка при сохранении изменений', 'Произошла ошибка при расчёте таблицы!'+#13#10+'Внесённые вами изменения будут отменены.',mtError, [mbOk], 0);
   Blok1_Form.PageControl1.TabIndex:=Num-1;
   If (TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Enabled=True) then
     TButton(FindComponent('CancelBtn' + IntToStr(Blok1_Form.PageControl1.TabIndex+1))).Click();
 end;
 on EFOpenError do
 begin
   Blok1_Form.PageControl1.TabIndex:=Num-1;
 MessageDlg('МкАРС 1.5 :: Ошибка при сохранении изменений', 'Произошла ошибка при открытии файла! Возможно он был удалён или же его никогда не существовало.',mtError, [mbOk], 0);
 end;
 end;
end;

end.
