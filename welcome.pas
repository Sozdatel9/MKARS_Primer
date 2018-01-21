unit welcome;
{$mode objfpc}{$H+}

interface

uses
Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
ComObj, Grids, Menus, Windows, ShellApi, INIFiles, LResources, Translations;
//подключаем компоненты для русификации кнопок и диалоговых окон
//при компиляции проекта в среде разработки Lazarus - модули LResources и Translations

type
{ TWelcomeForm }

TWelcomeForm = class(TForm)
block1lbl: TLabel;
GroupBox1: TGroupBox;
MainMenu1: TMainMenu;
MenuItem1: TMenuItem;
MenuItem2: TMenuItem;
MenuItem3: TMenuItem;
MenuItem4: TMenuItem;
MenuItem5: TMenuItem;
MenuItem6: TMenuItem;
MenuItem7: TMenuItem;
OkButton1: TButton;
Label1: TLabel;
selectedTableTitle: TLabel;
TableList1: TComboBox;
procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
procedure FormCreate(Sender: TObject);
procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
procedure FormResize(Sender: TObject);
procedure MenuItem2Click(Sender: TObject);
procedure MenuItem5Click(Sender: TObject);
procedure MenuItem6Click(Sender: TObject);
procedure MenuItem7Click(Sender: TObject);
procedure OkButton1Click(Sender: TObject);
procedure TableList1Change(Sender: TObject);
procedure TableList1DropDown(Sender: TObject);
procedure TableList1Enter(Sender: TObject);
procedure TableList1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
procedure TableList1MouseEnter(Sender: TObject);
procedure TableList1Select(Sender: TObject);
//procedure ShowTableDescription(Num: Integer);
private
{ private declarations }
public
{ public declarations }
end;

var
WelcomeForm: TWelcomeForm;
selectedTable1: string;
selectedTable2: string;
selectedTable3: string;
tempstring :string;
IniF :TINIFile;
s,s1,s2,i: integer;
implementation

uses aboutform, backupform, blok1;

{$R *.lfm}

{ TWelcomeForm }

//Функция, предназначенная для русификации стандартных надписей форм и кнопок
//при компиляции проекта в среде разработки Lazarus
function TranslateUnitResourceStrings: boolean;
var
  r: TLResource;
  POFile: TPOFile;
begin
  r:=LazarusResources.Find('lclstrconsts.ru','PO');
  POFile:=TPOFile.Create;
  try
    POFile.ReadPOText(r.Value);
    Result:=Translations.TranslateUnitResourceStrings('lclstrconsts',POFile);
  finally
    FreeAndNil(POFile);
  end;
end;

procedure TWelcomeForm.FormCreate(Sender: TObject);
begin
TranslateUnitResourceStrings;     //заменяем стандартные надписи кнопок и форм на русские
DecimalSeparator := ',';
//проверяем, существует ли файл MKARS.ini
if FileExists('MKARS.ini') then  //если существует, то загружаем из него названия таблиц, чтобы их мог выбрать пользователь
begin
IniF := TINIFile.Create('MKARS.ini'); //(вместо "Yes" - "Да", вместо "Cancel" - "Отмена" и т.д.
//подготавливаем выпадающий список, загружаем в него названия таблиц.
//Названия таблиц считываются из файла MKARS.ini
//Блок 1
s:=Inif.ReadInteger('Main', 'ChisloTablic_1', 30); //считываем количество таблиц, если не удалось считать, то по-умолчанию 30
i:=0;
for i:=1 to s do
begin
  tempstring:='Table'+IntToStr(i);
  TableList1.Items.Add(SysToUTF8(IniF.ReadString('Blok1',tempstring,'')));
end;
TableList1.ItemIndex:=0;
Application.Title:=WelcomeForm.Caption;
selectedTableTitle.Caption:=TableList1.Items[TableList1.ItemIndex];
end
else begin   //если файл не существует, выводим сообщение об ошибке
  WelcomeForm.Visible:=False;
  MessageDlg('МкАРС 1.5 :: Критическая ошибка', 'ОШИБКА:  Файл с информацией о таблицах MKARS.ini не найден !' + #13 + #13 + 'Программа будет закрыта.',mtError, [mbOk], 0);
  //Application.Terminate;  //выходим из программы
  Halt(0); //принудительно закрываем программу
end;
end;

procedure TWelcomeForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//при нажатии комбинации клавиш SHIFT + F1 показываем окно "О программе"
 If (Key = VK_F1) and (ssShift in Shift) then WelcomeForm.MenuItem6.Click
//при нажатии клавиши F1 вызываем справку
 Else if (Key = VK_F1) then
   begin
    WelcomeForm.MenuItem5.Click;
   end
//при нажатии комбинации клавиш CTRL + R показываем окно "Резервное копирование"
 Else if (Key = Ord('R')) and (ssCtrl in Shift) then
   begin
    WelcomeForm.MenuItem7.Click;
   end
//при нажатии клавиш ESC показываем окно "Выход из программы"
Else if (Key = VK_ESCAPE) then
  begin
   WelcomeForm.MenuItem2.Click;
  end;
end;

procedure TWelcomeForm.FormResize(Sender: TObject);
//при изменении размеров окна меняем ширину элементов
begin
 TableList1.Width:=Width-45;
 OkButton1.Left:=Width-38;
 GroupBox1.Width:=Width-9;
end;

procedure TWelcomeForm.FormCloseQuery(Sender: TObject; var CanClose: boolean);
var selectedBtn: Integer; //в этой переменной хранится нажатие выбранной пользователем кнопки
begin
selectedBtn:= MessageDlg('МкАРС 1.5 :: Выход из программы', 'Вы собираетесь выйти из программы.' + #13 + #13 + 'Выполнить резервное копирование данных прямо сейчас ?', mtConfirmation, [mbYes, mbNO, mbCancel] , 0);
if selectedBtn = mrYes then //Если нажата кнопка Да, то сперва выполняем резервное копирование данных, а затем выходим из программы
begin
BackupMainForm.Show();
BackupMainForm.CreateBackupBtn.Click();
BackupMainForm.Hide();
IniF.Free;
CanClose:=True;
Application.Terminate();
end
else if selectedBtn=mrNo then begin IniF.Free; CanClose:=True; Application.Terminate() end //Если нажата кнопка Нет, то сразу выходим из программы
else CanClose:=FALSE; //Если нажата кнопка "Отмена", то ничего не делаем
end;

procedure TWelcomeForm.MenuItem2Click(Sender: TObject);
begin
WelcomeForm.Close(); //выходим из программы
end;
procedure TWelcomeForm.MenuItem5Click(Sender: TObject);
begin
//проверяем, существует ли файл README.txt
if FileExists('README.txt') then  //если существует, то открываем его с помощью программы по умолчанию (Блокнот, Notepad++ и т.д).
begin
//открываем файл со справочной информацией README.txt
ShellExecute(0, 'open', 'README.txt', nil, nil, SW_SHOW);
end
else begin   //если файл не существует, выводим сообщение об ошибке
  MessageDlg('МкАРС 1.5 :: Ошибка', 'ОШИБКА !'+#13+#13+'Файл справки README.txt не найден !',mtError, [mbOk], 0);
end;
end;

procedure TWelcomeForm.MenuItem6Click(Sender: TObject);
begin
 AboutMainForm.Show(); //Показываем окно "О программе"
end;

procedure TWelcomeForm.MenuItem7Click(Sender: TObject);
begin
 BackupMainForm.Show(); //Показываем окно "Резервное копирование"
end;
procedure TWelcomeForm.OkButton1Click(Sender: TObject);
begin
DecimalSeparator := ','; //разделитель целой и дробной части числа
//Блок 1
selectedTable1:=TableList1.Items[TableList1.ItemIndex]; //Блок 1
tempstring:='';
tempstring:='Table'+IntToStr(TableList1.ItemIndex+1);   //считываем номер выбранной пользователем таблицы из списка
//переключаем нужную вкладку форму в зависимости от выбранной из выпадающего списка таблицы
if selectedTable1=SysToUTF8(IniF.ReadString('Blok1',tempstring,'')) then
begin
//переходим на вкладку, соответствующую выбранной из выпадающего списка таблице
Blok1_Form.PageControl1.TabIndex:=TableList1.ItemIndex;
//меняем заголовок программы на панели задач
Application.Title:=Blok1_Form.Caption;
//производим расчёты во всех 29 таблицах
Blok1_Form.PageControl1.TabIndex:=0;
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(1))).Click;
Blok1_Form.PageControl1.TabIndex:=1;
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(2))).Click;
Blok1_Form.PageControl1.TabIndex:=2;
TButton(Blok1_Form.FindComponent('CalculateBtn' + IntToStr(3))).Click;
Blok1_Form.PageControl1.TabIndex:=3;
//переходим на вкладку, соответствующую выбранной из выпадающего списка таблице
Blok1_Form.PageControl1.TabIndex:=TableList1.ItemIndex;
//показываем форму
Blok1_Form.Show;
WelcomeForm.Visible:=False; //скрываем форму "Добро пожаловать"
end;
end;

procedure ShowTableDescription(Num: Integer);
begin
 //Выводим название выбранной таблицы при наведении курсора, в соответствии с номером
 case Num of
  1 : begin
       WelcomeForm.TableList1.Hint := WelcomeForm.TableList1.Items[WelcomeForm.TableList1.ItemIndex];
       WelcomeForm.selectedTableTitle.Caption := WelcomeForm.TableList1.Items[WelcomeForm.TableList1.ItemIndex];
      end;
 end;
end;

procedure TWelcomeForm.TableList1Change(Sender: TObject);
begin
 ShowTableDescription(1);
end;

procedure TWelcomeForm.TableList1DropDown(Sender: TObject);
begin
 ShowTableDescription(1);
end;

procedure TWelcomeForm.TableList1Enter(Sender: TObject);
begin
 ShowTableDescription(1);
end;

procedure TWelcomeForm.TableList1MouseEnter(Sender: TObject);
begin
 ShowTableDescription(1);
end;

procedure TWelcomeForm.TableList1Select(Sender: TObject);
begin
 ShowTableDescription(1);
end;

procedure TWelcomeForm.TableList1KeyDown(Sender: TObject; var Key: Word;
 Shift: TShiftState);
begin
 if (Key = VK_RETURN) then
  begin
   WelcomeForm.OkButton1.Click;  //при нажатии Enter на клавиатуре открываем выбранную таблицу
  end;
end;


initialization
  {$I ulang_ru.lrs} //добавляем в проект файл с русским текстом, чтобы надписи на кнопках, формах были на русском языке.
end.
