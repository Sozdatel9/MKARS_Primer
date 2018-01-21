unit backupform;

{$mode objfpc}{$H+}

interface

uses
Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
ComObj, Menus, Windows, ShellAPI, LResources, ExtCtrls, zipper; //подключаем модуль zipper для поддержки ZIP-архивов
type
 { TBackupMainForm }
 TBackupMainForm = class(TForm)
  ListBox1: TListBox;
  RefreshBackups: TButton;
  CreateBackupBtn: TButton;
  Label2: TLabel;
  StatusGroupBox: TGroupBox;
  Label1: TLabel;
  RefreshBtn: TButton;
  RestoreBackupBtn: TButton;
  DeleteBtn: TButton;
  NazadBtn: TButton;
  ListBox2: TComboBox;
  procedure CreateBackupBtnClick(Sender: TObject);
  procedure DeleteBtnClick(Sender: TObject);
  procedure FormCreate(Sender: TObject);
  procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  procedure ListBox2Change(Sender: TObject);
  procedure NazadBtnClick(Sender: TObject);
  procedure RefreshBackupsClick(Sender: TObject);
  procedure RefreshBtnClick(Sender: TObject);
  procedure RestoreBackupBtnClick(Sender: TObject);
private
{ private declarations }
public
{ public declarations }
end;

var
 BackupMainForm: TBackupMainForm;
 OurZipper: TZipper;
 UnZipper: TUnZipper;
 CurrentDate:string;
 BackupFileName:string;
 SelectedBackup:string;
 NomerFayla:integer;
 KolichestvoFaylov:integer;
 ImyaUdalyaemogoFayla:string;
 szPathEntry:string;
 Min1, Hour1, Day1, Month1, Year1: string;
 TempString:string;
implementation
{ TBackupMainForm }
{$R *.lfm}

uses welcome;

{ TBackupMainForm }

procedure TBackupMainForm.CreateBackupBtnClick(Sender: TObject);
begin
 //проверяем, существует ли папка Backups
 if DirectoryExists('Backups') then begin end //если существует, то ничего не делаем
 else begin                                   //если нет, то создаём папку и в случае ошибки выводим соответствующее сообщение
 if CreateDir('Backups') = False then
   ShowMessage('При создании папки для хранения резервных копий произошла ошибка.');
 end;
 szPathEntry:=''; //обнуляем переменную с относительным путем. Это нужно для того, чтобы в резервную копию были добавлены только файлы, без учета папки Data\
 //создаём имя файла на основе текущей даты и времени
 KolichestvoFaylov:=ListBox1.Items.Count;;
 DateTimeToString(Currentdate, 'dd_mm_yyyyyyyyyy_hh_nn_ss', Now);
 BackupFileName:='backup_'+CurrentDate+'.zip';
 OurZipper := TZipper.Create;
 try
  //задаём имя архива
  OurZipper.FileName := 'Backups\'+BackupFileName;
  //добавляем файлы из выпадающего списка в архив
  for NomerFayla:=0 to KolichestvoFaylov-1 do
  begin
   OurZipper.Entries.AddFileEntry('Data\'+ListBox1.Items[NomerFayla], CreateRelativePath(ListBox1.Items[NomerFayla],szPathEntry));
  end;
  OurZipper.ZipAllFiles;
 finally
   OurZipper.Free;
 end;
 //обновляем список резервных копий
 RefreshBackups.Click();
 Label2.Caption:='Резервная копия ' + BackupFileName + ' была успешно создана';
end;

//Выводим месяц прописью. Например, вместо 2 - февраль, вместо 12 - декабрь и т.д
function Monthstr(S: Integer): string;
const
 Mes: array[1..12] of string = ('января', 'февраля', 'марта', 'апреля',
   'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября',
     'декабря');
begin
 Result := Mes[S];
end;

//функция удаления файлов (выводится стандартный диалог удаления файла)
function Recycle(const FileName: string; Wnd: HWND = 0): Boolean;
var
 FileOp: TSHFileOpStruct;
begin
 FillChar(FileOp, SizeOf(FileOp), 0);
 if Wnd = 0 then
   Wnd := Application.MainFormHandle;
 FileOp.Wnd := Wnd;
 FileOp.wFunc := FO_DELETE;
 FileOp.pFrom := PChar(FileName + #0#0);
 FileOp.fFlags := FOF_ALLOWUNDO or FOF_NOERRORUI or FOF_SILENT;
 Result := (SHFileOperation(FileOp) = 0) and (not
   FileOp.fAnyOperationsAborted);
end;

procedure TBackupMainForm.DeleteBtnClick(Sender: TObject);
var tempfilename, Path1:string;
begin
 ImyaUdalyaemogoFayla:=ListBox2.Items[Listbox2.ItemIndex]; //считываем имя резервной копии, которую хотим удалить
 //проверяем имя резервной копии, которую хотим удалить. Если это backup_default.zip, то выводим сообщение об ошибке.
 //Иначе: выводим окно "Подтверждение удаления файла"
 If ImyaUdalyaemogoFayla='backup_default.zip' then
 begin
 Label2.Caption:='ОШИБКА: данная резервная копия защищена от удаления !';
 MessageDlg('Ошибка при удалении резервной копии', 'ОШИБКА !'+#13+#13+'Данная резервная копия защищена от удаления !',mtError, [mbOk], 0);
 end
 else
 begin
 tempfilename:='\Backups\'+ImyaUdalyaemogoFayla; // формируем путь к удаляемому файлу
 Path1:=Application.ExeName; //полный путь и название запущенной программы
 Path1:=ExtractFileDir(Path1); //отбрасываем название программы. Остается путь.
 Path1:=Path1+tempfilename; //добавляем к пути имя удаляемого файла
 //Recycle(tempfilename, Handle);
 Recycle(Path1, Handle); //удаляем файл в корзину. При удалении появится окно "Подтверждение удаления файла"
 Label2.Caption:='Резервная копия ' + ImyaUdalyaemogoFayla + ' была успешно удалена';
 //обновляем список резервных копий
 RefreshBackups.Click();
 end;
end;

procedure ListFileDir(Path: string; FileList: TStrings);
var
 SR: TSearchRec;
 begin
  if SysUtils.FindFirst(Path + 'Data\*.xml', faAnyFile, SR) = 0 then
  begin
    repeat
      if (SR.Attr <> faDirectory) then
      begin
        FileList.Add(SR.Name);
      end;
    until SysUtils.FindNext(SR) <> 0;
    SysUtils.FindClose(SR);
  end;
  if SysUtils.FindFirst(Path + 'Data\Blok2\*.xml', faAnyFile, SR) = 0 then
  begin
    repeat
      if (SR.Attr <> faDirectory) then
      begin
        FileList.Add('Blok2\'+SR.Name);
      end;
    until SysUtils.FindNext(SR) <> 0;
    SysUtils.FindClose(SR);
  end;
  if SysUtils.FindFirst(Path + 'Data\Blok3\*.xml', faAnyFile, SR) = 0 then
  begin
    repeat
      if (SR.Attr <> faDirectory) then
      begin
        FileList.Add('Blok3\'+SR.Name);
      end;
    until SysUtils.FindNext(SR) <> 0;
    SysUtils.FindClose(SR);
  end;
 end;

procedure ListBackupsDir(Path: string; FileList: TStrings);
var
 SR: TSearchRec;
begin
 if SysUtils.FindFirst(Path + 'backup*.zip', faAnyFile, SR) = 0 then
 begin
   repeat
     if (SR.Attr <> faDirectory) then
     begin
       FileList.Add(SR.Name);
     end;
   until SysUtils.FindNext(SR) <> 0;
   SysUtils.FindClose(SR);
 end;
end;

procedure TBackupMainForm.FormCreate(Sender: TObject);
begin
 //проверяем, существует ли папка Backups
 if DirectoryExists('Backups')  then
 begin
 ListBackupsDir('Backups\', ListBox2.Items); //получаем список резервных копий
 ListBox2.ItemIndex:=0;
 ListBox1.Clear(); //очищаем список
 //формируем список файлов, которые нужно добавить в архив
 ListFileDir('', ListBox1.Items);
 //
 If ListBox2.Items[Listbox2.ItemIndex]='backup_default.zip' then
 begin
 ListBox2.Hint:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана 1 июля 2015 года в 12:45';
 Label2.Caption:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана 1 июля 2015 года в 12:45';
 end
 Else Begin
 //Минута
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-7);
 Delete(TempString, 1, Length(TempString)-2);
 Min1:=TempString;
 //Час
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-10);
 Delete(TempString, 1, Length(TempString)-2);
 Hour1:=TempString;
 if (StrToInt(Hour1)<10) then
 begin
 Delete(Hour1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //День
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, 9);
 Delete(TempString, 1, 7);
 Day1:=TempString;
 if (StrToInt(Day1)<10) then
 begin
 Delete(Day1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //Месяц
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, 12);
 Delete(TempString, 1, 10);
 Month1:=TempString;
 if (StrToInt(Month1)<10) then
 begin
 Delete(Month1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //Год
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-13);
 Delete(TempString, 1, 13);
 Year1:=TempString;
 //Выводим информацию о выбранной резервной копии
 ListBox2.Hint:='Резервная копия № ' + IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана ' + Day1+' '+Monthstr(StrToInt(Month1)) +' '+ Year1+' года в ' + Hour1+':'+Min1;
 Label2.Caption:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана ' + Day1+' '+Monthstr(StrToInt(Month1)) +' '+Year1+' года в ' +Hour1+':'+Min1;
 End;
 end //если существует, то ничего не делаем
 else begin                     //если нет, то создаём папку и в случае ошибки выводим соответствующее сообщение
 MessageDlg('МкАРС 1.5 :: Ошибка', 'ОШИБКА :: Папка для хранения резервных копий удалена или отсутствует !'+#13+#13+'Для продолжения работы нажмите ОК',mtError, [mbOk], 0);
 exit;
 end;
 //
end;

procedure TBackupMainForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 //при нажатии клавиши F1 вызываем справку
 if (Key = VK_F1) then BEGIN
   WelcomeForm.MenuItem5.Click;
 end;
 //обрабатываем нажатие клавиши ESC
 if (Key = VK_ESCAPE) then
  begin
    NazadBtn.Click();
  end;
end;

procedure TBackupMainForm.ListBox2Change(Sender: TObject);
begin
//Выводим в нижней части экрана информацию о выбранной резервной копии - имя и время создания.
//Время и дата при этом формируются из имени резервной копии.
 If ListBox2.Items[Listbox2.ItemIndex]='backup_default.zip' then
 begin
 ListBox2.Hint:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана 1 июля 2015 года в 12:45';
 Label2.Caption:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана 1 июля 2015 года в 12:45';
 end
 Else Begin
 //Минута
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-7);
 Delete(TempString, 1, Length(TempString)-2);
 Min1:=TempString;
 //Час
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-10);
 Delete(TempString, 1, Length(TempString)-2);
 Hour1:=TempString;
 if (StrToInt(Hour1)<10) then
 begin
 Delete(Hour1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //День
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, 9);
 Delete(TempString, 1, 7);
 Day1:=TempString;
 if (StrToInt(Day1)<10) then
 begin
 Delete(Day1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //Месяц
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, 12);
 Delete(TempString, 1, 10);
 Month1:=TempString;
 if (StrToInt(Month1)<10) then
 begin
 Delete(Month1, 1, 1);  //удаляем лишний разряд, если число меньше 10
 end;
 //Год
 TempString:='';
 TempString:= ListBox2.Items[Listbox2.ItemIndex];
 SetLength(TempString, Length(TempString)-13);
 Delete(TempString, 1, 13);
 Year1:=TempString;
 //Выводим информацию о выбранной резервной копии
 ListBox2.Hint:='Резервная копия № ' + IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана ' + Day1+' '+Monthstr(StrToInt(Month1)) +' '+ Year1+' года в ' + Hour1+':'+Min1;
 Label2.Caption:='Резервная копия № ' +IntToStr(ListBox2.ItemIndex+1) + #13 + #13 + 'Создана ' + Day1+' '+Monthstr(StrToInt(Month1)) +' '+Year1+' года в ' +Hour1+':'+Min1;
 end;
end;

procedure TBackupMainForm.NazadBtnClick(Sender: TObject);
begin
 WelcomeForm.Show();
 BackupMainForm.Hide();
end;

procedure TBackupMainForm.RefreshBackupsClick(Sender: TObject);
begin
 ListBox2.Clear(); //очищаем список
 ListBackupsDir('Backups\', ListBox2.Items); //получаем список резервных копий
 ListBox2.ItemIndex:=0;
 ListBox1.Clear(); //очищаем список
 //формируем список файлов, которые нужно добавить в резервную копию
 ListFileDir('', ListBox1.Items);
end;
procedure TBackupMainForm.RefreshBtnClick(Sender: TObject);
begin
 //обновляем список резервных копий
 RefreshBackups.Click();
end;

procedure TBackupMainForm.RestoreBackupBtnClick(Sender: TObject);
begin
 SelectedBackup:=ListBox2.Items[ListBox2.ItemIndex]; //считываем имя резервной копии, выбранной из списка
 UnZipper := TUnZipper.Create;
 try
  //задаём имя резервной копии
  UnZipper.FileName := 'Backups\'+SelectedBackup;
  //указываем, куда распаковывать
  UnZipper.OutputPath := 'Data\';
  UnZipper.Examine;
  //распаковываем резервную копию
  UnZipper.UnZipAllFiles;
 finally
  UnZipper.Free;
 end;
 Label2.Caption:='Резервная копия ' + SelectedBackup + ' была успешно восстановлена';
end;
end.
