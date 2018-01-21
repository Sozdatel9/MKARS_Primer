program MKARS;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, //this includes the LCL widgetset
  Forms, welcome, Windows, backupform, aboutform, unique_utils, blok1;

{$R *.res}
var
 MKARSApp: TUniqueInstance;
{После запуска приложения проверяем количество запущенных экземпляров приложения.
Если пользователь пытается повторно запустить приложение, то запуск блокируется.}

begin
 // Создаём объект с уникальным идентификатором
 MKARSApp:=TUniqueInstance.Create('MKARS150');
 // Проверяем, нет ли в системе объектов с таким же идентификатором
 if MKARSApp.IsRunInstance then
  begin
   MKARSApp.Free;
   Halt(1);
  end
 else
  // Если нет - регистрируем в системе наш идентификатор
  MKARSApp.RunListen;
 begin
 Application.Title:='МкАРС 1.5';
 RequireDerivedFormResource := True;
 Application.Initialize;
 Application.CreateForm(TWelcomeForm, WelcomeForm);
 Application.CreateForm(TAboutMainForm, AboutMainForm);
 Application.CreateForm(TBackupMainForm, BackupMainForm);
 Application.CreateForm(TBlok1_Form, Blok1_Form);
 Application.Run;
 end;
end.
