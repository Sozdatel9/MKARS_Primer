unit aboutform;

{$mode objfpc}{$H+}

interface

uses
 Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
 StdCtrls, ExtCtrls, Windows;

type

 { TAboutMainForm }

 TAboutMainForm = class(TForm)
  Label2: TLabel;
  Knopka_OK: TButton;
  Image1: TImage;
  Label1: TLabel;
  procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
  procedure Knopka_OKClick(Sender: TObject);
 private
  { private declarations }
 public
  { public declarations }
 end;

var
 AboutMainForm: TAboutMainForm;

implementation

{$R *.lfm}

{ TAboutMainForm }
uses welcome;

{ TAboutMainForm }

procedure TAboutMainForm.Knopka_OKClick(Sender: TObject);
begin
 WelcomeForm.Show();
 AboutMainForm.Hide();
end;

procedure TAboutMainForm.FormKeyDown(Sender: TObject; var Key: word;
 Shift: TShiftState);
begin
 //обрабатываем нажатие клавиши ESC
 if (Key = VK_ESCAPE) then
 begin
  AboutMainForm.Knopka_OK.Click();
 end;
 //при нажатии клавиши F1 вызываем справку
 if (Key = VK_F1) then
 begin
  WelcomeForm.MenuItem5.Click;
 end;
end;

end.
