program Importa_ICMS;

uses
  Forms,
  uImporta_ICMS in 'uImporta_ICMS.pas' {frmImportaICMS};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmImportaICMS, frmImportaICMS);
  Application.Run;
end.
