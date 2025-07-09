program GerarScript;
uses
  Vcl.Forms,
  uGerarScript in 'uGerarScript.pas' {frmGerarSqlQuery},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}
begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmGerarSqlQuery, frmGerarSqlQuery);
  Application.Run;
end.
