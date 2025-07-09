unit uGerarScript;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Menus,
  Vcl.ImgList, System.Win.Registry, Winapi.ShellApi, Vcl.ActnList,
  System.StrUtils, System.ImageList, Vcl.VirtualImageList,
  Vcl.BaseImageCollection, Vcl.ImageCollection;
type
  TfrmGerarSqlQuery = class(TForm)
    pnlBotoes: TPanel;
    GroupBox1: TGroupBox;
    btnGerarSQL: TSpeedButton ;
    btnLimpar: TSpeedButton ;
    btnSair: TSpeedButton ;
    btnAbrirArquivo: TSpeedButton ;
    odSQL: TOpenDialog;
    memoSQL: TMemo;
    memoDelphi: TMemo;
    ppOpcoesMemo: TPopupMenu;
    imgPopUp: TImageList;
    Selecionartudo1: TMenuItem;
    Copiar1: TMenuItem;
    Recortar1: TMenuItem;
    Colar1: TMenuItem;
    split: TSplitter;
    gbOpcoes: TGroupBox;
    ckbTryFinally: TCheckBox;
    ckbInsertUpdate: TCheckBox;
    gbNomeQuery: TGroupBox;
    edtNomeQuery: TEdit;
    ckbNomeQuery: TCheckBox;
    ckbParamByName: TCheckBox;
    pnlPesquisa: TPanel;
    pnlExpansor: TPanel;
    btnExpansor: TSpeedButton ;
    edtPesquisa: TEdit;
    lblPesquisa: TLabel;
    btnPesquisar: TSpeedButton ;
    fdPesquisar: TFindDialog;
    rbDelphi: TRadioButton;
    rbSQL: TRadioButton;
    SpeedButton1: TSpeedButton;
    imgColecao: TImageCollection;
    imgLista: TVirtualImageList;
    edtProgramador: TEdit;
    lblProgramador: TLabel;
    btnSalvar: TSpeedButton;
    sdgSQL: TSaveDialog;
    rdgScriptMCN: TRadioGroup;
    procedure ckbNomeQueryClick(Sender: TObject);
    procedure btnGerarSQLClick(Sender: TObject);
    procedure btnLimparClick(Sender: TObject);
    procedure memoSQLKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnSairClick(Sender: TObject);
    procedure ckbTryFinallyClick(Sender: TObject);
    procedure btnAbrirArquivoClick(Sender: TObject);
    procedure memoSQLChange(Sender: TObject);
    procedure Selecionartudo1Click(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure memoDelphiMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure Recortar1Click(Sender: TObject);
    procedure Colar1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ckbInsertUpdateClick(Sender: TObject);
    procedure memoSQLKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnExpansorClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure fdPesquisarFind(Sender: TObject);
    procedure btnPesquisarClick(Sender: TObject);
    procedure btnSalvarClick(Sender: TObject);
    procedure rdgScriptMCNClick(Sender: TObject);
  private
     { Private declarations }
    vNumLinhaSQL, vNumLinhaDelphi   : Integer;
    FSelPos                         : Integer;
    FMemoPesquisa                   : TMemo;
    ListaSQL: TStringList;
    procedure CriarRegistroWindows();
    procedure AbrirArquivoWindows();
    procedure WMDropFiles(var Msg: TMessage); message wm_DropFiles;
    procedure GerarSQL();
    procedure GerarScriptCriarCampo();
    procedure GerarScriptCriarChavePrimaria();
    procedure GerarScriptCriarChaveEstrangeira();
    procedure GerarScriptCriarTabela();
    function RetonarNomeTabela(): String;
    function RetonarNomeTabelaReferencia(): String;
    function RetonarNomeCampo(): String;
    function RetonarNomeCampoChaveEstrangeira(): String;
    function RetonarNomeCampoChavePrimaria(): String;
    function RetonarTipo(): String;
    function RetonarScriptTabela(): TStringList;
    function RetonarNomeTabelaCriar(): String;
    function ConcatenarQuery(ASQL: String): String;
    function NumeroLinhaDelphi(ALinhaSQL: String): Integer;
    function GerarParamByName(): String;
    function ConcatenarParamByName(AParametro: String): String;
    function Localizar(const StrOri, StrLoc: string; const PosInicial: Longint; DifMaieMin: Boolean = False;
                       ParaCima: Boolean = False; CoincidirPalavra: Boolean = False): Longint;
    procedure Pesquisar();
    procedure ListarArquivos(pDiretorio: String);
    function TemAtributo(Attr, Val: Integer): Boolean;
  public
    class function getQueryTratada(): String;
    { Public declarations }
  end;
var
  frmGerarSqlQuery: TfrmGerarSqlQuery;
implementation
uses
  Clipbrd;
{$R *.dfm}
class function TfrmGerarSqlQuery.getQueryTratada(): String;
begin
  frmGerarSqlQuery := TfrmGerarSqlQuery.Create(nil);
  try
    if frmGerarSqlQuery.ShowModal = mrOk then
      Result := frmGerarSqlQuery.memoDelphi.Lines.Text;
  finally
    FreeAndNil(frmGerarSqlQuery);
  end;
end;

procedure TfrmGerarSqlQuery.ListarArquivos(pDiretorio: String);
var
  SR: TSearchRec;
  ListRel: TStringList;
  vPos: Integer;
  vNumero: String;
begin
  ListRel := TStringList.Create;
  try
    if FindFirst(pDiretorio+ '\*.sql*', faAnyFile, SR) = 0 then
    begin
      repeat
        if (SR.Attr <> faDirectory) then
        begin
          vPos := Pos('_', SR.Name)+1;
          vNumero := Copy(SR.Name, vPos, Length(SR.Name));

          vPos := Pos('.', vNumero)-1;
          vNumero := Copy(vNumero, 1, vPos);

          ListRel.Add(Trim(vNumero));
        end;
      until
      FindNext(SR) <> 0;
      FindClose(SR);
    end;
    ListaSQL.Add(vNumero);
  finally
    ListRel.Free;
  end;
end;

function TfrmGerarSqlQuery.Localizar(const StrOri, StrLoc: string;
  const PosInicial: Integer; DifMaieMin, ParaCima,
  CoincidirPalavra: Boolean): Longint;
var
  i: Longint;
  Achou: Boolean;
  procedure ConferePalavraInteira();
  begin
    if Achou and CoincidirPalavra then
    begin
      if ((IfThen(i = 0, '', Copy(StrOri, i - 1             , 1)) <> '') and (Copy(StrOri, i              - 1, 1)[1] in ['0'..'9','A'..'Z','a'..'z'])) or
         ((IfThen(i = 0, '', Copy(StrOri, i + Length(StrLoc), 1)) <> '') and (Copy(StrOri, i + Length(StrLoc), 1)[1] in ['0'..'9','A'..'Z','a'..'z'])) then
        Achou := False;
    end;
  end;
begin
  Result := -1;
  if ParaCima then // se for para cima ele faz o for (loop) diminuindo o valor.
  begin
    for i := PosInicial - Length(StrLoc) downto 0 do
    begin
      if DifMaieMin then // a var achou deve ser TRUE para sair do looping achando a string
        Achou := StrLoc = Copy(StrOri, i, Length(StrLoc))
      else
        Achou := AnsiUpperCase(StrLoc) = AnsiUpperCase(Copy(StrOri, i, Length(StrLoc)));
      ConferePalavraInteira;
      if Achou then
      begin
        Result := i - 1; // contém a POSICAO do bicho.
        if Result < 0 then
          Result := 0;
        Break;
      end;
    end;
  end
  else  // Normal, do cursor para baixo
  for i := PosInicial to (Length(StrOri) - Length(StrLoc) + 1) do
  begin
    if DifMaieMin then
      Achou := StrLoc = Copy(StrOri, i, Length(StrLoc))
    else
      Achou := AnsiUpperCase(StrLoc) = AnsiUpperCase(Copy(StrOri, i, Length(StrLoc)));
    ConferePalavraInteira;
    if Achou then
    begin
      Result := i - 1;
      if Result < 0 then
        Result := 0;
      Break;
    end;
  end;
end;
procedure TfrmGerarSqlQuery.btnLimparClick(Sender: TObject);
begin
  memoSQL.Clear;
  memoDelphi.Clear;
end;
procedure TfrmGerarSqlQuery.btnPesquisarClick(Sender: TObject);
begin
  Pesquisar;
end;
procedure TfrmGerarSqlQuery.btnSairClick(Sender: TObject);
begin
  Close;
end;
procedure TfrmGerarSqlQuery.btnSalvarClick(Sender: TObject);
var
  vDiretorio: String;
  vNumero: Integer;
begin
  if Trim(memoDelphi.Text) = '' then
  begin
    ShowMessage('Não existe script para salvar!');
    Abort;
  end
  else
  begin
    ListaSQL := TStringList.Create;
    try
      vDiretorio := 'C:\MCNProjetos\MCNSoftware\Banco_de_dados\Scripts';
      ListarArquivos(vDiretorio);
      vNumero := StrToIntDef(Trim(ListaSQL.Text),0)+1;
      memoDelphi.Lines.SaveToFile(vDiretorio+'\Script_'+IntToStr(vNumero)+'.sql');
    finally
      ListaSQL.Free;
    end;
  end;
end;

procedure TfrmGerarSqlQuery.ckbInsertUpdateClick(Sender: TObject);
begin
  GerarSQL;
end;
procedure TfrmGerarSqlQuery.ckbNomeQueryClick(Sender: TObject);
begin
  edtNomeQuery.Enabled  := ckbNomeQuery.Checked;
  ckbTryFinally.Enabled := ckbNomeQuery.Checked;
  edtNomeQuery.Clear;
  GerarSQL;
end;
procedure TfrmGerarSqlQuery.ckbTryFinallyClick(Sender: TObject);
begin
  GerarSQL;
end;
procedure TfrmGerarSqlQuery.Colar1Click(Sender: TObject);
begin
  if  ActiveControl is TMemo then
  begin
    TMemo(ActiveControl).PasteFromClipboard;
    GerarSQL;
  end;
end;
function TfrmGerarSqlQuery.ConcatenarParamByName(AParametro: String): String;
begin
  if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) then
    Result := ('    ' + Trim(edtNomeQuery.Text) + '.ParamByName(''' + AParametro + ''').Value := EmptyStr;')
  else
    Result := ('  ParamByName(''' + AParametro + ''').Value := EmptyStr;');
end;
function TfrmGerarSqlQuery.ConcatenarQuery(ASQL: String): String;
begin
  if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) then
    Result := '    ' + Trim(edtNomeQuery.Text) + '.SQL.Add(''' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ''');'
  else
    Result := '  SQL.Add(''' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ''');' ;
end;
procedure TfrmGerarSqlQuery.Copiar1Click(Sender: TObject);
begin
  if ActiveControl is TMemo then
    TMemo(ActiveControl).CopyToClipboard
end;
procedure TfrmGerarSqlQuery.CriarRegistroWindows();
var
  VRegistro: TRegistry;
begin
  try
    VRegistro := TRegistry.Create;
    try
      VRegistro.RootKey := HKEY_CLASSES_ROOT;
      VRegistro.LazyWrite := False;
      {Define o nome interno e uma legenda para aparecer no Windows Explorer}
      VRegistro.OpenKey('\GerarScriptDelphi', True);
      VRegistro.WriteString('', 'Gerar Script - Arquivo SQL e Txt');
      VRegistro.CloseKey;
      VRegistro.OpenKey('GerarScriptDelphi\shell\open\command', True);
      VRegistro.WriteString('',ParamStr(0) + ' %1'); {NomeDoExe %1}
      VRegistro.CloseKey;
      {Define o ícone a ser usado no Windows Explorer}
      VRegistro.OpenKey('GerarScriptDelphi\DefaultIcon', True);
      VRegistro.WriteString('', ParamStr(0) + ',0');
      VRegistro.CloseKey;
      VRegistro.OpenKey('.txt', True);
      VRegistro.WriteString('', 'Arquivo de Texto');
      VRegistro.CloseKey;
      VRegistro.OpenKey('.SQL', True);
      VRegistro.WriteString('', 'Structured Query Language');
      VRegistro.CloseKey;
    finally
      VRegistro.CloseKey;
      FreeAndNil(VRegistro);
    end;
  except
  end;
end;
procedure TfrmGerarSqlQuery.fdPesquisarFind(Sender: TObject);
var
  P: Integer;
begin
  if rbDelphi.Checked then
    FMemoPesquisa := memoDelphi
  else
  if rbSQL.Checked then
    FMemoPesquisa := memoSQL;
  P:= Localizar(FMemoPesquisa.Text, fdPesquisar.FindText, FMemoPesquisa.SelStart + FMemoPesquisa.SelLength,
                frMatchCase in fdPesquisar.Options, not (frDown in fdPesquisar.Options), frWholeWord in fdPesquisar.Options);
  if P > -1 then
  begin
    FMemoPesquisa.SelStart  := P;
    FMemoPesquisa.SelLength := Length(fdPesquisar.FindText);
    FMemoPesquisa.SetFocus;
  end
  else
    FMemoPesquisa.SelStart  := 0;
end;
procedure TfrmGerarSqlQuery.FormCreate(Sender: TObject);
begin
  btnSalvar.Enabled := rdgScriptMCN.ItemIndex >= 0;
  //CriarRegistroWindows;
  DragAcceptFiles(Handle, True);
end;
procedure TfrmGerarSqlQuery.FormDestroy(Sender: TObject);
begin
  DragAcceptFiles(Handle, False);
end;
procedure TfrmGerarSqlQuery.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = Vk_F3) and (pnlPesquisa.Height = pnlPesquisa.Constraints.MaxHeight) then
    if (not (ssCtrl in Shift)) then
     Pesquisar;
  if (Key = Ord('F')) and (ssCtrl in Shift) then
  begin
    btnExpansor.Click;
    Key := 0;
  end;
  if (Key = VK_F9) then
  begin
    GerarSQL;
    Key := 0;
  end;
  if (Key = VK_DELETE) and ((ssCtrl in Shift) and (ssShift in Shift)) then
  begin
    memoSQL.Clear;
    memoDelphi.Clear;
    Key := 0;
  end;
end;
procedure TfrmGerarSqlQuery.FormShow(Sender: TObject);
begin
  AbrirArquivoWindows;
 // pnlPesquisa.Height := pnlPesquisa.Constraints.MinWidth;
end;
function TfrmGerarSqlQuery.GerarParamByName(): String;
var
  Posicao, PosEspaco, PosParenteses, PosVirgula, MenorPosicao : Integer;
  Parametro, Resto: String;
  ParamByName, ParamMemoSQL: TStringList;
begin
  Result := EmptyStr;
  ParamByName  := TStringList.Create;
  ParamMemoSQL := TStringList.Create;
  try
    ParamMemoSQL.Clear;
    ParamMemoSQL.Text := memoSQL.Lines.Text;
    Posicao := Pos(':', ParamMemoSQL.Text);
    Resto   := stringreplace(copy(ParamMemoSQL.Text, Posicao + 1, Length(ParamMemoSQL.Text)), sLineBreak, ' ', [rfReplaceAll]) + ' ';
    if ParamMemoSQL.Text <> EmptyStr then
    begin
      while Posicao > 0 do
      begin
        PosEspaco     := pos(' ', Resto);
        PosParenteses := pos(')', Resto);
        PosVirgula    := pos(',', Resto);
        if (PosEspaco > PosParenteses) and (PosParenteses > 0) then
          MenorPosicao := PosParenteses
        else
          MenorPosicao := PosEspaco;
        if (PosVirgula < MenorPosicao) and (PosVirgula > 0) then
          MenorPosicao := PosVirgula;
        Parametro := UpperCase(copy(Resto, 0, MenorPosicao - 1));
        if ParamByName.IndexOf(ConcatenarParamByName(Parametro)) < 0 then
          ParamByName.Add(ConcatenarParamByName(Parametro));
        Posicao := Pos(':', Resto);
        Resto   := copy(Resto, Posicao + 1, length(resto + ' ' ));
      end;
      Result := ParamByName.Text;
    end;
  finally
    FreeAndNil(ParamByName);
    FreeAndNil(ParamMemoSQL);
  end;
end;
procedure TfrmGerarSqlQuery.GerarSQL();
var
  i : integer;
  vSQL: TStringList;
begin
  vSQL := TStringList.Create;
  try
    MemoDelphi.Clear;
    if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) and (memoSQL.Text <> EmptyStr) then
    begin
      if ckbTryFinally.Checked then
      begin
        vSQL.Add('var');
        vSQL.Add('  '+Trim(edtNomeQuery.Text)+': TFDQuery;');
        vSQL.Add('begin');
        vSQL.Add('  '+Trim(edtNomeQuery.Text)+' := TFDQuery.Create(nil);');
        vSQL.Add('  try ');
        vSQL.Add('    '+Trim(edtNomeQuery.Text)+'.Connection := DM.Cnt;');
      end;
      vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.Close;');
      vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.SQL.Clear;');
    end
    else
    if memoSQL.Text <> EmptyStr then
    begin
      vSQL.Add('  Close;');
      vSQL.Add('  SQL.Clear;');
    end;
    for I := 0 to memoSQL.Lines.Count -1 do
      vSQL.Add(ConcatenarQuery(memoSQL.Lines[i]));
    if ckbParamByName.Checked then
    begin
      if GerarParamByName <> EmptyStr then
        vSQL.Add('    '+Trim(GerarParamByName));
    end;

    if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) and (memoSQL.Text <> EmptyStr) then
    begin
      case AnsiIndexStr(UpperCase(Copy(Trim(memoSQL.Text),1,6)), ['SELECT','INSERT','UPDATE']) of
        0: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.Open;');
        1..2: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.ExecSQL();');
      end;

      if ckbTryFinally.Checked then
      begin
        vSQL.Add('  finally ');
        vSQL.Add('    FreeAndNil('+ Trim(edtNomeQuery.Text) +');');
        vSQL.Add('  end; ');
      end;
    end
    else
    if memoSQL.Text <> EmptyStr then
    begin
      case AnsiIndexStr(UpperCase(Copy(Trim(memoSQL.Text),1,6)), ['SELECT','INSERT','UPDATE']) of
        0: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.Open;');
        1..2: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.ExecSQL();');
      end;
    end;
  finally
    memoDelphi.Text := vSQL.Text;
    FreeAndNil(vSQL);
  end;
end;
procedure TfrmGerarSqlQuery.GerarScriptCriarCampo();
var
  vTabela, vCampo, vTipo: String;
begin
  vTabela := RetonarNomeTabela();
  vCampo := RetonarNomeCampo();
  vTipo := RetonarTipo();

  memoDelphi.Lines.Add('/*');
  memoDelphi.Lines.Add('    @OBJETIVO: Criar campo '+vCampo+' na tabela '+vTabela);
  memoDelphi.Lines.Add('    @AUTOR...: '+Trim(edtProgramador.Text));
  memoDelphi.Lines.Add('    @DATA....: '+DateToStr(Date));
  memoDelphi.Lines.Add('*/');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('EXECUTE BLOCK');
  memoDelphi.Lines.Add('AS');
  memoDelphi.Lines.Add('BEGIN');
  memoDelphi.Lines.Add('  IF (EXISTE_O_CAMPO_NA_TABELA('''+vTabela+''','''+vCampo+''')=''NAO'') THEN');
  memoDelphi.Lines.Add('  BEGIN');
  memoDelphi.Lines.Add('    EXECUTE STATEMENT ''ALTER TABLE '+vTabela+' ADD '+vCampo+' '+vTipo+';'';');
  memoDelphi.Lines.Add('  END');
  memoDelphi.Lines.Add('END;');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('COMMIT WORK;');
end;

procedure TfrmGerarSqlQuery.GerarScriptCriarChaveEstrangeira();
var
  vTabela, vCampo, vTipo, NomeChave, NomeTabelaReferencia: String;
  vPos, vPosChave, Quantidade: Integer;
begin
  vTabela := RetonarNomeTabela();
  NomeTabelaReferencia := RetonarNomeTabelaReferencia();
  vCampo := RetonarNomeCampoChaveEstrangeira();
  vTipo := RetonarTipo();

  vPos := Pos('CONSTRAINT', memoSQL.Text)+10;
  vPosChave := Pos('FOREIGN', memoSQL.Text);
  Quantidade := (vPosChave-vPos);
  NomeChave := Copy(memoSQL.Text, vPos, Quantidade);
  NomeChave := Trim(NomeChave);

  memoDelphi.Lines.Add('/*');
  memoDelphi.Lines.Add('    @OBJETIVO: Criar chave estrangeira '+NomeChave+' na tabela '+vTabela);
  memoDelphi.Lines.Add('    @AUTOR...: '+Trim(edtProgramador.Text));
  memoDelphi.Lines.Add('    @DATA....: '+DateToStr(Date));
  memoDelphi.Lines.Add('*/');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('EXECUTE BLOCK');
  memoDelphi.Lines.Add('AS');
  memoDelphi.Lines.Add('BEGIN');
  memoDelphi.Lines.Add('  IF (EXISTE_A_CONSTRAINT('''+NomeChave+''')=''NAO'') THEN');
  memoDelphi.Lines.Add('  BEGIN');
  memoDelphi.Lines.Add('    EXECUTE STATEMENT ''ALTER TABLE '+vTabela+' ADD CONSTRAINT '+NomeChave+' FOREIGN KEY ('+vCampo+') REFERENCES '+NomeTabelaReferencia+';'';');
  memoDelphi.Lines.Add('  END');
  memoDelphi.Lines.Add('END;');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('COMMIT WORK;');
end;

procedure TfrmGerarSqlQuery.GerarScriptCriarChavePrimaria();
var
  vTabela, vCampo, vTipo, NomeChave: String;
  vPos, vPosChave, Quantidade: Integer;
begin
  vTabela := RetonarNomeTabela();
  vCampo := RetonarNomeCampoChavePrimaria();
  vTipo := RetonarTipo();

  vPos := Pos('CONSTRAINT', memoSQL.Text)+10;
  vPosChave := Pos('PRIMARY', memoSQL.Text);
  Quantidade := (vPosChave-vPos);
  NomeChave := Copy(memoSQL.Text, vPos, Quantidade);
  NomeChave := Trim(NomeChave);

  memoDelphi.Lines.Add('/*');
  memoDelphi.Lines.Add('    @OBJETIVO: Criar chave primaria '+NomeChave+' na tabela '+vTabela);
  memoDelphi.Lines.Add('    @AUTOR...: '+Trim(edtProgramador.Text));
  memoDelphi.Lines.Add('    @DATA....: '+DateToStr(Date));
  memoDelphi.Lines.Add('*/');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('EXECUTE BLOCK');
  memoDelphi.Lines.Add('AS');
  memoDelphi.Lines.Add('BEGIN');
  memoDelphi.Lines.Add('  IF (EXISTE_A_CONSTRAINT('''+NomeChave+''')=''NAO'') THEN');
  memoDelphi.Lines.Add('  BEGIN');
  memoDelphi.Lines.Add('    EXECUTE STATEMENT ''ALTER TABLE '+vTabela+' ADD CONSTRAINT '+NomeChave+' PRIMARY KEY ('+vCampo+')'+';'';');
  memoDelphi.Lines.Add('  END');
  memoDelphi.Lines.Add('END;');
  memoDelphi.Lines.Add('');
  memoDelphi.Lines.Add('COMMIT WORK;');
end;

procedure TfrmGerarSqlQuery.GerarScriptCriarTabela();
var
  vTabela: String;
  ScriptTabela: TStringList;
begin
  ScriptTabela := TStringList.Create;
  try
    vTabela := RetonarNomeTabelaCriar();
    ScriptTabela := RetonarScriptTabela();

    memoDelphi.Lines.Add('/*');
    memoDelphi.Lines.Add('    @OBJETIVO: Criar tabela '+vTabela);
    memoDelphi.Lines.Add('    @AUTOR...: '+Trim(edtProgramador.Text));
    memoDelphi.Lines.Add('    @DATA....: '+DateToStr(Date));
    memoDelphi.Lines.Add('*/');
    memoDelphi.Lines.Add('');
    memoDelphi.Lines.Add('EXECUTE BLOCK');
    memoDelphi.Lines.Add('AS');
    memoDelphi.Lines.Add('BEGIN');
    memoDelphi.Lines.Add('  IF (EXISTE_TABELA('''+vTabela+''') = ''N'') THEN');
    memoDelphi.Lines.Add('  BEGIN');
    memoDelphi.Lines.Add('    EXECUTE STATEMENT '''+ScriptTabela.Text+'');
    memoDelphi.Lines.Add('  END');
    memoDelphi.Lines.Add('END;');
    memoDelphi.Lines.Add('');
    memoDelphi.Lines.Add('COMMIT WORK;');
  finally
    ScriptTabela.Free;
  end;
end;

procedure TfrmGerarSqlQuery.memoDelphiMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight then
  begin
    Recortar1.Visible := not (Sender = memoDelphi);
    Colar1.Visible    := not (Sender = memoDelphi);
    TMemo(Sender).SetFocus;
  end;
end;
procedure TfrmGerarSqlQuery.memoSQLChange(Sender: TObject);
begin
//  memoDelphi.Lines[vNumLinhaDelphi] := ConcatenarQuery(memoSQL.Lines[vNumLinhaSQL]);
end;
procedure TfrmGerarSqlQuery.memoSQLKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = Ord('A')) and (ssCtrl in Shift) then
  begin
    TMemo(Sender).SelectAll;
    Key := 0;
  end;
  vNumLinhaSQL := memoSQL.CaretPos.Y;
  vNumLinhaDelphi := NumeroLinhaDelphi(ConcatenarQuery(memoSQL.Lines[vNumLinhaSQL]));
end;
procedure TfrmGerarSqlQuery.memoSQLKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = Ord('V')) and (ssCtrl in Shift) then
  begin
    GerarSQL;
    Key := 0;
  end;
  if (Key = Ord('X')) and (ssCtrl in Shift) then
  begin
    GerarSQL;
    Key := 0;
  end;
end;
function TfrmGerarSqlQuery.NumeroLinhaDelphi(ALinhaSQL: String): Integer;
var
  I: Integer;
begin
  for I := 0 to memoDelphi.Lines.Count -1 do
  begin
    if memoDelphi.Lines[I] = ALinhaSQL then
    begin
      Result := I;
      Break;
    end;
  end;
end;
procedure TfrmGerarSqlQuery.Pesquisar();
begin
  fdPesquisar.FindText := edtPesquisa.Text;
  fdPesquisar.Options  := [frDown,frShowHelp];
  fdPesquisarFind(Self);
end;
procedure TfrmGerarSqlQuery.rdgScriptMCNClick(Sender: TObject);
begin
  btnSalvar.Enabled := rdgScriptMCN.ItemIndex >= 0;
end;

procedure TfrmGerarSqlQuery.Recortar1Click(Sender: TObject);
begin
  if TMemo(ActiveControl) is TMemo then
  begin
    TMemo(ActiveControl).CutToClipboard;
    GerarSQL;
  end;
end;
function TfrmGerarSqlQuery.RetonarNomeCampo(): String;
var
  Posicao, PosicaoCampo: Integer;
  Script, Campo: String;
begin
  PosicaoCampo := Pos('ADD', memoSQL.Text)+3;

  Script := Copy(memoSQL.Text, PosicaoCampo, Length(memoSQL.Text)-1);
  Script := Trim(Script);

  Posicao := Pos(' ', Script);

  Campo := Copy(Script, 1, Posicao-1);
  Campo := Trim(Campo);

  Result := Campo;
end;

function TfrmGerarSqlQuery.RetonarNomeCampoChaveEstrangeira(): String;
var
  vPos, vPosChave, Quantidade: Integer;
  CampoChave: String;
begin
  vPos := Pos('(', memoSQL.Text)+1;
  vPosChave := Pos(')', memoSQL.Text);
  Quantidade := (vPosChave-vPos);

  CampoChave := Copy(memoSQL.Text, vPos, Quantidade);
  CampoChave := Trim(CampoChave);

  Result := CampoChave;
end;

function TfrmGerarSqlQuery.RetonarNomeCampoChavePrimaria(): String;
var
  vPos, vPosChave, Quantidade: Integer;
  CampoChave: String;
begin
  vPos := Pos('(', memoSQL.Text)+1;
  vPosChave := Pos(')', memoSQL.Text);
  Quantidade := (vPosChave-vPos);

  CampoChave := Copy(memoSQL.Text, vPos, Quantidade);
  CampoChave := Trim(CampoChave);

  Result := CampoChave;
end;

function TfrmGerarSqlQuery.RetonarNomeTabela(): String;
var
  PosicaoTabela, PosicaoCampo: Integer;
  Script, Tabela: String;
begin
  PosicaoTabela := Pos('ALTER TABLE', memoSQL.Text)+11;

  Script := Copy(memoSQL.Text, PosicaoTabela, Length(memoSQL.Text)-1);
  Script := Trim(Script);

  PosicaoCampo := Pos('ADD', Script);

  Tabela := Copy(Script, 1, PosicaoCampo-1);
  Tabela := Trim(Tabela);

  Result := Tabela;
end;

function TfrmGerarSqlQuery.RetonarNomeTabelaCriar(): String;
var
  vPos, vPosChave, Quantidade: Integer;
  Tabela: String;
begin
  vPos := Pos('TABLE', memoSQL.Text)+5;
  vPosChave := Pos('(', memoSQL.Text);
  Quantidade := (vPosChave-vPos);

  Tabela := Copy(memoSQL.Text, vPos, Quantidade);
  Tabela := Trim(Tabela);

  Result := Tabela;
end;

function TfrmGerarSqlQuery.RetonarNomeTabelaReferencia(): String;
var
  vPos: Integer;
  CampoChave: String;
begin
  vPos := Pos('REFERENCES', memoSQL.Text)+10;

  CampoChave := Copy(memoSQL.Text, vPos, Length(memoSQL.Text));
  CampoChave := Trim(CampoChave);

  Result := CampoChave;
end;

function TfrmGerarSqlQuery.RetonarScriptTabela(): TStringList;
var
  i: Integer;
begin
  Result := TStringList.Create;

  memoSQL.Text := Trim(StringReplace(UpperCase(memoSQL.Text), 'NOW', '''NOW''', [rfReplaceAll, rfIgnoreCase]));

  try
    for i := 0 to Pred(memoSQL.Lines.Count) do
      if i = 0 then
        Result.Add(memoSQL.Lines[i])
      else
      if i = Pred(memoSQL.Lines.Count) then
        Result.Add(DupeString(' ', 10)+memoSQL.Lines[i]+''';')
      else
        Result.Add(DupeString(' ', 10)+memoSQL.Lines[i])
  finally

  end;
end;

function TfrmGerarSqlQuery.RetonarTipo(): String;
var
  Posicao, PosicaoCampo: Integer;
  Script, Tipo: String;
begin
  PosicaoCampo := Pos('ADD', memoSQL.Text)+3;

  Script := Copy(memoSQL.Text, PosicaoCampo, Length(memoSQL.Text)-1);
  Script := Trim(Script);

  Posicao := Pos(' ', Script);

  Tipo := Copy(Script, Posicao, Length(memoSQL.Text)-1);
  Tipo := Trim(Tipo);

  Result := Tipo;
end;

procedure TfrmGerarSqlQuery.Selecionartudo1Click(Sender: TObject);
begin
  if ActiveControl is TMemo then
    TMemo(ActiveControl).SelectAll;
end;
function TfrmGerarSqlQuery.TemAtributo(Attr, Val: Integer): Boolean;
begin
  Result := Attr and Val = Val;
end;

procedure TfrmGerarSqlQuery.WMDropFiles(var Msg: TMessage);
var
  BufferSize: word;
  Drop: HDROP;
  FileName: string;
  Pt: TPoint;
  RctMemo: TRect;
begin
  { Pega o manipulador (handle) da operação
    "arrastar e soltar" (drag-and-drop) }
  Drop := Msg.wParam;
  { Pega o retângulo do Memo }
  RctMemo := memoSQL.BoundsRect;
  if PtInRect(RctMemo, Pt) then
  begin
    { Obtém o comprimento necessário para o nome do arquivo,
      sem contar o caractere nulo do fim da string.
      O segundo parâmetro (zero) indica o primeiro arquivo da lista }
    BufferSize := DragQueryFile(Drop, 0, nil, 0);
    SetLength(FileName, BufferSize +1); { O +1 é p/ nulo do fim da string }
    if DragQueryFile(Drop, 0, PChar(FileName), BufferSize+1) = BufferSize then
    begin
      memoSQL.Lines.LoadFromFile(string(PChar(FileName)));
      GerarSQL;
    end;
  end;
  Msg.Result := 0;
end;
procedure TfrmGerarSqlQuery.AbrirArquivoWindows();
begin
  {Se o primeiro parâmetro for um nome de arquivo existente...}
  if FileExists(ParamStr(1)) then
  begin
    memoSQL.Lines.LoadFromFile(ParamStr(1));
    GerarSQL;
  end;
end;
procedure TfrmGerarSqlQuery.btnAbrirArquivoClick(Sender: TObject);
begin
  odSQL.InitialDir := GetCurrentDir;
  if odSQL.Execute(Handle) then
  begin
    memoSQL.Lines.LoadFromFile(odSQL.FileName);
    GerarSQL;
  end;
end;
procedure TfrmGerarSqlQuery.btnExpansorClick(Sender: TObject);
begin
 { if pnlPesquisa.Height = pnlPesquisa.Constraints.MaxHeight then
    pnlPesquisa.Height := pnlPesquisa.Constraints.MinWidth
  else
    pnlPesquisa.Height := pnlPesquisa.Constraints.MaxHeight;}
  edtPesquisa.Clear;
  if edtPesquisa.CanFocus then
    edtPesquisa.SetFocus;
end;
procedure TfrmGerarSqlQuery.btnGerarSQLClick(Sender: TObject);
begin
  case rdgScriptMCN.ItemIndex of
    0: GerarScriptCriarTabela();
    1: GerarScriptCriarCampo();
    2: GerarScriptCriarChavePrimaria();
    3: GerarScriptCriarChaveEstrangeira();
  else
    GerarSQL();
  end
end;
end.
