unit uImporta_ICMS;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZConnection,
  StdCtrls, Mask, ToolEdit, Buttons, ComCtrls, DBXpress, FMTBcd, SqlExpr,
  ADODB, Grids, DBGrids, ExtCtrls, ppDB, ppCtrls, ppStrtch, ppMemo,
  ppPrnabl, ppClass, ppBands, ppCache, ppProd, ppReport, ppComm, ppRelatv,
  ppDBPipe, StrUtils, ComObj;

type
  TfrmImportaICMS = class(TForm)
    Status: TStatusBar;
    btnConfirma: TBitBtn;
    BitBtn1: TBitBtn;
    lblInforme: TLabel;
    edtCaminho: TFilenameEdit;
    conecta_siscon_aux: TZConnection;
    qryDestino: TZQuery;
    qryImportacao: TZQuery;
    ListBox1: TListBox;
    Conecta_SIAT: TADOConnection;
    qryBuscaSIAT: TADOQuery;
    BitBtn2: TBitBtn;
    qryNome: TZQuery;
    lblInicio: TLabel;
    lblFim: TLabel;
    lblDestino: TLabel;
    BitBtn4: TBitBtn;
    rgOpcoes: TRadioGroup;
    qryVerifica: TZQuery;
    con_siscon: TZConnection;
    qryTributo: TZQuery;
    qryRegistro: TZQuery;
    ppDBPipeline1: TppDBPipeline;
    ppReport1: TppReport;
    DataSource1: TDataSource;
    ZQuery1: TZQuery;
    ppHeaderBand1: TppHeaderBand;
    ppDetailBand1: TppDetailBand;
    ppFooterBand1: TppFooterBand;
    ppGroup1: TppGroup;
    ppGroupHeaderBand1: TppGroupHeaderBand;
    ppGroupFooterBand1: TppGroupFooterBand;
    ppLabel1: TppLabel;
    ppDBMemo1: TppDBMemo;
    ppLabel2: TppLabel;
    ppDBText1: TppDBText;
    ppDBText2: TppDBText;
    ppDBText3: TppDBText;
    ppDBText4: TppDBText;
    ppDBText5: TppDBText;
    ppDBText6: TppDBText;
    ppDBText7: TppDBText;
    ppDBText8: TppDBText;
    ppDBText9: TppDBText;
    ppDBText10: TppDBText;
    ppDBText11: TppDBText;
    ppDBText12: TppDBText;
    ppDBText13: TppDBText;
    ppDBText14: TppDBText;
    ppDBCalc1: TppDBCalc;
    ppDBCalc2: TppDBCalc;
    qryDados: TADOQuery;
    BitBtn3: TBitBtn;
    BitBtn5: TBitBtn;
    con_gsrf: TZConnection;
    qryGSRF: TZQuery;
    qryGSRF2: TZQuery;
    qryGSRF3: TZQuery;
    Edit1: TEdit;
    qryTomador: TZQuery;
    strgDados: TStringGrid;
    qryAutonomo: TZQuery;
    Conecta_NFSE: TADOConnection;
    qryNFSE: TADOQuery;
    conecta_150: TZConnection;
    qryTeste: TZQuery;
    conecta_simplesnacional: TZConnection;
    qryApuracao: TZQuery;
    qryPessoaSN: TZQuery;
    BitBtn6: TBitBtn;
    conecta_brasil: TZConnection;
    qryBrasil: TZQuery;
    BitBtn7: TBitBtn;
    conecta_local_siscon: TZConnection;
    qryRegistro65_local: TZQuery;
    Label1: TLabel;
    procedure BitBtn1Click(Sender: TObject);
    procedure btnConfirmaClick(Sender: TObject);

    procedure FileSearch(const PathName, FileName : string; const InDir :  boolean);
    function TrazNomeCredenciado(cnpj:string):string;
    function TrazNomeCredenciado_Postgres(cnpj:string):string;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);

    procedure ProcessaCartaoCredito;
    procedure ProcessaNotaFiscalEletronica;
    procedure rgOpcoesClick(Sender: TObject);
    procedure GravaArquivoTexto(Reg:string);

    procedure Agrupa_CartaoCredito;
    procedure Agrupa_NotaFiscal;
    procedure Agrupa_Administradora;
    procedure Atualiza_SimplesNacional;
    procedure Atualiza_RecolherRetido;
    procedure Processa_RegistroPagamento;

    procedure Atualiza_Totalizados;//Marcar registros que estão com o valor total do mês. (17/07/2013 - Carlito)
    function TrazID_Tributo(codtrb:string):integer;
    function TrazID_GrupoTributo(codtrb:string):integer;

    procedure Atualiza_Pessoa_Siat;
    function TrazDadosPessoaSiat(tipo:string):string;
    procedure BitBtn3Click(Sender: TObject);

    procedure ProcessaPagamentos;//Alimenta a tabela "pagamentos" (postgre) no banco "siscon".
    procedure TrazDatas;//Traz data lançamento pelo codlnc na tabela siatthe.tbllnc

    procedure ProcessaRelatorioArrecadacao;//Alimenta a tabela "relatorio_arrecadacao" (postgre) no banco "siscon".

    procedure Migra_TBLITR;
    procedure Migra_TBLUOR;

    procedure Processa_RelatorioCreditoGeral; //18/11/2013 Carlito
    procedure Endereco_Cepisa; // Ler o arquivo enviado da cepisa e grava em uma tabela do postgres 25/11/2013 Carlito
    procedure Endereco_Receita;
    procedure con_sisconBeforeConnect(Sender: TObject); // Ler o arquivo enviado da receita e grava em uma tabela do postgres 25/11/2013 Carlito
    function TrazIDPessoaExterna(cpfcnpj:string):integer;
    function VerificaUC_Cepisa(uc:string):boolean;

    procedure Insere_PessoaSistemaSiat;// 29/11/2013 Carlito
    procedure Insere_TabelaEcoAtv;
    procedure Insere_TabelaAtv;// 29/11/2013 Carlito

    procedure Insere_ISSPago; // 02/12/2013
    function TrazIDPessoa_Sistema_Siat(codpes:integer):integer;

    procedure Atualiza_IDPessoaNF;//atualizar o campo pessoa_sistema_siat_id nas tabelas de nota_fiscal e registro65
    function TrazIDPessoa_Sist_Siat(cpfcnpj:string):integer;
    procedure BitBtn5Click(Sender: TObject); //Nesta função eu passo como parâmetro o cpfcnpj da nota ou reg65
    procedure Agrupa_ISS_Pago;
    procedure con_gsrfBeforeConnect(Sender: TObject);
    procedure Insere_Pessoa_GSRF;
    function TrazID_Municipio(cidade:string):integer;
    procedure RetiraNotaFiscalDuplicada;

    procedure Processa_ArrecadacaoGrupoLocal; //Carlito 02/04/2014
    procedure Endereco_Tomadores;// Carlito --> 16/04/2014

    procedure Processa_Rendimentos_Autonomos; // 22/08/2014
    function PesquisaAutonomoCadastrado(cpf:string):boolean;
    procedure Atualiza_NFE_Autonomos; //01/09/2014
    procedure Atualiza_CMC_Autonomos; // verifico se existe no SIAT e dou update no campo "insmun" da tabela "autonomo_receita"05/12/2014
    procedure Importa_Planilha_NotasFiscaisSAT;
    procedure conecta_150BeforeConnect(Sender: TObject);//Importação da planilha para poder fazer o agrupamento por ano e cpf.

    procedure ProcessaMalha_SimplesNacional;//18/03/2015 Carlito
    procedure Separador_RegistrosSN; // Pega cada campo so simples nacional nos arquivos txt para alimentar as tabelas;

    procedure AlimentaNotaFiscalMensal;
    procedure conecta_siscon_auxBeforeConnect(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject); // Aqui vou na base de nota fiscal e alimento o siscon-15/01/2016

    procedure Cruzamento_NFE_Cartao;
    function Verifica_CartaoMensal(cnpj:string;ano,mes:integer):boolean;

    procedure AtualizaEnderecoReceita;
    procedure Brasil_GrauRisco;
    procedure BitBtn7Click(Sender: TObject);
    procedure conecta_brasilBeforeConnect(Sender: TObject);
    procedure conecta_local_sisconBeforeConnect(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;
                                          
var
  frmImportaICMS: TfrmImportaICMS;
  vCont, vID, i, vIDPessoa, x, y, k, r : integer;
  ArqTexto : TextFile;
  Entrada, vTipo, vNomeArquivo, vCaminhoDestino, vNomeMes, vArquivoTexto, vCpfCnpj, vInscMun,
  vNomePessoa, vCodGtr, Item, vGen1, vGen2, vCPF : String;
  Linha  : Longint;
  vDataHora : TDateTime;

  vNome, vTipoLogradouro, vLogradouro, vNumero,
  vComplemento, vBairro, vCEP, vMunicipio, vSituacao, vUF : String;

  vCod_local, vCod_setor, vCod_rota, vCod_sequencia, vUC,
  vReferencia, vFD, vTp_motivo, vClasse, vSit_fatura, vConsumo_kwh,
  vValor_importe, vValor_cosip, vHead, vCNPJ  : String;

  vUm, vDois, vTres, vQuatro, vCinco, vSeis, vSete, vOito, vNove, vDez, vOnze, vDoze,
  vTreze, vQuatorze, vQuinze, vDezesseis, vDezessete, vDezoito, vDezenove, vVinte,
  vVinteUm, vVinteDois, vVinteTres, vVinteQuatro, vVinteCinco, vVinteSeis,vVinteSete, vVinteOito,
  vVinteNove, vTrinta, vTrintaUm, vTrintaDois, vCNPJ_Matriz, vCNPJ_Filial, vPA, vTP_Atividade  : String;

implementation

{$R *.dfm}

function StrZeroString(Zeros: string;
  Quant: integer): String;
{Insere Zeros à frente de uma string}
var
I,Tamanho:integer;
aux: string;
begin
  aux := zeros;
  Tamanho := length(ZEROS);
  ZEROS:='';
  for I:=1 to quant-tamanho do
      ZEROS:=ZEROS + '0';
      aux := zeros + aux;
      StrZeroString := aux;

end;


function XlsToStringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
    xlCellTypeLastCell = $0000000B;
var
    XLApp, Sheet: OLEVariant;
    RangeMatrix: Variant;
begin
Result:=False;
//Cria Excel- OLE Object
XLApp:=CreateOleObject('Excel.Application');
try
    //Esconde Excel
    XLApp.Visible:=False;
    //Abre o Workbook
    XLApp.Workbooks.Open(AXLSFile);
    Sheet:=XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    //Pegar o número da última linha
    x:=XLApp.ActiveCell.Row;
    //Pegar o número da última coluna
    y:=XLApp.ActiveCell.Column;
    //Seta Stringgrid linha e coluna
    AGrid.RowCount:=x;
    AGrid.ColCount:=y;
    //Associaca a variant WorkSheet com a variant do Delphi
    RangeMatrix:=XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
    //Cria o loop para listar os registros no TStringGrid
    k:=1;
    repeat
        for r:=1 to y do
            AGrid.Cells[(r - 1),(k - 1)]:=RangeMatrix[K, R];
        Inc(k,1);
    until k > x;
    RangeMatrix:=Unassigned;
finally
    //Fecha o Excel
    if not VarIsEmpty(XLApp) then
        begin
        XLApp.Quit;
        XLAPP:=Unassigned;
        Sheet:=Unassigned;
        Result:=True;
        end;
    end;
end;

//Função para substituir caracteres especiais.

function TrocaCaracterEspecial(aTexto : string; aLimExt : boolean) : string;
const
  //Lista de caracteres especiais
  xCarEsp: array[1..38] of String = ('á', 'à', 'ã', 'â', 'ä','Á', 'À', 'Ã', 'Â', 'Ä',
                                     'é', 'è','É', 'È','í', 'ì','Í', 'Ì',
                                     'ó', 'ò', 'ö','õ', 'ô','Ó', 'Ò', 'Ö', 'Õ', 'Ô',
                                     'ú', 'ù', 'ü','Ú','Ù', 'Ü','ç','Ç','ñ','Ñ');
  //Lista de caracteres para troca
  xCarTro: array[1..38] of String = ('a', 'a', 'a', 'a', 'a','A', 'A', 'A', 'A', 'A',
                                     'e', 'e','E', 'E','i', 'i','I', 'I',
                                     'o', 'o', 'o','o', 'o','O', 'O', 'O', 'O', 'O',
                                     'u', 'u', 'u','u','u', 'u','c','C','n', 'N');
  //Lista de Caracteres Extras
  xCarExt: array[1..48] of string = ('<','>','!','@','#','$','%','¨','&','*',
                                     '(',')','_','+','=','{','}','[',']','?',
                                     ';',':',',','|','*','"','~','^','´','`',
                                     '¨','æ','Æ','ø','£','Ø','ƒ','ª','º','¿',
                                     '®','½','¼','ß','µ','þ','ý','Ý');
var
  xTexto : string;
  i : Integer;
begin
   xTexto := aTexto;
   for i:=1 to 38 do
     xTexto := StringReplace(xTexto, xCarEsp[i], xCarTro[i], [rfreplaceall]);
   //De acordo com o parâmetro aLimExt, elimina caracteres extras.  
   if (aLimExt) then
     for i:=1 to 48 do
       xTexto := StringReplace(xTexto, xCarExt[i], '', [rfreplaceall]);  
   Result := xTexto;
end;
function AnsiToAscii ( str: String ): String;
var
  i : Integer;
begin
    for i := 1 to Length ( str ) do
    case str[i] of
    'á': str[i] := 'a';
    'é': str[i] := 'e';
    'í': str[i] := 'i';
    'ó': str[i] := 'o';
    'ú': str[i] := 'u';
    'à': str[i] := 'a';
    'è': str[i] := 'e';
    'ì': str[i] := 'i';
    'ò': str[i] := 'o';
    'ù': str[i] := 'u';
    'â': str[i] := 'a';
    'ê': str[i] := 'e';
    'î': str[i] := 'i';
    'ô': str[i] := 'o';
    'û': str[i] := 'u';
    'ä': str[i] := 'a';
    'ë': str[i] := 'e';
    'ï': str[i] := 'i';
    'ö': str[i] := 'o';
    'ü': str[i] := 'u';
    'ã': str[i] := 'a';
    'õ': str[i] := 'o';
    'ñ': str[i] := 'n';
    'ç': str[i] := 'c';
    'Á': str[i] := 'A';
    'É': str[i] := 'E';
    'Í': str[i] := 'I';
    'Ó': str[i] := 'O';
    'Ú': str[i] := 'U';
    'À': str[i] := 'A';
    'È': str[i] := 'E';
    'Ì': str[i] := 'I';
    'Ò': str[i] := 'O';
    'Ù': str[i] := 'U';
    'Â': str[i] := 'A';
    'Ê': str[i] := 'E';
    'Î': str[i] := 'I';
    'Ô': str[i] := 'O';
    'Û': str[i] := 'U';
    'Ä': str[i] := 'A';
    'Ë': str[i] := 'E';
    'Ï': str[i] := 'I';
    'Ö': str[i] := 'O';
    'Ü': str[i] := 'U';
    'Ã': str[i] := 'A';
    'Õ': str[i] := 'O';
    'Ñ': str[i] := 'N';
    'Ç': str[i] := 'C';
 end;
  Result := str;
end;

procedure TfrmImportaICMS.BitBtn1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmImportaICMS.btnConfirmaClick(Sender: TObject);
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now)+' -> '+datetostr(date);


  if rgOpcoes.ItemIndex = 0 then
    ProcessaCartaoCredito
  else if rgOpcoes.ItemIndex = 1 then
    //ProcessaNotaFiscalEletronica
    AlimentaNotaFiscalMensal //15/01/2016
  else if rgOpcoes.ItemIndex = 2 then
    begin
    Showmessage('Não faz nada, a opção anterior já faz isso.');
    //Agrupa_NotaFiscal;
    //Atualiza_SimplesNacional;
    //Atualiza_RecolherRetido;
    end
  else if rgOpcoes.ItemIndex = 3 then
    Processa_RegistroPagamento
  else if rgOpcoes.ItemIndex = 4 then
    ProcessaPagamentos
  else if rgOpcoes.ItemIndex = 5 then
    ProcessaRelatorioArrecadacao
  else if rgOpcoes.ItemIndex = 6 then
    Processa_RelatorioCreditoGeral
  else if rgOpcoes.ItemIndex = 7 then
    Endereco_Receita
  else if rgOpcoes.ItemIndex = 8 then
    Agrupa_CartaoCredito
  else if rgOpcoes.ItemIndex = 9 then
    Endereco_Cepisa
  else if rgOpcoes.ItemIndex = 10 then
    Insere_ISSPago
  else if rgOpcoes.ItemIndex = 11 then
    Agrupa_ISS_Pago
  else if rgOpcoes.ItemIndex = 12 then
    Processa_ArrecadacaoGrupoLocal
  else if rgOpcoes.ItemIndex = 13 then
    Endereco_Tomadores
  else if rgOpcoes.ItemIndex = 14 then
    Processa_Rendimentos_Autonomos
  else if rgOpcoes.ItemIndex = 15 then
    Atualiza_NFE_Autonomos
  else if rgOpcoes.ItemIndex = 16 then
    Atualiza_CMC_Autonomos
  else if rgOpcoes.ItemIndex = 17 then
    ProcessaMalha_SimplesNacional;

  lblFim.Caption := 'FINAL: '+timetostr(now)+' -> '+datetostr(date);

  //ShowMessage('Final da Importação!');
  ListBox1.Clear;

end;

procedure TfrmImportaICMS.FileSearch(const PathName, FileName: string;
  const InDir: boolean);
var Rec  : TSearchRec;
    Path : string;
begin
Path := IncludeTrailingBackslash(PathName);
if FindFirst(Path + FileName, faAnyFile - faDirectory, Rec) = 0 then
 try
   repeat
     ListBox1.Items.Add(Path + Rec.Name);
   until FindNext(Rec) <> 0;
 finally
   FindClose(Rec);
 end;

If not InDir then Exit;

if FindFirst(Path + '*.*', faDirectory, Rec) = 0 then
 try
   repeat
    if ((Rec.Attr and faDirectory) <> 0)  and (Rec.Name<>'.') and (Rec.Name<>'..') then
     FileSearch(Path + Rec.Name, FileName, True);
   until FindNext(Rec) <> 0;
 finally
   FindClose(Rec);
 end;

end;

function TfrmImportaICMS.TrazNomeCredenciado(cnpj: string): string;
begin
  Result := '';

  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select nomrazres from SIATTHE.tblpes   ');
  qryBuscaSIAT.SQL.Add('where cpfcnpj=:cnpj     ');
  qryBuscaSIAT.Parameters.ParamByName('cnpj').Value := cnpj;
  qryBuscaSIAT.open;

  if qryBuscaSIAT.RecordCount > 0 then
    Result := Trim(qryBuscaSIAT.fieldbyname('nomrazres').AsString);


end;

procedure TfrmImportaICMS.BitBtn2Click(Sender: TObject);
begin
  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select * from registro65 order by id ');
  qryDestino.open;

  vCont := 0;

  while not qryDestino.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := IntToStr(vCont)+' - ID.: '+IntToStr(qryDestino.fieldbyname('id').Value);
    Application.ProcessMessages;


    qryImportacao.close;
    qryImportacao.sql.Clear;
    qryImportacao.sql.add('update registro65 set nome_credenciado =:nome  ');
    qryImportacao.sql.add('where id =:id     ');
    qryImportacao.ParamByName('id').Value   := qryDestino.fieldbyname('id').Asinteger;
    qryImportacao.ParamByName('nome').Value := TrazNomeCredenciado(qryDestino.fieldbyname('cnpj_mf').Value);
    qryImportacao.ExecSQL;


    qryDestino.Next;
    end;

  showmessage('Acabou...');



end;

function TfrmImportaICMS.TrazNomeCredenciado_Postgres(
  cnpj: string): string;
begin

  Result := '';

  qryNome.Close;
  qryNome.SQL.Clear;
  qryNome.SQL.Add('select nome from pessoa   ');
  qryNome.SQL.Add('where cnpj=:cnpj     ');
  qryNome.ParamByName('cnpj').Value := cnpj;
  qryNome.open;

  if qryNome.RecordCount > 0 then
    Result := AnsiToAscii(Trim(qryNome.fieldbyname('nome').AsString));

end;

procedure TfrmImportaICMS.BitBtn4Click(Sender: TObject);
begin

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select mes, ano, id from agrupamento order by id ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryImportacao.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryDestino.close;
    qryDestino.sql.Clear;
    qryDestino.sql.add('update agrupamento set data =:dt  ');
    qryDestino.sql.add('where id =:id     ');
    qryDestino.ParamByName('id').Value   := qryImportacao.fieldbyname('id').Asinteger;
    qryDestino.ParamByName('dt').Value   := qryImportacao.fieldbyname('ano').Asstring+'-'+qryImportacao.fieldbyname('mes').AsString+'-01';

    qryDestino.ExecSQL;

    qryImportacao.next;
    end;


   showmessage('acabou..');
end;

procedure TfrmImportaICMS.ProcessaCartaoCredito;
begin
  Linha   := 0;
  Entrada := '';  vNomeArquivo := '';
  vCaminhoDestino := 'C:\SEMF\Cartao\Processados\';
   // 'D:\SEMF\Cartao\Processados\';
  //  'C:\SEMF\Cartao\Processados\';

  vCont := 0;
  vID   := 0;

  lblDestino.Caption := 'Caminho Destino: '+vCaminhoDestino;

  FileSearch(edtCaminho.Text,'*.txt', false);

  for i := 0 to ListBox1.Items.Count-1 do
    begin

    vNomeArquivo := copy(listbox1.Items[i],25,50);
    Linha := 0;

    If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
      begin
      AssignFile(ArqTexto,edtCaminho.Text + Trim(vNomeArquivo));
      Reset(ArqTexto);

      vDataHora := now;

      while not Eof(ArqTexto) do
        begin
        Linha := Linha + 1;
        Readln(ArqTexto,Entrada);

        vTipo  := Copy(Entrada,1,2);
        vCont := vCont + 1;

        Status.Panels[0].Text := 'Registros: ' + IntToStr(Linha)+' Tipo: '+vTipo;
        Application.ProcessMessages;


        if vTipo = '10' then
          begin
          qryDestino.Close;
          qryDestino.SQL.Clear;
          qryDestino.SQL.Add('INSERT INTO registro10(id, tipo, cnpj_mf, insc_estadual, nome_administradora, municipio, ');
          qryDestino.SQL.Add('  uf, fax, data_inicial, data_final, cod_convenio, cod_natureza_operacao, cod_finalidade_arquivo,nome_arquivo) ');

          qryDestino.SQL.Add('VALUES (nextval(''registro10_id_seq''), :tipo, :cnpj_mf, :insc_estadual, :nome_administradora, :municipio, ');
          qryDestino.SQL.Add('  :uf, :fax, :data_inicial, :data_final, :cod_convenio, :cod_natureza_operacao, :cod_finalidade_arquivo,:nome_arquivo) ');

          qryDestino.ParamByName('tipo').Value                := Copy(Entrada,1,2);
          qryDestino.ParamByName('cnpj_mf').Value             := Copy(Entrada,3,14);
          qryDestino.ParamByName('insc_estadual').Value       := Copy(Entrada,17,14);
          qryDestino.ParamByName('nome_administradora').Value := AnsiToAscii(Copy(Entrada,31,35));
          qryDestino.ParamByName('municipio').Value := Copy(Entrada,66,30);
          qryDestino.ParamByName('uf').Value        := Copy(Entrada,96,2);
          qryDestino.ParamByName('fax').Value       := Copy(Entrada,98,10);
          qryDestino.ParamByName('data_inicial').Value           := Copy(Entrada,108,4)+'-'+Copy(Entrada,112,2)+'-'+Copy(Entrada,114,2);
          qryDestino.ParamByName('data_final').Value             := Copy(Entrada,116,4)+'-'+Copy(Entrada,120,2)+'-'+Copy(Entrada,122,2);
          qryDestino.ParamByName('cod_convenio').Value           := Copy(Entrada,124,1);
          qryDestino.ParamByName('cod_natureza_operacao').Value  := Copy(Entrada,125,1);
          qryDestino.ParamByName('cod_finalidade_arquivo').Value := Copy(Entrada,126,1);
          qryDestino.ParamByName('nome_arquivo').Value := Trim(vNomeArquivo);
          qryDestino.ExecSQL;

          qryImportacao.Close;
          qryImportacao.SQL.Clear;
          qryImportacao.SQL.Add('select Max(id) as Ultimo from registro10');
          qryImportacao.open;

          vID := qryImportacao.fieldbyname('ultimo').AsInteger;


          end
        else if vTipo = '11' then
          begin
          qryDestino.Close;
          qryDestino.SQL.Clear;
          qryDestino.SQL.Add('INSERT INTO registro11(id, tipo, logradouro, numero, complemento, bairro, ');
          qryDestino.SQL.Add(' cep, nome_contato, telefone,registro10_id)                         ');
          qryDestino.SQL.Add('VALUES (nextval(''registro11_id_seq''), :tipo, :logradouro, :numero, :complemento, :bairro, ');
          qryDestino.SQL.Add(' :cep, :nome_contato, :telefone,:registro10_id)                         ');

          qryDestino.ParamByName('tipo').Value         := Copy(Entrada,1,2);
          qryDestino.ParamByName('logradouro').Value   := AnsiToAscii(Copy(Entrada,3,34));
          qryDestino.ParamByName('numero').Value       := Copy(Entrada,37,5);
          qryDestino.ParamByName('complemento').Value  := AnsiToAscii(Copy(Entrada,42,22));
          qryDestino.ParamByName('bairro').Value       := AnsiToAscii(Copy(Entrada,64,15));
          qryDestino.ParamByName('cep').Value          := Copy(Entrada,79,8);
          qryDestino.ParamByName('nome_contato').Value := AnsiToAscii(Copy(Entrada,87,28));
          qryDestino.ParamByName('telefone').Value     := Copy(Entrada,115,12);
          qryDestino.ParamByName('registro10_id').Value := vID;
          qryDestino.ExecSQL;

          end
        else if vTipo = '65' then
          begin
          vCpfCnpj := '';

          qryRegistro65_local.Close;
          qryRegistro65_local.SQL.Clear;
          qryRegistro65_local.SQL.Add('INSERT INTO registro65(id, tipo, cnpj_mf, insc_estadual, data, numero_documento, numero_autorizacao, ');
          qryRegistro65_local.SQL.Add(' natureza_operacao, tipo_operacao, valor_operacao, modelo_documento,          ');
          qryRegistro65_local.SQL.Add(' num_doc_fiscal, num_cadastro_estabelecimento, branco1, branco2, branco3, registro10_id, nome_credenciado,     ');
          qryRegistro65_local.SQL.Add('  pessoa_sistema_siat_id)                  ');

          qryRegistro65_local.SQL.Add('VALUES (nextval(''registro65_id_seq''), :tipo, :cnpj_mf, :insc_estadual, :data, :numero_documento, :numero_autorizacao, ');
          qryRegistro65_local.SQL.Add(' :natureza_operacao, :tipo_operacao, :valor_operacao, :modelo_documento,          ');
          qryRegistro65_local.SQL.Add(' :num_doc_fiscal, :num_cadastro_estabelecimento, :branco1, :branco2, :branco3,:registro10_id, :nome_credenciado,     ');
          qryRegistro65_local.SQL.Add(' :pid)                  ');

          qryRegistro65_local.ParamByName('tipo').Value                := Copy(Entrada,1,2);
          qryRegistro65_local.ParamByName('cnpj_mf').Value       := Copy(Entrada,3,14);

          vCpfCnpj := Copy(Entrada,3,14);

          qryRegistro65_local.ParamByName('insc_estadual').Value := Copy(Entrada,17,14);
          qryRegistro65_local.ParamByName('data').Value          := Copy(Entrada,31,4)+'-'+Copy(Entrada,35,2)+'-'+ Copy(Entrada,37,2);
          qryRegistro65_local.ParamByName('numero_documento').Value := Copy(Entrada,39,18);
          qryRegistro65_local.ParamByName('numero_autorizacao').Value := Copy(Entrada,39,18);
          qryRegistro65_local.ParamByName('natureza_operacao').Value  := Copy(Entrada,57,1);
          qryRegistro65_local.ParamByName('tipo_operacao').Value      := Copy(Entrada,58,1);
          qryRegistro65_local.ParamByName('valor_operacao').Value     := strtofloat(Copy(Entrada,59,13))/100;
          qryRegistro65_local.ParamByName('modelo_documento').Value   := Copy(Entrada,72,2);
          qryRegistro65_local.ParamByName('num_doc_fiscal').Value     := Copy(Entrada,74,10);
          qryRegistro65_local.ParamByName('num_cadastro_estabelecimento').Value := Copy(Entrada,84,20);
          qryRegistro65_local.ParamByName('branco1').Value := Copy(Entrada,84,43);
          qryRegistro65_local.ParamByName('branco2').Value := Copy(Entrada,104,23);
          qryRegistro65_local.ParamByName('branco3').Value := Copy(Entrada,106,21);

          qryRegistro65_local.ParamByName('registro10_id').Value    := vID;
          qryRegistro65_local.ParamByName('nome_credenciado').Value := TrazNomeCredenciado_Postgres(Copy(Entrada,3,14));
          qryRegistro65_local.ParamByName('pid').Value              := TrazIDPessoa_Sist_Siat(vCpfCnpj);

          qryRegistro65_local.ExecSQL;

          end
        else if vTipo = '66' then
          begin
          qryDestino.Close;
          qryDestino.SQL.Clear;
          qryDestino.SQL.Add('INSERT INTO registro66(id, tipo, cnpj_mf, insc_estadual, anomes_referencia, mesano_referencia, ');
          qryDestino.SQL.Add(' montante_cartao_credito, montante_cartao_debito, branco, registro10_id)             ');
          qryDestino.SQL.Add('VALUES (nextval(''registro66_id_seq''), :tipo, :cnpj_mf, :insc_estadual, :anomes_referencia, :mesano_referencia, ');
          qryDestino.SQL.Add(' :montante_cartao_credito, :montante_cartao_debito, :branco, :registro10_id)             ');

          qryDestino.ParamByName('tipo').Value                := Copy(Entrada,1,2);

          qryDestino.ParamByName('cnpj_mf').Value := Copy(Entrada,3,14);
          qryDestino.ParamByName('insc_estadual').Value := Copy(Entrada,17,14);
          qryDestino.ParamByName('anomes_referencia').Value := Copy(Entrada,31,6);
          qryDestino.ParamByName('mesano_referencia').Value := Copy(Entrada,31,6);
          qryDestino.ParamByName('montante_cartao_credito').Value := strtofloat(Copy(Entrada,37,18))/100;
          qryDestino.ParamByName('montante_cartao_debito').Value := strtofloat(Copy(Entrada,55,18))/100;
          qryDestino.ParamByName('branco').Value := Copy(Entrada,73,54);
          qryDestino.ParamByName('registro10_id').Value := vID;

          qryDestino.ExecSQL;

          end
        else if vTipo = '90' then
          begin
          qryDestino.Close;
          qryDestino.SQL.Clear;
          qryDestino.SQL.Add('INSERT INTO registro90(id, tipo, cnpj_mf, insc_estadual, tipo_totalizado65, total_65, ');
          qryDestino.SQL.Add(' tipo_totalizado66, total_66, total_geral, total_registros, brancos,numero_registro_90, registro10_id, nome_arquivo)      ');
          qryDestino.SQL.Add('VALUES (nextval(''registro90_id_seq''), :tipo, :cnpj_mf, :insc_estadual, :tipo_totalizado65, :total_65, ');
          qryDestino.SQL.Add(' :tipo_totalizado66, :total_66, :total_geral, :total_registros, :brancos,:numero_registro_90, :registro10_id, :nome_arquivo)      ');

          qryDestino.ParamByName('tipo').Value                := Copy(Entrada,1,2);

          qryDestino.ParamByName('cnpj_mf').Value := Copy(Entrada,3,14);
          qryDestino.ParamByName('insc_estadual').Value := Copy(Entrada,17,14);
          qryDestino.ParamByName('tipo_totalizado65').Value := Copy(Entrada,31,2);
          qryDestino.ParamByName('total_65').Value := Copy(Entrada,33,8);
          qryDestino.ParamByName('tipo_totalizado66').Value := Copy(Entrada,41,2);
          qryDestino.ParamByName('total_66').Value := Copy(Entrada,43,8);
          qryDestino.ParamByName('total_geral').Value := Copy(Entrada,51,2);
          qryDestino.ParamByName('total_registros').Value := Copy(Entrada,53,8);
          qryDestino.ParamByName('brancos').Value := Copy(Entrada,61,65);
          qryDestino.ParamByName('numero_registro_90').Value := Copy(Entrada,126,1);

          qryDestino.ParamByName('registro10_id').Value := vID;
          qryDestino.ParamByName('nome_arquivo').Value := Trim(vNomeArquivo);

          qryDestino.ExecSQL;


          end;

        Status.Panels[1].Text := IntToStr(vCont)+' Arq.: '+Trim(vNomeArquivo);

        Application.ProcessMessages;

        end;//while


      end; //If FileExists

      CloseFile(ArqTexto);

      //Move arquivo lido
      If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
        MoveFile(Pchar(edtCaminho.Text+Trim(vNomeArquivo)),Pchar(vCaminhoDestino+Trim(vNomeArquivo)));


    end;//for ...

end;

procedure TfrmImportaICMS.ProcessaNotaFiscalEletronica;
var vNome, vSerie : string;
begin
  Linha   := 0;
  Entrada := '';  vNomeArquivo := '';
  vCaminhoDestino := 'C:\SEMF\NFE\Processados\';

  vCont := 0;
  vID   := 0;

  lblDestino.Caption := 'Caminho Destino: '+vCaminhoDestino;

  FileSearch(edtCaminho.Text,'*.spd', false);

  for i := 0 to ListBox1.Items.Count-1 do
    begin

    vNomeArquivo := copy(listbox1.Items[i],22,50);
    Linha := 0;

    If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
      begin
      AssignFile(ArqTexto,edtCaminho.Text + Trim(vNomeArquivo));
      Reset(ArqTexto);

      while not Eof(ArqTexto) do
        begin
        Linha := Linha + 1;
        Readln(ArqTexto,Entrada);

        vTipo  := Copy(Entrada,1,1);
        vCont := vCont + 1;

        vHead := Copy(Entrada,1,1);

       
        if vHead = 'H' then
          vCNPJ := Copy(Entrada,41,14);

        vSerie := Trim(Copy(Entrada,12,2));


        Status.Panels[0].Text := 'Registros: ' + IntToStr(Linha)+' Tipo: '+vTipo;
        Application.ProcessMessages;

        if (vTipo = 'E') and (vSerie = 'NF') then
          begin
          vCpfCnpj := '';

          qryDestino.Close;
          qryDestino.SQL.Clear;
          qryDestino.SQL.Add('INSERT INTO nota_fiscal(id, tipo, data_emissao, serie, modelo, natureza, numero_notafiscal, valor_notafiscal, ');
          qryDestino.SQL.Add('  valor_servico, tipo_recolhimento, aliquota_iss, situacao, tipo_documento,cnpj_cpf, ');
          qryDestino.SQL.Add('  nome_reduzido, nome_cidade,uf_cidade, codigo_pagamento,codigo_atividade,         ');
          qryDestino.SQL.Add('  numero_finalnota, codigo_obra, cidade_prestacao, uf_prestacao, ');
          qryDestino.SQL.Add('  codigo_servico, percentual_deducao, tomador_simples, tomador_orgao, tipo_operacao, nome_arquivo, linha_arquivo, ');
          qryDestino.SQL.Add('  pessoa_sistema_siat_id, cnpj)                  ');
          qryDestino.SQL.Add('VALUES (nextval(''nota_fiscal_id_seq''), :tipo, :data_emissao, :serie, :modelo, :natureza, :numero_notafiscal, :valor_notafiscal, ');
          qryDestino.SQL.Add('  :valor_servico, :tipo_recolhimento, :aliquota_iss, :situacao, :tipo_documento,:cnpj_cpf, ');
          qryDestino.SQL.Add('  :nome_reduzido, :nome_cidade,:uf_cidade, :codigo_pagamento,:codigo_atividade, ');
          qryDestino.SQL.Add('  :numero_finalnota, :codigo_obra, :cidade_prestacao, :uf_prestacao, ');
          qryDestino.SQL.Add('  :codigo_servico, :percentual_deducao, :tomador_simples, :tomador_orgao, :tipo_operacao, :nome_arquivo,:linha_arquivo, ');
          qryDestino.SQL.Add('  :pid, :cnpj)                 ');

          qryDestino.ParamByName('tipo').Value := Trim(vTipo);

          qryDestino.ParamByName('data_emissao').Value := Copy(Entrada,8,4)+'-'+Copy(Entrada,5,2)+'-'+Copy(Entrada,2,2);

          qryDestino.ParamByName('serie').Value             := vSerie;
          qryDestino.ParamByName('modelo').Value            := Trim(Copy(Entrada,14,1));
          qryDestino.ParamByName('natureza').Value          := Trim(Copy(Entrada,15,1));
          qryDestino.ParamByName('numero_notafiscal').Value := Trim(Copy(Entrada,16,9));
          qryDestino.ParamByName('valor_notafiscal').Value  := Copy(Entrada,25,15);
          qryDestino.ParamByName('valor_servico').Value     := Copy(Entrada,40,15);
          qryDestino.ParamByName('tipo_recolhimento').Value := Trim(Copy(Entrada,55,1));
          qryDestino.ParamByName('aliquota_iss').Value      := Copy(Entrada,56,5);
          qryDestino.ParamByName('situacao').Value          := Trim(Copy(Entrada,61,1));
          qryDestino.ParamByName('tipo_documento').Value    := Trim(Copy(Entrada,62,3));
          qryDestino.ParamByName('cnpj_cpf').Value          := vCNPJ;

          vCpfCnpj :=  Trim(Copy(Entrada,65,14));

          vNome := Trim(Copy(Entrada,79,40));
          vNome := StringReplace( Trim(vNome), ''''  , '' , [rfReplaceAll]);
          vNome := AnsiToAscii(vNome);
          qryDestino.ParamByName('nome_reduzido').Value     := UpperCase(vNome);

          qryDestino.ParamByName('nome_cidade').Value       := UpperCase(Trim(Copy(Entrada,119,30)));
          qryDestino.ParamByName('uf_cidade').Value         := Trim(Copy(Entrada,149,2));
          qryDestino.ParamByName('codigo_pagamento').Value  := Trim(Copy(Entrada,151,5));
          qryDestino.ParamByName('codigo_atividade').Value  := Trim(Copy(Entrada,156,10));
          qryDestino.ParamByName('numero_finalnota').Value  := Trim(Copy(Entrada,166,9));
          qryDestino.ParamByName('codigo_obra').Value       := Trim(Copy(Entrada,175,15));
          qryDestino.ParamByName('cidade_prestacao').Value  := UpperCase(Trim(Copy(Entrada,190,30)));
          qryDestino.ParamByName('uf_prestacao').Value      := Trim(Copy(Entrada,220,2));
          qryDestino.ParamByName('codigo_servico').Value    := Trim(Copy(Entrada,222,10));
          qryDestino.ParamByName('percentual_deducao').Value:= Copy(Entrada,232,5);
          qryDestino.ParamByName('tomador_simples').Value   := Trim(Copy(Entrada,237,1));
          qryDestino.ParamByName('tomador_orgao').Value     := Trim(Copy(Entrada,238,1));
          qryDestino.ParamByName('tipo_operacao').Value     := Trim(Copy(Entrada,239,1));

          qryDestino.ParamByName('nome_arquivo').Value  := Trim(vNomeArquivo);
          qryDestino.ParamByName('linha_arquivo').Value := Linha;
          qryDestino.ParamByName('pid').Value           := TrazIDPessoa_Sist_Siat(vCNPJ);
          qryDestino.ParamByName('cnpj').Value          := vCpfCnpj;


//          qryDestino.ExecSQL;

          Try
            Try
            qryDestino.ExecSQL;
            Except
           // vArquivoTexto := 'C:\SEMF\EXCESSOES.TXT';
           // GravaArquivoTexto(Entrada);

            End;
          Finally
          End;



          end;//if vTipo...
        Edit1.Text := Trim(vNomeArquivo);
        lblFim.Caption := 'FINAL: '+timetostr(now)+' -> '+datetostr(date);

        Status.Panels[1].Text := IntToStr(vCont)+' Arq.: '+Trim(vNomeArquivo);
        Application.ProcessMessages;

        end;//while not Eof...


      end; //If FileExists

      CloseFile(ArqTexto);

      //Move arquivo lido
      If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
        MoveFile(Pchar(edtCaminho.Text+Trim(vNomeArquivo)),Pchar(vCaminhoDestino+Trim(vNomeArquivo)));


    end;//for ...

end;

procedure TfrmImportaICMS.rgOpcoesClick(Sender: TObject);
begin
  if rgOpcoes.ItemIndex = 0 then
    edtCaminho.Text := 'C:\SEMF\Cartao\Arquivos\'
  else if rgOpcoes.ItemIndex = 1 then
    edtCaminho.Text := 'C:\SEMF\NFE\Arquivos\'
  else if rgOpcoes.ItemIndex = 17 then
    edtCaminho.Text := 'C:\SEMF\Simples Nacional\Arquivos\';



end;

procedure TfrmImportaICMS.GravaArquivoTexto(Reg: string);
var
   vArq : TextFile;

begin
   AssignFile(vArq, vArquivoTexto);

   if FileExists(vArquivoTexto) Then
     Append(vArq)
   else
     Rewrite(vArq);


  Writeln(vArq, Reg);

  CloseFile(vArq);

end;

procedure TfrmImportaICMS.Agrupa_NotaFiscal;
begin


  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select serie, cnpj_cpf, sum(valor_notafiscal) as valor_notafiscal,              ');
  qryImportacao.SQL.Add('sum(valor_servico) as valor_servico, sum(aliquota_iss) as aliquota_iss,         ');
  qryImportacao.SQL.Add('extract(YEAR from data_emissao) as ano, extract(MONTH from data_emissao) as mes ');
  qryImportacao.SQL.Add('from nota_fiscal                                                  ');
  qryImportacao.SQL.Add('where cnpj_cpf <> '''' and serie = ''NF'' and situacao = ''N''    ');      //and extract(YEAR from data_emissao) = 2012
  qryImportacao.SQL.Add('group by serie, cnpj_cpf, ano, mes              ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  vDataHora := now;

  while not qryImportacao.eof do
    begin

    vNomeMes := '';
    case qryImportacao.fieldbyname('mes').Value of
      1:  vNomeMes := 'JANEIRO';
      2:  vNomeMes := 'FEVEREIRO';
      3:  vNomeMes := 'MARCO';
      4:  vNomeMes := 'ABRIL';
      5:  vNomeMes := 'MAIO';
      6:  vNomeMes := 'JUNHO';
      7:  vNomeMes := 'JULHO';
      8:  vNomeMes := 'AGOSTO';
      9:  vNomeMes := 'SETEMBRO';
      10: vNomeMes := 'OUTUBRO';
      11: vNomeMes := 'NOVEMBRO';
      12: vNomeMes := 'DEZEMBRO';
    end;

    vCont := vCont + 1;

    Status.Panels[1].Text := '1-Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;


    qryDestino.Close;
    qryDestino.SQL.Clear;
    qryDestino.SQL.Add('INSERT INTO notafiscal_agrupada(id, cnpj_cpf, valor_notafiscal, valor_servico, ');
    qryDestino.SQL.Add('aliquota_iss, ano, mes, nome_mes, data, datahora_processamento)                                      ');
    qryDestino.SQL.Add('VALUES (nextval(''notafiscal_agrupada_id_seq''), :cnpj_cpf, :valor_notafiscal, :valor_servico,  ');
    qryDestino.SQL.Add(':aliquota_iss, :ano, :mes, :nome_mes, :data, :datahora_processamento)                  ');

    qryDestino.ParamByName('cnpj_cpf').Value         := qryImportacao.fieldbyname('cnpj_cpf').Value;
    //qryDestino.ParamByName('nome_reduzido').Value    := qryImportacao.fieldbyname('nome_reduzido').Value;
    qryDestino.ParamByName('valor_notafiscal').Value := qryImportacao.fieldbyname('valor_notafiscal').Value;
    qryDestino.ParamByName('valor_servico').Value    := qryImportacao.fieldbyname('valor_servico').Value;
    qryDestino.ParamByName('aliquota_iss').Value     := qryImportacao.fieldbyname('aliquota_iss').Value;

    qryDestino.ParamByName('ano').Value              := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value              := qryImportacao.fieldbyname('mes').Value;
    qryDestino.ParamByName('nome_mes').Value         := vNomeMes;
    qryDestino.ParamByName('data').Value             := qryImportacao.fieldbyname('ano').Asstring+'-'+
                                                        qryImportacao.fieldbyname('mes').AsString+'-01';
    qryDestino.ParamByName('datahora_processamento').Value := vDataHora;



    Try
      Try
      qryDestino.ExecSQL;
      Except
      vArquivoTexto := 'C:\SEMF\AGRUPAMENTOS.TXT';
      GravaArquivoTexto(qryImportacao.fieldbyname('cnpj_cpf').Value);

      End;
    Finally
    End;


    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);





end;

procedure TfrmImportaICMS.Atualiza_SimplesNacional;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select serie, cnpj_cpf, extract(YEAR from data_emissao) as ano, ');
  qryImportacao.SQL.Add('extract(MONTH from data_emissao) as mes, tomador_simples ');
  qryImportacao.SQL.Add('from nota_fiscal where cnpj_cpf <> '''' and serie = ''NF'' and situacao = ''N''  ');
  qryImportacao.SQL.Add('group by serie, cnpj_cpf, ano, mes, tomador_simples             ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  while not qryImportacao.eof do
    begin

    vNomeMes := '';
    vCont := vCont + 1;

    Status.Panels[1].Text := '2-Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryDestino.Close;
    qryDestino.SQL.Clear;
    qryDestino.SQL.Add('update notafiscal_agrupada set tomador_simples =:tsn ');
    qryDestino.SQL.Add('where cnpj_cpf =:cnpj_cpf                            ');
    qryDestino.SQL.Add('  and ano =:ano and mes =:mes   ');
    qryDestino.ParamByName('tsn').Value      := qryImportacao.fieldbyname('tomador_simples').Value;
    qryDestino.ParamByName('cnpj_cpf').Value := qryImportacao.fieldbyname('cnpj_cpf').Value;
    qryDestino.ParamByName('ano').Value      := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value      := qryImportacao.fieldbyname('mes').Value;

    qryDestino.ExecSQL;



    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);


end;

procedure TfrmImportaICMS.Atualiza_RecolherRetido;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select serie, cnpj_cpf, tipo_recolhimento, sum(valor_notafiscal) as valor_notafiscal,  ');
  qryImportacao.SQL.Add('extract(YEAR from data_emissao) as ano, extract(MONTH from data_emissao) as mes ');
  qryImportacao.SQL.Add('from nota_fiscal where tipo_recolhimento <> '''' and serie = ''NF'' and situacao = ''N''  ');
  qryImportacao.SQL.Add('group by serie, cnpj_cpf, tipo_recolhimento, ano, mes                                  ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  while not qryImportacao.eof do
    begin

    vNomeMes := '';
    vCont := vCont + 1;

    Status.Panels[1].Text := '3-Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryDestino.Close;
    qryDestino.SQL.Clear;
    qryDestino.SQL.Add('update notafiscal_agrupada set  ');

    if Trim(qryImportacao.fieldbyname('tipo_recolhimento').AsString) = 'A' then
      qryDestino.SQL.Add(' valor_recolher =:recolher      ')
    else if Trim(qryImportacao.fieldbyname('tipo_recolhimento').AsString) = 'R' then
      qryDestino.SQL.Add(' valor_retido =:retido ');


    qryDestino.SQL.Add('where cnpj_cpf =:cnpj_cpf       ');
    qryDestino.SQL.Add('  and ano =:ano and mes =:mes   ');

    if Trim(qryImportacao.fieldbyname('tipo_recolhimento').AsString) = 'A' then
      qryDestino.ParamByName('recolher').Value := qryImportacao.fieldbyname('valor_notafiscal').Value
    else if Trim(qryImportacao.fieldbyname('tipo_recolhimento').AsString) = 'R' then
      qryDestino.ParamByName('retido').Value   := qryImportacao.fieldbyname('valor_notafiscal').Value;

    qryDestino.ParamByName('cnpj_cpf').Value := qryImportacao.fieldbyname('cnpj_cpf').Value;
    qryDestino.ParamByName('ano').Value      := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value      := qryImportacao.fieldbyname('mes').Value;

    qryDestino.ExecSQL;



    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);


end;

procedure TfrmImportaICMS.Agrupa_CartaoCredito;
begin

  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select cnpj_mf as cnpj_credenciado,nome_credenciado, sum(valor_operacao) as valor_operacao, ');
  qryImportacao.SQL.Add('extract(YEAR from data) as ano, extract(MONTH from data) as mes                             ');
  qryImportacao.SQL.Add('from registro65 where totalizado = ''f''                                                    ');

  qryImportacao.SQL.Add('  AND data >= ''2013-10-01''                                                                ');
  qryImportacao.SQL.Add('  AND data <= ''2014-12-31''                                                                ');

  qryImportacao.SQL.Add('group by cnpj_mf,nome_credenciado, ano, mes                                                 ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  while not qryImportacao.eof do
    begin

    vNomeMes := '';
    case qryImportacao.fieldbyname('mes').Value of
      1:  vNomeMes := 'JANEIRO';
      2:  vNomeMes := 'FEVEREIRO';
      3:  vNomeMes := 'MARCO';
      4:  vNomeMes := 'ABRIL';
      5:  vNomeMes := 'MAIO';
      6:  vNomeMes := 'JUNHO';
      7:  vNomeMes := 'JULHO';
      8:  vNomeMes := 'AGOSTO';
      9:  vNomeMes := 'SETEMBRO';
      10: vNomeMes := 'OUTUBRO';
      11: vNomeMes := 'NOVEMBRO';
      12: vNomeMes := 'DEZEMBRO';
    end;

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;


    qryDestino.Close;
    qryDestino.SQL.Clear;
    qryDestino.SQL.Add('INSERT INTO agrupamento(id, cnpj_credenciado, nome_credenciado, valor_operacao, ano, mes, nome_mes, data) ');
    qryDestino.SQL.Add('VALUES (nextval(''agrupamento_id_seq''), :cnpj_credenciado, :nome_credenciado, :valor_operacao, :ano, :mes, :nome_mes,:data) ');
    qryDestino.ParamByName('cnpj_credenciado').Value := qryImportacao.fieldbyname('cnpj_credenciado').Value;
    qryDestino.ParamByName('nome_credenciado').Value := AnsiToAscii(qryImportacao.fieldbyname('nome_credenciado').Value);
    qryDestino.ParamByName('valor_operacao').Value   := qryImportacao.fieldbyname('valor_operacao').Value;
    qryDestino.ParamByName('ano').Value              := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value              := qryImportacao.fieldbyname('mes').Value;
    qryDestino.ParamByName('nome_mes').Value         := vNomeMes;
    qryDestino.ParamByName('data').Value             := qryImportacao.fieldbyname('ano').Asstring+'-'+
                                                        qryImportacao.fieldbyname('mes').AsString+'-01';

    Try
      Try
      qryDestino.ExecSQL;
      Except
      vArquivoTexto := 'C:\SEMF\AGRUPA_CARTAO.TXT';
      GravaArquivoTexto(qryImportacao.fieldbyname('cnpj_credenciado').Value);

      End;
    Finally
    End;

    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);

  showmessage('FINAL...');

end;

procedure TfrmImportaICMS.Agrupa_Administradora;
var
  VlrTaxa, VlrIss : Double;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select r10.cnpj_mf as cnpj_administradora,                ');
  qryImportacao.SQL.Add('r10.nome_administradora,                                  ');
  qryImportacao.SQL.Add('extract(MONTH from r65.data) as mes, extract(YEAR from r65.data) as ano, ');
  qryImportacao.SQL.Add('sum(r65.valor_operacao) as valoroperacao                 ');
  qryImportacao.SQL.Add('from registro10 r10                                      ');
  qryImportacao.SQL.Add('inner join registro65 r65 on r10.id = r65.registro10_id  ');

//qryImportacao.SQL.Add('WHERE extract(YEAR from r65.data) = 2013   ');
//qryImportacao.SQL.Add('  and r10.cnpj_mf = ''01425787000104''     ');

  qryImportacao.SQL.Add('group by r10.cnpj_mf, r10.nome_administradora, ano, mes  ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  while not qryImportacao.eof do
    begin

    vNomeMes := '';
    case qryImportacao.fieldbyname('mes').Value of
      1:  vNomeMes := 'JANEIRO';
      2:  vNomeMes := 'FEVEREIRO';
      3:  vNomeMes := 'MARCO';
      4:  vNomeMes := 'ABRIL';
      5:  vNomeMes := 'MAIO';
      6:  vNomeMes := 'JUNHO';
      7:  vNomeMes := 'JULHO';
      8:  vNomeMes := 'AGOSTO';
      9:  vNomeMes := 'SETEMBRO';
      10: vNomeMes := 'OUTUBRO';
      11: vNomeMes := 'NOVEMBRO';
      12: vNomeMes := 'DEZEMBRO';
    end;

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    VlrTaxa := 0; VlrIss := 0;

    qryDestino.Close;
    qryDestino.SQL.Clear;

    qryDestino.SQL.Add('INSERT INTO administradora_agrupada(id, cnpj_administradora, nome_administradora,   ');
    qryDestino.SQL.Add(' valor_operacao, valor_taxa, valor_iss, ano, mes, nome_mes, data)                   ');
    qryDestino.SQL.Add('VALUES (nextval(''administradora_agrupada_id_seq''), :cnpj_administradora, :nome_administradora, ');
    qryDestino.SQL.Add(' :valor_operacao, :valor_taxa, :valor_iss, :ano, :mes, :nome_mes, :data)       ');

    qryDestino.ParamByName('cnpj_administradora').Value  := qryImportacao.fieldbyname('cnpj_administradora').Value;
    qryDestino.ParamByName('nome_administradora').Value  := qryImportacao.fieldbyname('nome_administradora').Value;
    qryDestino.ParamByName('valor_operacao').Value       := qryImportacao.fieldbyname('valoroperacao').Value;

    VlrTaxa := qryImportacao.fieldbyname('valoroperacao').Value * 0.03;
    VlrIss := VlrTaxa * 0.05;

    qryDestino.ParamByName('valor_taxa').Value    := VlrTaxa;
    qryDestino.ParamByName('valor_iss').Value     := VlrIss;

    qryDestino.ParamByName('ano').Value              := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value              := qryImportacao.fieldbyname('mes').Value;
    qryDestino.ParamByName('nome_mes').Value         := vNomeMes;
    qryDestino.ParamByName('data').Value             := qryImportacao.fieldbyname('ano').Asstring+'-'+
                                                        qryImportacao.fieldbyname('mes').AsString+'-01';


    Try
      Try
      qryDestino.ExecSQL;
      Except
      vArquivoTexto := 'C:\SEMF\AGRUPA_ADMINISTRADORA.TXT';
      GravaArquivoTexto(qryImportacao.fieldbyname('cnpj_administradora').Value);

      End;
    Finally
    End;


    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);

  showmessage('FINAL...');

end;

procedure TfrmImportaICMS.Atualiza_Totalizados;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select r65.data, r65.valor_operacao, r66.montante_cartao_credito,   ');
  qryImportacao.SQL.Add('r66.montante_cartao_debito, r65.natureza_operacao, r65.tipo_operacao, r65.registro10_id,  ');
  qryImportacao.SQL.Add('r65.nome_credenciado, r65.id, r65.cnpj_mf ');
  qryImportacao.SQL.Add('from registro65 r65, registro66 r66                                 ');
  qryImportacao.SQL.Add('where r65.valor_operacao = r66.montante_cartao_credito              ');
  qryImportacao.SQL.Add('  and r65.registro10_id = r66.registro10_id                         ');
  qryImportacao.SQL.Add('  and r65.cnpj_mf = r66.cnpj_mf                                     ');
  qryImportacao.SQL.Add('  order by r65.data                                                 ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;

  while not qryImportacao.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryVerifica.Close;
    qryVerifica.SQL.Clear;
    qryVerifica.SQL.Add('select id from registro10 where id =:id and cnpj_mf =:cnpj ');
    qryVerifica.ParamByName('id').Value   := qryImportacao.fieldbyname('registro10_id').Value;
    qryVerifica.ParamByName('cnpj').Value := '01027058000191';
    qryVerifica.open;

    if qryVerifica.RecordCount > 0 then
      begin

      qryDestino.close;
      qryDestino.sql.Clear;
      qryDestino.sql.add('update registro65 set totalizado =:tot  ');
      qryDestino.sql.add('where registro10_id =:id and valor_operacao =:vlr     ');
      qryDestino.ParamByName('tot').Value := True;
      qryDestino.ParamByName('id').Value  := qryImportacao.fieldbyname('registro10_id').Value;
      qryDestino.ParamByName('vlr').Value := qryImportacao.fieldbyname('montante_cartao_credito').Value;
      qryDestino.ExecSQL;

      qryDestino.close;
      qryDestino.sql.Clear;
      qryDestino.sql.add('update registro65 set totalizado =:tot  ');
      qryDestino.sql.add('where registro10_id =:id and valor_operacao =:vlr     ');
      qryDestino.ParamByName('tot').Value := True;
      qryDestino.ParamByName('id').Value  := qryImportacao.fieldbyname('registro10_id').Value;
      qryDestino.ParamByName('vlr').Value := qryImportacao.fieldbyname('montante_cartao_debito').Value;
      qryDestino.ExecSQL;


      end;
    qryImportacao.Next;
    end;//while ...

  lblFim.Caption := 'FINAL: '+timetostr(now);

  showmessage('FINAL...');

end;

procedure TfrmImportaICMS.Processa_RegistroPagamento;
begin

  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select x.codtrb,sum(x.valorTotal) as valortotal,sum(x.Juros) as juros,sum(x.Multa) as multa,');
  qryBuscaSIAT.SQL.Add('sum(x.ValAtu) as valorAtualizado, sum(x.vlrLancado) as vlrLanc, count(*) as quantidade,     ');
  qryBuscaSIAT.SQL.Add(' case when cad.tipcad = ''I'' then ''IMOBILIARIO''                                          ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''M'' then ''MOBILIARIO''                                           ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''P'' then ''PESSOA''                                               ');
  qryBuscaSIAT.SQL.Add(' else ''IMOBILIARIO'' end as tpcad, sum(x.Desconto) as desconto,                            ');

  qryBuscaSIAT.SQL.Add('  TO_CHAR(x.dtMov,''YYYY'') as ano, TO_CHAR(x.dtMov,''MM'') as mes                          ');

  qryBuscaSIAT.SQL.Add('	from (select cad.codcad codcad, cad.tipcad tipcad,t.codtrb,                               ');
  qryBuscaSIAT.SQL.Add('	t.desmin as descricao, m.DATMVA as dtMov,  count(*) QtdParcelas,                          ');
  qryBuscaSIAT.SQL.Add('	dpp.identi as identi, sum(dpp.vallanmoe) vlrLancado,                                      ');
  qryBuscaSIAT.SQL.Add('	sum(dpp.vallanmoe) + sum(case when da.codtrb is not null then d.valdoc else 0 end) ValLan,');
  qryBuscaSIAT.SQL.Add('	sum(dpp.vallanmoe + dpp.atumon + dpp.jurfin) ValAtu,                                      ');
  qryBuscaSIAT.SQL.Add('	sum(dpp.jurmor) Juros, sum(dpp.mulmor) Multa, sum(dpp.descon) Desconto,                   ');
  qryBuscaSIAT.SQL.Add('	sum(dp.valemodca) Emolumento, sum(dpp.valpago) valorTotal                                 ');
  qryBuscaSIAT.SQL.Add('	from SIATTHE.TBLMVA m                                                                     ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLMVALTA ml                                                           ');
  qryBuscaSIAT.SQL.Add('	    on m.codmva = ml.codmva                                                               ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLDCM d                                                               ');
  qryBuscaSIAT.SQL.Add('	    on ml.codmvalta = d.codmvalta                                                         ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMPAG dp                                                            ');
  qryBuscaSIAT.SQL.Add('	on d.coddcm = dp.coddcm                                                                   ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMPAGPAR dpp                                                        ');
  qryBuscaSIAT.SQL.Add('	    on dp.coddcmpag = dpp.coddcmpag                                                       ');
  qryBuscaSIAT.SQL.Add('	left join siatthe.tblcad cad                                                              ');
  qryBuscaSIAT.SQL.Add('	    on dpp.codcad = cad.codcad                                                            ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMAJS da                                                            ');
  qryBuscaSIAT.SQL.Add('	    on (d.coddcm = da.coddcm)                                                             ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLTRB t                                                               ');
  qryBuscaSIAT.SQL.Add('	    on (t.codtrb = dpp.codtrb or t.codtrb = da.codtrb)                                    ');
 {
  qryBuscaSIAT.SQL.Add('      inner join SIATTHE.tbltrbgtr tg     ');
  qryBuscaSIAT.SQL.Add('    on t.codtrb = tg.codtrb               ');
  qryBuscaSIAT.SQL.Add('  inner join SIATTHE.tblgtr g             ');
  qryBuscaSIAT.SQL.Add('    on g.codgtr = tg.codgtr               ');
   }

  qryBuscaSIAT.SQL.Add('	group by t.codtrb, t.desmin,  cad.tipcad, cad.codcad, m.DATMVA, dpp.identi        ');
  qryBuscaSIAT.SQL.Add('	) x left join siatthe.tblcad cad on cad.codcad = x.codcad                                 ');

//   --     where TO_DATE(x.dtMov,'DD/MM/YYYY') between '05/08/2013' and '06/08/2013'

  //qryBuscaSIAT.SQL.Add('WHERE x.dtMov between to_date (''01/01/2013'', ''DD/MM/YYYY'')    ');
 // qryBuscaSIAT.SQL.Add('  AND to_date (''10/01/2013'', ''DD/MM/YYYY'')                    ');

  qryBuscaSIAT.SQL.Add('group by                                                                                    ');

  qryBuscaSIAT.SQL.Add(' case when cad.tipcad = ''I'' then ''IMOBILIARIO''                                          ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''M'' then ''MOBILIARIO''                                           ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''P'' then ''PESSOA''                                               ');
  qryBuscaSIAT.SQL.Add(' else ''IMOBILIARIO'' end,                                                                  ');

  qryBuscaSIAT.SQL.Add(' x.codtrb, TO_CHAR(x.dtMov,''YYYY''), TO_CHAR(x.dtMov,''MM'')                               ');


  qryBuscaSIAT.SQL.Add('        order by x.codtrb, TO_CHAR(x.dtMov,''YYYY''), TO_CHAR(x.dtMov,''MM'')               ');

  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;
  vDataHora := now;

  if qryBuscaSIAT.RecordCount > 0 then
    begin
    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('truncate table registro_pagamento   ');
    qryRegistro.ExecSQL;
    end;

  while not qryBuscaSIAT.eof do
    begin

    vNomeMes := '';
    case qryBuscaSIAT.fieldbyname('mes').Value of
      1:  vNomeMes := 'JANEIRO';
      2:  vNomeMes := 'FEVEREIRO';
      3:  vNomeMes := 'MARCO';
      4:  vNomeMes := 'ABRIL';
      5:  vNomeMes := 'MAIO';
      6:  vNomeMes := 'JUNHO';
      7:  vNomeMes := 'JULHO';
      8:  vNomeMes := 'AGOSTO';
      9:  vNomeMes := 'SETEMBRO';
      10: vNomeMes := 'OUTUBRO';
      11: vNomeMes := 'NOVEMBRO';
      12: vNomeMes := 'DEZEMBRO';
    end;

    vCont := vCont + 1;

    vCodGtr := '';
    qryDados.Close;
    qryDados.SQL.Clear;
    qryDados.SQL.Add('select codgtr from SIATTHE.tbltrbgtr where codtrb =:cod order by datalt desc    ');
    qryDados.Parameters.ParamByName('cod').Value := qryBuscaSIAT.fieldbyname('codtrb').asstring;
    qryDados.open;
    vCodGtr := qryDados.fieldbyname('codgtr').AsString;


    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO registro_pagamento(id, tipo_cadastro, tributo_id, valor_pago, ano, mes, data,                    ');
    qryRegistro.SQL.Add('  usuario_id, datahora_processamento, nome_mes, codtrb, grupo_tributo_id,                                    ');
    qryRegistro.SQL.Add('  valor_atualizado, juros, multa, valor_lancado, quantidade, desconto)                                       ');
    qryRegistro.SQL.Add('VALUES (nextval(''registro_pagamento_id_seq''), :tipo_cadastro, :tributo_id, :valor_pago, :ano, :mes, :data, ');
    qryRegistro.SQL.Add('  :usuario_id, :datahora_processamento, :nome_mes, :codtrb, :grupo,                                          ');
    qryRegistro.SQL.Add('  :valor_atualizado, :juros, :multa, :valor_lancado, :quantidade, :desconto)                                 ');

    qryRegistro.ParamByName('tipo_cadastro').Asstring := qryBuscaSIAT.fieldbyname('tpcad').Asstring;
    qryRegistro.ParamByName('valor_pago').AsFloat  := qryBuscaSIAT.fieldbyname('valortotal').AsFloat;
    qryRegistro.ParamByName('tributo_id').Value    := TrazID_Tributo(qryBuscaSIAT.fieldbyname('codtrb').AsString);
    qryRegistro.ParamByName('ano').Value           := qryBuscaSIAT.fieldbyname('ano').Value;
    qryRegistro.ParamByName('mes').Value           := qryBuscaSIAT.fieldbyname('mes').Value;
    qryRegistro.ParamByName('data').Value          := qryBuscaSIAT.fieldbyname('ano').Asstring+'-'+
                                                      qryBuscaSIAT.fieldbyname('mes').AsString+'-01';
    qryRegistro.ParamByName('usuario_id').Value    := 0;
    qryRegistro.ParamByName('datahora_processamento').Value := vDataHora;
    qryRegistro.ParamByName('nome_mes').Value    := vNomeMes;
    qryRegistro.ParamByName('codtrb').Value      := Trim(qryBuscaSIAT.fieldbyname('codtrb').AsString);
    qryRegistro.ParamByName('grupo').Value       := TrazID_GrupoTributo(vCodGtr);
    qryRegistro.ParamByName('valor_atualizado').AsFloat  := qryBuscaSIAT.fieldbyname('valorAtualizado').AsFloat;
    qryRegistro.ParamByName('juros').AsFloat       := qryBuscaSIAT.fieldbyname('juros').AsFloat;
    qryRegistro.ParamByName('multa').AsFloat       := qryBuscaSIAT.fieldbyname('multa').AsFloat;
    qryRegistro.ParamByName('valor_lancado').AsFloat := qryBuscaSIAT.fieldbyname('vlrLanc').AsFloat;
    qryRegistro.ParamByName('quantidade').AsInteger  := qryBuscaSIAT.fieldbyname('quantidade').AsInteger;
    qryRegistro.ParamByName('desconto').AsFloat      := qryBuscaSIAT.fieldbyname('desconto').AsFloat;

    Try
      Try
      qryRegistro.ExecSQL;
      Except
      //vArquivoTexto := 'C:\SEMF\REGISTRO_PAGAMENTO.TXT';

      //GravaArquivoTexto(qryBuscaSIAT.fieldbyname('tpcad').Value+'-'+qryBuscaSIAT.fieldbyname('codtrb').Value+'-'+
      //                  qryBuscaSIAT.fieldbyname('ano').Value+'-'+qryBuscaSIAT.fieldbyname('mes').Value);

      End;
    Finally
    End;

    qryBuscaSIAT.Next;

    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

function TfrmImportaICMS.TrazID_Tributo(codtrb: string): integer;
begin
  Result := 0;

  if trim(codtrb) <> '' then
    begin
    qryTributo.Close;
    qryTributo.SQL.Clear;
    qryTributo.SQL.Add('select id from tributo   ');
//    qryTributo.SQL.Add('where cast( cast(codtrb as text) as integer)=:cod     ');
    qryTributo.SQL.Add('where codtrb=:cod     ');
    qryTributo.ParamByName('cod').Value := Trim(codtrb);
    qryTributo.open;

    if qryTributo.RecordCount > 0 then
      Result := qryTributo.fieldbyname('id').AsInteger;

    end;
end;

function TfrmImportaICMS.TrazID_GrupoTributo(codtrb: string): integer;
begin
  Result := 13;

  if trim(codtrb) <> '' then
    begin
    qryTributo.Close;
    qryTributo.SQL.Clear;
    qryTributo.SQL.Add('select id from grupo_tributo   ');
    qryTributo.SQL.Add('where codgtr=:cod     ');
    qryTributo.ParamByName('cod').Value := Trim(codtrb);
    qryTributo.open;

    if qryTributo.RecordCount > 0 then
      Result := qryTributo.fieldbyname('id').AsInteger;

    end;

end;

procedure TfrmImportaICMS.Atualiza_Pessoa_Siat;
begin


  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select cad.tipcad, cad.codimo, cad.codeco,cad.codcad,           ');
  qryBuscaSIAT.SQL.Add('pes.codpes, pes.tippes, pes.nomrazres, pes.docnum, pes.cpfcnpj  ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.tblcad cad, SIATTHE.tblpes pes                     ');
  qryBuscaSIAT.SQL.Add('where cad.codpes = pes.codpes order by pes.codpes               ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' Cod.Pes: '+qryBuscaSIAT.fieldbyname('codpes').AsString;
    Application.ProcessMessages;


    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO pessoa_siat(id, nome, documento, inscricaomunicipal, tipocadastro, tipopessoa,codcad)   ');
    qryRegistro.SQL.Add('VALUES (nextval(''pessoa_siat_id_seq''), :nome, :documento, :inscricaomunicipal, :tipocadastro, :tipopessoa,:codcad)  ');

    vNomePessoa := AnsiToAscii(Trim(qryBuscaSIAT.fieldbyname('nomrazres').AsString));
    vNomePessoa := TrocaCaracterEspecial(vNomePessoa,true);

    lblDestino.Caption := vNomePessoa;

    qryRegistro.ParamByName('nome').Value  := vNomePessoa;


    TrazDadosPessoaSiat(qryBuscaSIAT.fieldbyname('tipcad').AsString);

    qryRegistro.ParamByName('documento').Value      := Trim(vCpfCnpj);

    qryRegistro.ParamByName('inscricaomunicipal').Value := Trim(vInscMun);


    qryRegistro.ParamByName('tipocadastro').Value    := Trim(qryBuscaSIAT.fieldbyname('tipcad').AsString);
    qryRegistro.ParamByName('tipopessoa').Value      := Trim(qryBuscaSIAT.fieldbyname('tippes').AsString);
    qryRegistro.ParamByName('codcad').Value          := qryBuscaSIAT.fieldbyname('codcad').AsInteger;

    qryRegistro.ExecSQL;

    qryBuscaSIAT.next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

function TfrmImportaICMS.TrazDadosPessoaSiat(tipo:string): string;
begin

  Result := '';
  vCpfCnpj := '';
  vInscMun := '';

  if trim(tipo) <> 'I' then
    begin
    qryDados.Close;
    qryDados.SQL.Clear;
    qryDados.SQL.Add('select insmun, cpfcnpj from siatthe.tbleco   ');
    qryDados.SQL.Add('where codpes=:cod     ');
    qryDados.Parameters.ParamByName('cod').Value := qryBuscaSIAT.fieldbyname('codpes').asinteger;
    qryDados.open;

    if qryDados.RecordCount > 0 then
      begin
      vInscMun := qryDados.fieldbyname('insmun').AsString;
      vCpfCnpj := qryDados.fieldbyname('cpfcnpj').AsString;
      end;

    end
  else
    begin
    qryDados.Close;
    qryDados.SQL.Clear;
    qryDados.SQL.Add('select insimo from siatthe.tblimo   ');
    qryDados.SQL.Add('where codimo=:cod     ');
    qryDados.Parameters.ParamByName('cod').Value := qryBuscaSIAT.fieldbyname('codimo').asinteger;
    qryDados.open;

    if qryDados.RecordCount > 0 then
      begin
      vInscMun := qryDados.fieldbyname('insimo').AsString;
      vCpfCnpj := qryBuscaSIAT.fieldbyname('cpfcnpj').AsString;
      end;


    end;
end;

procedure TfrmImportaICMS.BitBtn3Click(Sender: TObject);
begin
Migra_TBLUOR;
//Atualiza_Pessoa_Siat;
end;

procedure TfrmImportaICMS.ProcessaPagamentos;
begin


  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select distinct cad.tipcad, x.codtrb, x.valorPago, x.dtMov, x.ValLan, x.Juros, x.Multa, ');
  qryBuscaSIAT.SQL.Add(' x.codcad, x.codlnc, x.ValAtu                                                           ');
  qryBuscaSIAT.SQL.Add('from (select cad.codcad codcad, cad.tipcad tipcad,t.codtrb,                             ');
  qryBuscaSIAT.SQL.Add('t.desmin as descricao, m.DATMVA as dtMov,                                               ');
  qryBuscaSIAT.SQL.Add('dpp.identi as identi,                                                                   ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe + case when da.codtrb is not null then d.valdoc else 0 end ValLan,        ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe + dpp.atumon + dpp.jurfin ValAtu,                                         ');
  qryBuscaSIAT.SQL.Add('dpp.jurmor Juros, dpp.mulmor Multa, dpp.descon Desconto,                                ');
  qryBuscaSIAT.SQL.Add('dp.valemodca Emolumento, dpp.valpago valorPago, dpp.codlnc                              ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.TBLMVA m                                                                   ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLMVALTA ml                                                         ');
  qryBuscaSIAT.SQL.Add('    on m.codmva = ml.codmva                                                             ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLDCM d                                                             ');
  qryBuscaSIAT.SQL.Add('    on ml.codmvalta = d.codmvalta                                                       ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAG dp                                                          ');
  qryBuscaSIAT.SQL.Add('on d.coddcm = dp.coddcm                                                                 ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAGPAR dpp                                                      ');
  qryBuscaSIAT.SQL.Add('    on dp.coddcmpag = dpp.coddcmpag                                                     ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblcad cad                                                            ');
  qryBuscaSIAT.SQL.Add('    on dpp.codcad = cad.codcad                                                          ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMAJS da                                                          ');
  qryBuscaSIAT.SQL.Add('    on (d.coddcm = da.coddcm)                                                           ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLTRB t                                                             ');
  qryBuscaSIAT.SQL.Add('    on (t.codtrb = dpp.codtrb or t.codtrb = da.codtrb)                                  ');
  qryBuscaSIAT.SQL.Add('	) x left join siatthe.tblcad cad on cad.codcad = x.codcad                             ');

 // qryBuscaSIAT.SQL.Add('where x.dtMov between to_date(''01/01/2013'', ''dd/MM/yyyy'') and to_date(''30/09/2013'', ''dd/MM/yyyy'')  ');

  qryBuscaSIAT.SQL.Add(' where x.valorPago > 0                                                                 ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' Cod.Cad: '+qryBuscaSIAT.fieldbyname('codcad').AsString;
    Application.ProcessMessages;


    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO pagamentos(id, tipocadastro, tributo_id, datavalidade, datalancamento, ');
    qryRegistro.SQL.Add(' datapagamento, valoratualizado, valorpago, valorlancado, valormulta, valorjuros,  ');
//    qryRegistro.SQL.Add(' tipodivida_id, pessoasiat_id, situacao_divida_id,     ');
    qryRegistro.SQL.Add(' faixa_id, grupo_tributo_id)     ');

    qryRegistro.SQL.Add('VALUES (nextval(''pagamentos_id_seq''), :tipocadastro, :tributo_id, :datavalidade, :datalancamento, ');
    qryRegistro.SQL.Add(' :datapagamento, :valoratualizado, :valorpago, :valorlancado, :valormulta, :valorjuros,  ');
//    qryRegistro.SQL.Add(' :tipodivida_id, :pessoasiat_id, :situacao,                            ');

    qryRegistro.SQL.Add('  :faixa_id, :grupo)                           ');

    qryRegistro.ParamByName('tipocadastro').Value     := qryBuscaSIAT.fieldbyname('tipcad').AsString;
    qryRegistro.ParamByName('tributo_id').Value       := TrazID_Tributo(qryBuscaSIAT.fieldbyname('codtrb').AsString);

    TrazDatas; //Traz data lançamento pelo codlnc na tabela siatthe.tbllnc

    if qryDados.fieldbyname('datven').Asstring <> '' then
      qryRegistro.ParamByName('datavalidade').Value     := qryDados.fieldbyname('datven').AsDateTime
    else
      qryRegistro.ParamByName('datavalidade').Value     := null;

    if qryDados.fieldbyname('datbla').AsString <> '' then
      qryRegistro.ParamByName('datalancamento').Value   := qryDados.fieldbyname('datbla').AsDateTime
    else
      qryRegistro.ParamByName('datalancamento').Value   := null;

    if qryDados.fieldbyname('datpag').Asstring <> '' then
      qryRegistro.ParamByName('datapagamento').Value    := qryDados.fieldbyname('datpag').AsDateTime
    else
      qryRegistro.ParamByName('datapagamento').Value    := null;

    qryRegistro.ParamByName('valoratualizado').Value  := qryBuscaSIAT.fieldbyname('ValAtu').AsFloat;

    qryRegistro.ParamByName('valorpago').Value        := qryBuscaSIAT.fieldbyname('valorPago').AsFloat;
    qryRegistro.ParamByName('valorlancado').Value     := qryBuscaSIAT.fieldbyname('ValLan').AsFloat;

    qryRegistro.ParamByName('valormulta').Value       := qryBuscaSIAT.fieldbyname('Multa').AsFloat;
    qryRegistro.ParamByName('valorjuros').Value       := qryBuscaSIAT.fieldbyname('Juros').AsFloat;

    //qryRegistro.ParamByName('tipodivida_id').Value    := qryBuscaSIAT.fieldbyname('').AsInteger;

    //qryRegistro.ParamByName('pessoasiat_id').Value    := qryBuscaSIAT.fieldbyname('').AsInteger;
    //qryRegistro.ParamByName('situacao').Value         := qryBuscaSIAT.fieldbyname('').AsInteger;

    if qryBuscaSIAT.fieldbyname('ValLan').AsFloat <=50 then
      qryRegistro.ParamByName('faixa_id').Value := 1
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 50) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 100) then
      qryRegistro.ParamByName('faixa_id').Value := 2
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 100) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 200) then
      qryRegistro.ParamByName('faixa_id').Value := 3
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 200) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 500) then
      qryRegistro.ParamByName('faixa_id').Value := 4
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 500) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 1000) then
      qryRegistro.ParamByName('faixa_id').Value := 5
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 1000) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 5000) then
      qryRegistro.ParamByName('faixa_id').Value := 6
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 5000) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 10000) then
      qryRegistro.ParamByName('faixa_id').Value := 7
    else if (qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 10000) and (qryBuscaSIAT.fieldbyname('ValLan').AsFloat <= 50000) then
      qryRegistro.ParamByName('faixa_id').Value := 8
    else if qryBuscaSIAT.fieldbyname('ValLan').AsFloat > 50000 then
      qryRegistro.ParamByName('faixa_id').Value := 9;

    qryRegistro.ParamByName('grupo').Value    := TrazID_GrupoTributo(vCodGtr);


    qryRegistro.ExecSQL;


    qryBuscaSIAT.next;
    end; //while


end;

procedure TfrmImportaICMS.Migra_TBLITR;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select coditr,tipcad,tipitr,descom, desres, desred, sigla, desmin  ');
  qryBuscaSIAT.SQL.Add('from siatthe.tblitr order by coditr                                ');

  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryBuscaSIAT.Next;
    end;//while

end;

procedure TfrmImportaICMS.TrazDatas;
begin
   {
  vCodGtr := '';
  qryDados.Close;
  qryDados.SQL.Clear;
  qryDados.SQL.Add('select codgtr from SIATTHE.tbltrbgtr where codtrb =:cod order by datalt desc    ');
  qryDados.Parameters.ParamByName('cod').Value := qryBusca.fieldbyname('codtrb').asstring;
  qryDados.open;
  vCodGtr := qryDados.fieldbyname('codgtr').AsString; }


  qryDados.Close;
  qryDados.SQL.Clear;
  qryDados.SQL.Add('select par.datpag, lnc.datbla, par.datven      ');
  qryDados.SQL.Add('from siatthe.tbllnc lnc, siatthe.tbllncpar par ');
  qryDados.SQL.Add('where lnc.codlnc = par.codlnc                  ');
  qryDados.SQL.Add('  and lnc.codlnc =:cod                         ');
  qryDados.Parameters.ParamByName('cod').Value := qryBuscaSIAT.fieldbyname('codlnc').asinteger;
  qryDados.open;

end;

procedure TfrmImportaICMS.ProcessaRelatorioArrecadacao;
begin


  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select distinct x.codtrb, x.valorPago, x.dtMov, x.ValLan, x.Juros, x.Multa, ');

  qryBuscaSIAT.SQL.Add(' case when cad.tipcad = ''I'' then ''IMOBILIARIO''                                          ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''M'' then ''MOBILIARIO''                                           ');
  qryBuscaSIAT.SQL.Add('      when cad.tipcad = ''P'' then ''PESSOA''                                               ');
  qryBuscaSIAT.SQL.Add(' else ''IMOBILIARIO'' end as tpcad,                                                         ');

  qryBuscaSIAT.SQL.Add(' x.codcad, x.codlnc, x.ValAtu, x.Desconto                                               ');
  qryBuscaSIAT.SQL.Add('from (select cad.codcad codcad, cad.tipcad tipcad,t.codtrb,                             ');
  qryBuscaSIAT.SQL.Add('t.desmin as descricao, m.DATMVA as dtMov,                                               ');
  qryBuscaSIAT.SQL.Add('dpp.identi as identi,                                                                   ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe + case when da.codtrb is not null then d.valdoc else 0 end ValLan,        ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe + dpp.atumon + dpp.jurfin ValAtu,                                         ');
  qryBuscaSIAT.SQL.Add('dpp.jurmor Juros, dpp.mulmor Multa, dpp.descon Desconto,                                ');
  qryBuscaSIAT.SQL.Add('dp.valemodca Emolumento, dpp.valpago valorPago, dpp.codlnc                              ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.TBLMVA m                                                                   ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLMVALTA ml                                                         ');
  qryBuscaSIAT.SQL.Add('    on m.codmva = ml.codmva                                                             ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLDCM d                                                             ');
  qryBuscaSIAT.SQL.Add('    on ml.codmvalta = d.codmvalta                                                       ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAG dp                                                          ');
  qryBuscaSIAT.SQL.Add('on d.coddcm = dp.coddcm                                                                 ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAGPAR dpp                                                      ');
  qryBuscaSIAT.SQL.Add('    on dp.coddcmpag = dpp.coddcmpag                                                     ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblcad cad                                                            ');
  qryBuscaSIAT.SQL.Add('    on dpp.codcad = cad.codcad                                                          ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMAJS da                                                          ');
  qryBuscaSIAT.SQL.Add('    on (d.coddcm = da.coddcm)                                                           ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLTRB t                                                             ');
  qryBuscaSIAT.SQL.Add('    on (t.codtrb = dpp.codtrb or t.codtrb = da.codtrb)                                  ');
  qryBuscaSIAT.SQL.Add('	) x left join siatthe.tblcad cad on cad.codcad = x.codcad                             ');
  qryBuscaSIAT.SQL.Add(' where x.codcad > 0 and x.valorPago > 0                                                 ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    TrazDatas;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' Cod.Cad: '+qryBuscaSIAT.fieldbyname('codcad').AsString;
    lblFim.Caption := '...: '+timetostr(now)+' -> '+datetostr(date);
    Application.ProcessMessages;


    qryRegistro.Close;
    qryRegistro.SQL.Clear;

    qryRegistro.SQL.Add('INSERT INTO relatorio_arrecadacao(id, tipocadastro, tributo_id, valor_total, valor_principal, ');
    qryRegistro.SQL.Add(' juros, multa, desconto, faixa_id, grupo_tributo_id, data_processamento, codcad,              ');
    qryRegistro.SQL.Add(' datavalidade, datalancamento, datapagamento)                                                 ');
    qryRegistro.SQL.Add('VALUES (nextval(''relatorio_arrecadacao_id_seq''), :tipocadastro, :tributo_id, :valor_total, :valor_principal, ');
    qryRegistro.SQL.Add(' :juros, :multa, :desconto, :faixa_id, :grupo, :data_processamento, :codcad,              ');
    qryRegistro.SQL.Add(' :datavalidade, :datalancamento, :datapagamento)                                          ');

    qryRegistro.ParamByName('tipocadastro').Value     := qryBuscaSIAT.fieldbyname('tpcad').AsString;
    qryRegistro.ParamByName('tributo_id').Value       := TrazID_Tributo(qryBuscaSIAT.fieldbyname('codtrb').AsString);

    qryRegistro.ParamByName('valor_total').AsFloat    := qryBuscaSIAT.fieldbyname('valorPago').AsFloat;
    qryRegistro.ParamByName('valor_principal').Value  := qryBuscaSIAT.fieldbyname('ValLan').AsFloat;

    qryRegistro.ParamByName('juros').Value            := qryBuscaSIAT.fieldbyname('Juros').AsFloat;
    qryRegistro.ParamByName('multa').Value            := qryBuscaSIAT.fieldbyname('Multa').AsFloat;
    qryRegistro.ParamByName('desconto').Value         := qryBuscaSIAT.fieldbyname('Desconto').AsFloat;


    if qryBuscaSIAT.fieldbyname('valorPago').AsFloat <=50 then
      qryRegistro.ParamByName('faixa_id').Value := 1
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 50) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 100) then
      qryRegistro.ParamByName('faixa_id').Value := 2
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 100) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 200) then
      qryRegistro.ParamByName('faixa_id').Value := 3
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 200) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 500) then
      qryRegistro.ParamByName('faixa_id').Value := 4
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 500) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 1000) then
      qryRegistro.ParamByName('faixa_id').Value := 5
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 1000) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 5000) then
      qryRegistro.ParamByName('faixa_id').Value := 6
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 5000) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 10000) then
      qryRegistro.ParamByName('faixa_id').Value := 7
    else if (qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 10000) and (qryBuscaSIAT.fieldbyname('valorPago').AsFloat <= 50000) then
      qryRegistro.ParamByName('faixa_id').Value := 8
    else if qryBuscaSIAT.fieldbyname('valorPago').AsFloat > 50000 then
      qryRegistro.ParamByName('faixa_id').Value := 9;

    qryRegistro.ParamByName('grupo').Value                 := TrazID_GrupoTributo(vCodGtr);
    qryRegistro.ParamByName('data_processamento').Value    := now;
    qryRegistro.ParamByName('codcad').Value                := qryBuscaSIAT.fieldbyname('codcad').AsString;

    if qryDados.fieldbyname('datven').Asstring <> '' then
      qryRegistro.ParamByName('datavalidade').Value     := qryDados.fieldbyname('datven').AsDateTime
    else
      qryRegistro.ParamByName('datavalidade').Value     := null;

    if qryDados.fieldbyname('datbla').AsString <> '' then
      qryRegistro.ParamByName('datalancamento').Value   := qryDados.fieldbyname('datbla').AsDateTime
    else
      qryRegistro.ParamByName('datalancamento').Value   := null;

    if qryDados.fieldbyname('datpag').Asstring <> '' then
      qryRegistro.ParamByName('datapagamento').Value    := qryDados.fieldbyname('datpag').AsDateTime
    else
      qryRegistro.ParamByName('datapagamento').Value    := null;

    qryRegistro.ExecSQL;

    qryBuscaSIAT.next;
    end; //while

end;

procedure TfrmImportaICMS.Migra_TBLUOR;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select coduor, codorg, nomcom from siatthe.tbluor order by coduor  ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO tbluor(id, coduor, codorg, nomcom)                  ');
    qryRegistro.SQL.Add('VALUES (nextval(''tbluor_id_seq''), :coduor, :codorg, :nomcom)  ');

    vNomePessoa := '';
    vNomePessoa := AnsiToAscii(Trim(qryBuscaSIAT.fieldbyname('nomcom').AsString));
    vNomePessoa := TrocaCaracterEspecial(vNomePessoa,true);


    qryRegistro.ParamByName('coduor').Value := Trim(qryBuscaSIAT.fieldbyname('coduor').AsString);
    qryRegistro.ParamByName('codorg').Value := Trim(qryBuscaSIAT.fieldbyname('codorg').AsString);
    qryRegistro.ParamByName('nomcom').Value := Trim(vNomePessoa);


    qryRegistro.ExecSQL;


    qryBuscaSIAT.Next;
    end;//while
   showmessage('Fim...');
end;

procedure TfrmImportaICMS.Processa_RelatorioCreditoGeral;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select                                                                                                     ');
  qryBuscaSIAT.SQL.Add(' case when cad.tipcad = ''I'' then ''IMOBILIARIO''                                                         ');
  qryBuscaSIAT.SQL.Add('     when cad.tipcad = ''M'' then ''MOBILIARIO''                                                           ');
  qryBuscaSIAT.SQL.Add('     when cad.tipcad = ''P'' then ''PESSOA''                                                               ');
  qryBuscaSIAT.SQL.Add(' else ''IMOBILIARIO'' end as tipoCadastro,                                                                 ');
  qryBuscaSIAT.SQL.Add(' trb.codtrb, trb.desmin,                                                                                   ');
  qryBuscaSIAT.SQL.Add(' lp.vallanmoe as valorLancado, calc.valatu as valorAtualizado, calc.valMul as multa,                       ');
  qryBuscaSIAT.SQL.Add(' calc.valJur as juros,calc.valDes as desconto, calc.valtot as valorTotal,                                  ');
  qryBuscaSIAT.SQL.Add(' lp.datven as dataValidade, lp.datpag as dataPagamento,lp.datbla as dataLancamento,                        ');
  qryBuscaSIAT.SQL.Add(' lp.ajuiza, lp.divatv,lp.exectd, lp.situac, cad.codcad,                                                    ');
//  qryBuscaSIAT.SQL.Add(' trunc((months_between(sysdate, lp.datven))/12) as anos, ');

  //FAIXA
  qryBuscaSIAT.SQL.Add(' case when lp.vallanmoe <= 50  then 1                                                                      ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 50 and lp.vallanmoe <= 100  then 2                                               ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 100 and lp.vallanmoe <= 200 then 3                                               ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 200 and lp.vallanmoe <= 500 then 4                                               ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 500 and lp.vallanmoe <= 1000 then 5                                              ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 1000 and lp.vallanmoe <= 5000 then 6                                             ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 5000 and lp.vallanmoe <= 10000 then 7                                            ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 10000 and lp.vallanmoe <= 50000 then 8                                           ');
  qryBuscaSIAT.SQL.Add('      when lp.vallanmoe > 50000 then 9 end as faixa, l.codlnc,                                             ');

  //TIPO DÍVIDA
  qryBuscaSIAT.SQL.Add(' case when lp.exectd  = ''S''  then 5                                             '); //DIVIDA EXECUTADA

  qryBuscaSIAT.SQL.Add('      when lp.exectd  = ''N'' and lp.ajuiza = ''S'' then 4                        '); //DIVIDA AJUIZADA

  qryBuscaSIAT.SQL.Add('      when lp.exectd  = ''N'' and lp.ajuiza = ''N'' and lp.divatv = ''S'' then 3  '); //DIVIDA ATIVA

  qryBuscaSIAT.SQL.Add('      when (lp.ajuiza = ''N'' and lp.divatv = ''N'' and lp.exectd = ''N'') and    ');
  qryBuscaSIAT.SQL.Add('      (TO_CHAR(lp.datven,''YYYY'') = TO_CHAR(sysdate,''YYYY'') )then 1            '); //DIVIDA ANO

  qryBuscaSIAT.SQL.Add('      when (lp.ajuiza = ''N'' and lp.divatv = ''N'' and lp.exectd = ''N'') and    ');
  qryBuscaSIAT.SQL.Add('      (TO_CHAR(lp.datven,''YYYY'') > TO_CHAR(sysdate,''YYYY'') ) then 2           '); //DIVIDA POSTERIOR

  qryBuscaSIAT.SQL.Add('      when (lp.ajuiza = ''N'' and lp.divatv = ''N'' and lp.exectd = ''N'') and    ');
  qryBuscaSIAT.SQL.Add('      (trunc((months_between(sysdate, lp.datven))/12) >= 5 ) then 6               '); //DIVIDA PRESCRITA

//  qryBuscaSIAT.SQL.Add('      when (lp.ajuiza = ''N'' and lp.divatv = ''N'' and lp.exectd = ''N'') and    ');
//  qryBuscaSIAT.SQL.Add('      (TO_CHAR(lp.datven,''YYYY'') < TO_CHAR(sysdate,''YYYY'') ) then 8           '); //DIVIDA ANTERIOR

  qryBuscaSIAT.SQL.Add('     else 8 end as tipoDivida                                                ');

  qryBuscaSIAT.SQL.Add('from siatthe.tblLnc l                                                                                       ');
  qryBuscaSIAT.SQL.Add(' inner join SIATTHE.tbllncpar lp on lp.codlnc = l.codlnc                                                    ');
  qryBuscaSIAT.SQL.Add(' inner join SIATTHE.tblcad cad on l.codcad = cad.codcad                                                     ');
  qryBuscaSIAT.SQL.Add(' inner join SIATTHE.tbltrb trb on trb.codtrb = l.codtrb                                                     ');
  qryBuscaSIAT.SQL.Add(' inner join table(siatthe.PckgTribCalc.CalculaParcelaObj(lp.codlncpar,trunc(current_date))) calc on 1 = 1   ');
  qryBuscaSIAT.SQL.Add('where lp.situac in (''ABER'',''SUSP'',''SUSC'')                                                             ');
  qryBuscaSIAT.SQL.Add(' and lp.vallanmoe > 0                                                                                       ');
  qryBuscaSIAT.SQL.Add(' and cad.TIPCAD in (''I'', ''M'', ''P'') and lp.codlnc >= 20542997                                          ');

//  qryBuscaSIAT.SQL.Add(' and lp.datven between  to_date(''01/11/2013'', ''DD/MM/YYYY'') and to_date(''18/11/2013'', ''DD/MM/YYYY'') ');
//  qryBuscaSIAT.SQL.Add(' group by                                                                                                   ');
//  qryBuscaSIAT.SQL.Add('case when cad.tipcad = ''I'' then ''IMOBILIARIO''                                                           ');
//  qryBuscaSIAT.SQL.Add('     when cad.tipcad = ''M'' then ''MOBILIARIO''                                                            ');
//  qryBuscaSIAT.SQL.Add('     when cad.tipcad = ''P'' then ''PESSOA''                                                                ');
//  qryBuscaSIAT.SQL.Add(' else ''IMOBILIARIO'' end, trb.codtrb, trb.desmin                                                           ');
  qryBuscaSIAT.SQL.Add(' order by lp.codlnc                                                                                        ');

  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;
  vDataHora := now;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    vCodGtr := '';
    qryDados.Close;
    qryDados.SQL.Clear;
    qryDados.SQL.Add('select codgtr from SIATTHE.tbltrbgtr where codtrb =:cod order by datalt desc    ');
    qryDados.Parameters.ParamByName('cod').Value := qryBuscaSIAT.fieldbyname('codtrb').asstring;
    qryDados.open;
    vCodGtr := qryDados.fieldbyname('codgtr').AsString;


    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' CODLNC: '+qryBuscaSIAT.fieldbyname('codlnc').asstring;
    Application.ProcessMessages;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;

    qryRegistro.SQL.Add('INSERT INTO credito_geral(id, tipo_cadastro, tributo_id, datavalidade, datalancamento,        ');
    qryRegistro.SQL.Add(' datapagamento, quantidade, valor_lancado, valor_atualizado, juros,                           ');
    qryRegistro.SQL.Add(' multa, desconto, valor_total, datahora_processamento, codtrb,                                ');
    qryRegistro.SQL.Add(' grupo_tributo_id, codcad, faixa_id, tipo_divida_id, situacao_divida_id,                      ');
    qryRegistro.SQL.Add(' exectd, ajuiza, divatv, situac)                                                              ');

    qryRegistro.SQL.Add('VALUES (nextval(''credito_geral_id_seq''), :tipo_cadastro, :tributo_id, :datavalidade,        ');
    qryRegistro.SQL.Add(' :datalancamento, :datapagamento, :quantidade, :valor_lancado, :valor_atualizado, :juros,     ');
    qryRegistro.SQL.Add(' :multa, :desconto, :valor_total, :datahora_processamento, :codtrb,                           ');
    qryRegistro.SQL.Add(' :grupo, :codcad, :faixa_id, :tipo_divida_id, :situacao_divida,                               ');
    qryRegistro.SQL.Add(' :exectd, :ajuiza, :divatv, :situac)                                                          ');


    qryRegistro.ParamByName('tipo_cadastro').Asstring   := qryBuscaSIAT.fieldbyname('tipoCadastro').Asstring;
    qryRegistro.ParamByName('tributo_id').Value         := TrazID_Tributo(qryBuscaSIAT.fieldbyname('codtrb').AsString);

    if qryBuscaSIAT.fieldbyname('dataValidade').Asstring <> '' then
      qryRegistro.ParamByName('datavalidade').Value     := qryBuscaSIAT.fieldbyname('dataValidade').AsDateTime
    else
      qryRegistro.ParamByName('datavalidade').Value     := null;

    if qryBuscaSIAT.fieldbyname('dataLancamento').AsString <> '' then
      qryRegistro.ParamByName('datalancamento').Value   := qryBuscaSIAT.fieldbyname('dataLancamento').AsDateTime
    else
      qryRegistro.ParamByName('datalancamento').Value   := null;

    if qryBuscaSIAT.fieldbyname('dataPagamento').Asstring <> '' then
      qryRegistro.ParamByName('datapagamento').Value    := qryBuscaSIAT.fieldbyname('dataPagamento').AsDateTime
    else
      qryRegistro.ParamByName('datapagamento').Value    := null;


    qryRegistro.ParamByName('quantidade').AsInteger     := 1;
    qryRegistro.ParamByName('valor_lancado').AsFloat    := qryBuscaSIAT.fieldbyname('valorLancado').AsFloat;
    qryRegistro.ParamByName('valor_atualizado').AsFloat := qryBuscaSIAT.fieldbyname('valorAtualizado').AsFloat;
    qryRegistro.ParamByName('juros').AsFloat            := qryBuscaSIAT.fieldbyname('juros').AsFloat;
    qryRegistro.ParamByName('multa').AsFloat            := qryBuscaSIAT.fieldbyname('multa').AsFloat;
    qryRegistro.ParamByName('desconto').AsFloat         := qryBuscaSIAT.fieldbyname('desconto').AsFloat;
    qryRegistro.ParamByName('valor_total').AsFloat      := qryBuscaSIAT.fieldbyname('valorTotal').AsFloat;
    qryRegistro.ParamByName('datahora_processamento').Value := vDataHora;
    qryRegistro.ParamByName('codtrb').Value             := Trim(qryBuscaSIAT.fieldbyname('codtrb').AsString);

    qryRegistro.ParamByName('grupo').Value              := TrazID_GrupoTributo(vCodGtr);
    qryRegistro.ParamByName('codcad').Value             := qryBuscaSIAT.fieldbyname('codcad').AsInteger;
    qryRegistro.ParamByName('faixa_id').Value           := qryBuscaSIAT.fieldbyname('faixa').AsInteger;

    if Trim(qryBuscaSIAT.fieldbyname('situac').AsString) = 'SUSP' then
      qryRegistro.ParamByName('tipo_divida_id').Value     := 7
    else
      qryRegistro.ParamByName('tipo_divida_id').Value     := qryBuscaSIAT.fieldbyname('tipoDivida').AsInteger;


    if Trim(qryBuscaSIAT.fieldbyname('situac').AsString) = 'SUSP' then
      qryRegistro.ParamByName('situacao_divida').Value     := 3 //Suspensa
    else if qryBuscaSIAT.fieldbyname('tipoDivida').AsInteger = 6 then
      qryRegistro.ParamByName('situacao_divida').Value     := 2  //Prescrita
    else
      qryRegistro.ParamByName('situacao_divida').Value     := 1;//Nomarl

    qryRegistro.ParamByName('exectd').Value             := Trim(qryBuscaSIAT.fieldbyname('exectd').AsString);
    qryRegistro.ParamByName('ajuiza').Value             := Trim(qryBuscaSIAT.fieldbyname('ajuiza').AsString);
    qryRegistro.ParamByName('divatv').Value             := Trim(qryBuscaSIAT.fieldbyname('divatv').AsString);
    qryRegistro.ParamByName('situac').Value             := Trim(qryBuscaSIAT.fieldbyname('situac').AsString);


    Try
      Try
      qryRegistro.ExecSQL;
      Except

      End;
    Finally
    End;

    qryBuscaSIAT.Next;

    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

procedure TfrmImportaICMS.Endereco_Cepisa;
begin
  Status.Panels[0].Text := 'Apagando cpf vazio... ';
  
  qryRegistro.close;
  qryRegistro.sql.Clear;
  qryRegistro.sql.add('delete from pessoa_externa where cpf_cnpj = ''''   ');
  qryRegistro.ExecSQL;

  Linha   := 0;  Entrada := '';  vTipo := ''; vCont := 0;

  vCod_local := ''; vCod_setor := ''; vCod_rota := ''; vCod_sequencia := ''; vUC := '';
  vReferencia := ''; vFD := ''; vTp_motivo := ''; vClasse := ''; vSit_fatura := ''; vConsumo_kwh := '';

  vValor_importe := ''; vValor_cosip := '';

  AssignFile(ArqTexto,edtCaminho.Text);
  Reset(ArqTexto);

  vDataHora := now;


  while not Eoln(ArqTexto) do
    begin
    Linha := Linha + 1;
    Readln(ArqTexto,Entrada);

    Status.Panels[0].Text := 'Registros: ' + IntToStr(Linha);

    If PosEx(',', Entrada) <> 0 then
      Item := LeftStr(Entrada, PosEx(',', Entrada) - 1);

    vCod_local := Item;


    If Pos(',', Entrada) <> 0 then
      Item := Copy(Entrada, Pos(',', Entrada)+1, (Length(Entrada)-Pos(',',Entrada)));
    vCod_setor := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vCod_rota := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vCod_sequencia := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vUC := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vNome := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vCpfCnpj := LeftStr(Item, PosEx(',', Item) - 1);


     //separando tipo do documento
    vGen1 :=  vCpfCnpj;
    If PosEx('|', vCpfCnpj) <> 0 then
      vTipo := LeftStr(vCpfCnpj, PosEx('|', vCpfCnpj) - 1);

    vGen1 := '';
    vGen1 := vTipo;

    if Trim(vGen1) = 'CPF' then
      vTipo := 'PF'
    else
      vTipo := 'PJ';


    If Pos('|', vCpfCnpj) <> 0 then
      vGen2 := Trim(Copy(vCpfCnpj, Pos('|', vCpfCnpj)+1, (Length(vCpfCnpj)-Pos('|',vCpfCnpj))));

    vCpfCnpj := Trim(vGen2);



    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vLogradouro := LeftStr(Item, PosEx(',', Item) - 1);


     //separando logradouro do número
    If PosEx('|', vLogradouro) <> 0 then
      vGen1 := LeftStr(vLogradouro, PosEx('|', vLogradouro) - 1);

    If Pos('|', vLogradouro) <> 0 then
      vGen2 := Trim(Copy(vLogradouro, Pos('|', vLogradouro)+1, (Length(vLogradouro)-Pos('|',vLogradouro))));

    vLogradouro := Trim(vGen1);
    vNumero     := Trim(vGen2);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vComplemento := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vCEP := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vBairro := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vReferencia := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vFD := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vTp_motivo := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vClasse := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vSit_fatura := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vConsumo_kwh := LeftStr(Item, PosEx(',', Item) - 1);


    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vValor_importe := LeftStr(Item, PosEx(',', Item) - 1);

    If Pos(',', Item) <> 0 then
      Item := Copy(Item, Pos(',', Item)+1, (Length(Item)-Pos(',',Item)));
    vValor_cosip := Item;


    vIDPessoa := TrazIDPessoaExterna(vCpfCnpj);

    if vIDPessoa > 0 then
      begin

      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('update pessoa_externa set cepisa =:cepisa  ');
      qryRegistro.sql.add('where id =:id     ');
      qryRegistro.ParamByName('id').Value     := vIDPessoa;
      qryRegistro.ParamByName('cepisa').Value := True;
      qryRegistro.ExecSQL;


      end// if vIDPessoa > 0 then
    else
      begin

      qryRegistro.Close;
      qryRegistro.SQL.Clear;
      qryRegistro.SQL.Add('INSERT INTO pessoa_externa(id, nome, tipo, cpf_cnpj, cepisa, receita, situacao_receita)            ');
      qryRegistro.SQL.Add('VALUES (nextval(''pessoa_externa_id_seq''), :nome, :tipo, :cpf_cnpj, :cepisa, :receita, :situacao) ');

      qryRegistro.ParamByName('nome').Value     := Trim(vNome);
      qryRegistro.ParamByName('tipo').Value     := Trim(vTipo);
      qryRegistro.ParamByName('cpf_cnpj').Value := Trim(vCpfCnpj);
      qryRegistro.ParamByName('cepisa').Value   := True;
      qryRegistro.ParamByName('receita').Value  := False;
      qryRegistro.ParamByName('situacao').Value := '';
      qryRegistro.ExecSQL;

      //pesquisando
      qryTributo.Close;
      qryTributo.SQL.Clear;
      qryTributo.SQL.Add('select Max(id) as Ultimo from pessoa_externa');
      qryTributo.open;

      vIDPessoa := qryTributo.fieldbyname('ultimo').AsInteger;

      end;

    if not VerificaUC_Cepisa(Trim(vUC)) then
      begin
      qryRegistro.Close;
      qryRegistro.SQL.Clear;

      qryRegistro.SQL.Add('INSERT INTO endereco_cepisa(id, uc, cod_local, cod_setor, cod_rota, cod_sequencia,   ');
      qryRegistro.SQL.Add(' logradouro, complemento, bairro, cep, referencia, fd, tp_motivo, classe,            ');
      qryRegistro.SQL.Add(' sit_fatura, consumo_kwh, valor_importe, valor_cosip, pessoa_externa_id,             ');
      qryRegistro.SQL.Add(' datahora_processamento, numero)                                                     ');

      qryRegistro.SQL.Add('VALUES (nextval(''endereco_cepisa_id_seq''), :uc, :cod_local, :cod_setor, :cod_rota,     ');
      qryRegistro.SQL.Add(' :cod_sequencia, :logradouro, :complemento, :bairro, :cep, :referencia, :fd, :tp_motivo, ');
      qryRegistro.SQL.Add(' :classe, :sit_fatura, :consumo_kwh, :valor_importe, :valor_cosip, :pessoa,              ');
      qryRegistro.SQL.Add(' :datahora, :numero)                                                                     ');

      qryRegistro.ParamByName('uc').Value            := Trim(vUC);
      qryRegistro.ParamByName('cod_local').Value     := Trim(vCod_local);
      qryRegistro.ParamByName('cod_setor').Value     := Trim(vCod_setor);
      qryRegistro.ParamByName('cod_rota').Value      := Trim(vCod_rota);
      qryRegistro.ParamByName('cod_sequencia').Value := Trim(vCod_sequencia);
      qryRegistro.ParamByName('logradouro').Value    := Trim(vLogradouro);
      qryRegistro.ParamByName('complemento').Value   := Trim(vComplemento);
      qryRegistro.ParamByName('bairro').Value        := Trim(vBairro);
      qryRegistro.ParamByName('cep').Value           := Trim(vCEP);
      qryRegistro.ParamByName('referencia').Value    := Trim(vReferencia);
      qryRegistro.ParamByName('fd').Value            := Trim(vFD);
      qryRegistro.ParamByName('tp_motivo').Value     := Trim(vTp_motivo);
      qryRegistro.ParamByName('classe').Value        := Trim(vClasse);
      qryRegistro.ParamByName('sit_fatura').Value    := Trim(vSit_fatura);
      qryRegistro.ParamByName('consumo_kwh').Value   := Trim(vConsumo_kwh);
      qryRegistro.ParamByName('valor_importe').Value := vValor_importe;
      qryRegistro.ParamByName('valor_cosip').Value   := vValor_cosip;
      qryRegistro.ParamByName('pessoa').Value        := vIDPessoa;
      qryRegistro.ParamByName('datahora').Value      := vDataHora;
      qryRegistro.ParamByName('numero').Value        := Trim(vNumero);


      Try
        Try
        qryRegistro.ExecSQL;
        Except

        End;
        Finally
      End;

      end;//if not VerificaUC_Cepisa(Trim(vUC)) then

    vCont := vCont + 1;
    Status.Panels[1].Text := IntToStr(vCont)+' -> '+Trim(vCpfCnpj);
    Application.ProcessMessages;

    end;

  CloseFile(ArqTexto);
  ShowMessage('Final da Importação');


end;


procedure TfrmImportaICMS.Endereco_Receita;
begin
  Linha   := 0;  Entrada := '';  vTipo := ''; vCont := 0;

  AssignFile(ArqTexto,edtCaminho.Text);
  Reset(ArqTexto);

  vDataHora := now;


  while not Eoln(ArqTexto) do
    begin
    Linha := Linha + 1;
    Readln(ArqTexto,Entrada);

    Status.Panels[0].Text := 'Registros: ' + IntToStr(Linha);

    If PosEx('~', Entrada) <> 0 then
      Item := LeftStr(Entrada, PosEx('~', Entrada) - 1);

    vCpfCnpj := Item;

    Delete(vCpfCnpj,ansipos('.',vCpfCnpj),1);  //retira as mascaras se houver
    Delete(vCpfCnpj,ansipos('.',vCpfCnpj),1);
    Delete(vCpfCnpj,ansipos('-',vCpfCnpj),1);
    Delete(vCpfCnpj,ansipos('/',vCpfCnpj),1);

    if Length(vCpfCnpj) <= 11 then
      vTipo := 'PF'
    else
      vTipo := 'PJ';

    If Pos('~', Entrada) <> 0 then
      Item := Copy(Entrada, Pos('~', Entrada)+1, (Length(Entrada)-Pos('~',Entrada)));
    vNome := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vTipoLogradouro := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vLogradouro := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vNumero := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vComplemento := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vBairro := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vCEP := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vMunicipio := LeftStr(Item, PosEx('~', Item) - 1);

    If Pos('~', Item) <> 0 then
      Item := Copy(Item, Pos('~', Item)+1, (Length(Item)-Pos('~',Item)));
    vSituacao := Trim(Item);

     //separando uf do município
    vUF :=  vMunicipio;
    If PosEx('-', vMunicipio) <> 0 then
      Item := LeftStr(vMunicipio, PosEx('-', vMunicipio) - 1);

    vMunicipio := Trim(Item);

    If Pos('-', vUF) <> 0 then
      vUF := Trim(Copy(vUF, Pos('-', vUF)+1, (Length(vUF)-Pos('-',vUF))));

    vIDPessoa := TrazIDPessoaExterna(vCpfCnpj);

    if vIDPessoa > 0 then
      begin

      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('update pessoa_externa set receita =:receita  ');
      qryRegistro.sql.add('where id =:id     ');
      qryRegistro.ParamByName('id').Value      := vIDPessoa;
      qryRegistro.ParamByName('receita').Value := True;
      qryRegistro.ExecSQL;


      end// if vIDPessoa > 0 then
    else
      begin

      qryRegistro.Close;
      qryRegistro.SQL.Clear;
      qryRegistro.SQL.Add('INSERT INTO pessoa_externa(id, nome, tipo, cpf_cnpj, cepisa, receita, situacao_receita)            ');
      qryRegistro.SQL.Add('VALUES (nextval(''pessoa_externa_id_seq''), :nome, :tipo, :cpf_cnpj, :cepisa, :receita, :situacao) ');

      qryRegistro.ParamByName('nome').Value     := Trim(vNome);
      qryRegistro.ParamByName('tipo').Value     := Trim(vTipo);
      qryRegistro.ParamByName('cpf_cnpj').Value := Trim(vCpfCnpj);
      qryRegistro.ParamByName('cepisa').Value   := False;
      qryRegistro.ParamByName('receita').Value  := True;
      qryRegistro.ParamByName('situacao').Value := Trim(vSituacao);
      qryRegistro.ExecSQL;

      //pesquisando
      qryTributo.Close;
      qryTributo.SQL.Clear;
      qryTributo.SQL.Add('select Max(id) as Ultimo from pessoa_externa');
      qryTributo.open;

      vIDPessoa := qryTributo.fieldbyname('ultimo').AsInteger;

      end;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO endereco_receita(id, tipo_logradouro, logradouro, numero,                       ');
    qryRegistro.SQL.Add(' complemento, bairro, cep, municipio, uf, datahora_processamento, pessoa_externa_id)        ');

    qryRegistro.SQL.Add('VALUES (nextval(''endereco_receita_id_seq''), :tipo_logradouro, :logradouro, :numero,       ');
    qryRegistro.SQL.Add(' :complemento, :bairro, :cep, :municipio, :uf, :datahora, :pessoa)                          ');

    qryRegistro.ParamByName('tipo_logradouro').Value := Trim(vTipoLogradouro);
    qryRegistro.ParamByName('logradouro').Value      := Trim(vLogradouro);
    qryRegistro.ParamByName('numero').Value          := Trim(vNumero);
    qryRegistro.ParamByName('complemento').Value     := Trim(vComplemento);
    qryRegistro.ParamByName('bairro').Value          := Trim(vBairro);
    qryRegistro.ParamByName('cep').Value             := Trim(vCEP);
    qryRegistro.ParamByName('municipio').Value       := Trim(vMunicipio);
    qryRegistro.ParamByName('uf').Value              := Trim(vUF);
    qryRegistro.ParamByName('datahora').Value        := vDataHora;
    qryRegistro.ParamByName('pessoa').Value          := vIDPessoa;


    Try
      Try
      qryRegistro.ExecSQL;
      Except

      End;
      Finally
    End;


    vCont := vCont + 1;
    Status.Panels[1].Text := IntToStr(vCont);
    Application.ProcessMessages;

    end;

  CloseFile(ArqTexto);
  ShowMessage('Final da Importação');

end;

procedure TfrmImportaICMS.con_sisconBeforeConnect(Sender: TObject);
begin
  con_siscon.Properties.Add('cliente_enconding=latin1');
  con_siscon.Properties.Add('codepage=latin1');

end;


function TfrmImportaICMS.TrazIDPessoaExterna(cpfcnpj: string): integer;
begin
  Result := 0;

  if trim(cpfcnpj) <> '' then
    begin
    qryTributo.Close;
    qryTributo.SQL.Clear;
    qryTributo.SQL.Add('select id from pessoa_externa   ');
    qryTributo.SQL.Add('where cpf_cnpj=:cpf     ');
    qryTributo.ParamByName('cpf').Value := Trim(cpfcnpj);
    qryTributo.open;

    if qryTributo.RecordCount > 0 then
      Result := qryTributo.fieldbyname('id').AsInteger;

    end;

end;

function TfrmImportaICMS.VerificaUC_Cepisa(uc: string): boolean;
begin
  Result := False;

  if trim(uc) <> '' then
    begin
    qryTributo.Close;
    qryTributo.SQL.Clear;
    qryTributo.SQL.Add('select id from endereco_cepisa where uc=:uc   ');
    qryTributo.ParamByName('uc').Value := Trim(uc);
    qryTributo.open;

    if qryTributo.RecordCount > 0 then
      Result := True;

    end;

end;

procedure TfrmImportaICMS.Insere_PessoaSistemaSiat;
var vAtu : integer;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select                                           ');
  qryBuscaSIAT.SQL.Add('eco.codeco,                                      ');
  qryBuscaSIAT.SQL.Add('eco.insmun,                                      ');
  qryBuscaSIAT.SQL.Add('eco.cpfcnpj,                                     ');
  qryBuscaSIAT.SQL.Add('case length(eco.cpfcnpj)                         ');
  qryBuscaSIAT.SQL.Add(' when 14 then siatthe.mascaracnpj(eco.cpfcnpj)   ');
  qryBuscaSIAT.SQL.Add(' when 11 then siatthe.mascaracpf(eco.cpfcnpj)    ');
  qryBuscaSIAT.SQL.Add(' else ''Erro''                                   ');
  qryBuscaSIAT.SQL.Add('end as cpfcnpj_formatado,                        ');
  qryBuscaSIAT.SQL.Add('eco.nomraz as razao_social,                      ');
  qryBuscaSIAT.SQL.Add('eco.nomfan as nome_fantasia,                     ');
  qryBuscaSIAT.SQL.Add('eco.tippes as tipo_pessoa,                       ');
  qryBuscaSIAT.SQL.Add('eco.cnpjmat as cnpj_matriz,                      ');
  qryBuscaSIAT.SQL.Add('case length(eco.cnpjmat)                         ');
  qryBuscaSIAT.SQL.Add('      when 14 then siatthe.mascaracnpj(eco.cpfcnpj) ');
  qryBuscaSIAT.SQL.Add('      when 11 then siatthe.mascaracpf(eco.cpfcnpj)  ');
  qryBuscaSIAT.SQL.Add('      when null then null                           ');
  qryBuscaSIAT.SQL.Add('      else null                                     ');
  qryBuscaSIAT.SQL.Add('end as cnpj_matriz_formatado,                       ');
  qryBuscaSIAT.SQL.Add('eco.codnjr as codigo_natureza_juridica,             ');
  qryBuscaSIAT.SQL.Add('njr.desnjr as natureza_juridica,                    ');
  qryBuscaSIAT.SQL.Add('case siatthe.recuperaSituacaoSiat(eco.codEco,''S'',''TIPSIT'') ');
  qryBuscaSIAT.SQL.Add('    when ''A'' then ''Ativa''                       ');
  qryBuscaSIAT.SQL.Add('    when ''E'' then ''Encerrada''                   ');
  qryBuscaSIAT.SQL.Add('    when ''S'' then ''Suspensa''                    ');
  qryBuscaSIAT.SQL.Add('    else ''Não Mapeada''                            ');
  qryBuscaSIAT.SQL.Add('end as situacao,                                    ');
  qryBuscaSIAT.SQL.Add('case eco.tipins                                     ');
  qryBuscaSIAT.SQL.Add('      when ''OM'' then ''Outros Municipios''        ');
  qryBuscaSIAT.SQL.Add('      when ''N'' then ''Normal''                    ');
  qryBuscaSIAT.SQL.Add('      when ''UA'' then ''Unidade Agregada''         ');
  qryBuscaSIAT.SQL.Add('      when ''OF'' then ''Oficio''                   ');
  qryBuscaSIAT.SQL.Add('      when ''UT'' then ''Unidade Temporaria''       ');
  qryBuscaSIAT.SQL.Add('end as tipo_inscricao,                              ');
  qryBuscaSIAT.SQL.Add('to_char(eco.datcons,''dd/mm/yyyy'') as data_abertura, ');
  qryBuscaSIAT.SQL.Add('case ecoelo.tipimo                                    ');
  qryBuscaSIAT.SQL.Add('    when ''R'' then ''Residencial''                   ');
  qryBuscaSIAT.SQL.Add('    when ''C'' then ''Comercial''                     ');
  qryBuscaSIAT.SQL.Add('    when ''M'' then ''Misto''                         ');
  qryBuscaSIAT.SQL.Add('end as tipo_imovel,                                   ');
  qryBuscaSIAT.SQL.Add('siatthe.RecuperaDsfEnum(ecoelo.tiplog, ''D'', ''RS'') as tipo_logradouro, ');
  qryBuscaSIAT.SQL.Add('ecoelo.nomlog as logradouro,                          ');
  qryBuscaSIAT.SQL.Add('ecoelo.numero as numero,                              ');
  qryBuscaSIAT.SQL.Add('ecoelo.comple as complemento,                         ');
  qryBuscaSIAT.SQL.Add('ecoelo.nombai as bairro,                              ');
  qryBuscaSIAT.SQL.Add('ecoelo.cep as cep,                                    ');
  qryBuscaSIAT.SQL.Add('ecoelo.arefunati as area_func_atividade,              ');
  qryBuscaSIAT.SQL.Add('case eco.tipest                                       ');
  qryBuscaSIAT.SQL.Add('    when ''M'' then ''Sede/Matriz''                   ');
  qryBuscaSIAT.SQL.Add('    when ''F'' then ''Filial''                        ');
  qryBuscaSIAT.SQL.Add('    else ''Não Mapeado''                              ');
  qryBuscaSIAT.SQL.Add('end as tipo_estabelecimento,                          ');
  qryBuscaSIAT.SQL.Add('case                                                  ');
  qryBuscaSIAT.SQL.Add('    when (select count(*) from                        ');
  qryBuscaSIAT.SQL.Add('    (select aaa.* from siatthe.tblecoatv aaa where aaa.situac=''A'' and (aaa.datfim is null or aaa.datfim>sysdate)) ecoatv ');
  qryBuscaSIAT.SQL.Add('left join                      ');
  qryBuscaSIAT.SQL.Add('(select bbb.* from siatthe.tblatv bbb where bbb.situac=''A'' and bbb.ultniv=''S'') atv on atv.codatv=ecoatv.codatv ');
  qryBuscaSIAT.SQL.Add('left join                      ');
  qryBuscaSIAT.SQL.Add('(select max(fff.codatvcfg) as codatvcfg,fff.codatv from siatthe.tblatvcfg fff group by fff.codatv) maxatvcfg on maxatvcfg.codatv=atv.codatv ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblatvcfg atvcfg on atvcfg.codatvcfg=maxatvcfg.codatvcfg  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbllsv lsv on lsv.codlsv=atvcfg.codlsv      ');
  qryBuscaSIAT.SQL.Add('where ecoatv.codeco=eco.codeco and lsv.numlsv is not null)>0  ');
  qryBuscaSIAT.SQL.Add('and                            ');
  qryBuscaSIAT.SQL.Add('    (select count(*) from      ');
  qryBuscaSIAT.SQL.Add('    (select aaa.* from siatthe.tblecoatv aaa where aaa.situac=''A'' and (aaa.datfim is null or aaa.datfim>sysdate)) ecoatv ');
  qryBuscaSIAT.SQL.Add('left join                       ');
  qryBuscaSIAT.SQL.Add('(select bbb.* from siatthe.tblatv bbb where bbb.situac=''A'' and bbb.ultniv=''S'') atv on atv.codatv=ecoatv.codatv  ');
  qryBuscaSIAT.SQL.Add('left join                       ');
  qryBuscaSIAT.SQL.Add('(select max(fff.codatvcfg) as codatvcfg,fff.codatv from siatthe.tblatvcfg fff group by fff.codatv) maxatvcfg on maxatvcfg.codatv=atv.codatv  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblatvcfg atvcfg on atvcfg.codatvcfg=maxatvcfg.codatvcfg  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbllsv lsv on lsv.codlsv=atvcfg.codlsv  ');
  qryBuscaSIAT.SQL.Add('where ecoatv.codeco=eco.codeco and lsv.numlsv is null)=0  ');
  qryBuscaSIAT.SQL.Add('then ''P'' ');

  qryBuscaSIAT.SQL.Add('when (select count(*) from ');
  qryBuscaSIAT.SQL.Add('(select aaa.* from siatthe.tblecoatv aaa where aaa.situac=''A'' and (aaa.datfim is null or aaa.datfim>sysdate)) ecoatv  ');
  qryBuscaSIAT.SQL.Add('left join                 ');
  qryBuscaSIAT.SQL.Add('(select bbb.* from siatthe.tblatv bbb where bbb.situac=''A'' and bbb.ultniv=''S'') atv on atv.codatv=ecoatv.codatv      ');
  qryBuscaSIAT.SQL.Add('left join                 ');
  qryBuscaSIAT.SQL.Add('(select max(fff.codatvcfg) as codatvcfg,fff.codatv from siatthe.tblatvcfg fff group by fff.codatv) maxatvcfg on maxatvcfg.codatv=atv.codatv  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblatvcfg atvcfg on atvcfg.codatvcfg=maxatvcfg.codatvcfg   ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbllsv lsv on lsv.codlsv=atvcfg.codlsv                     ');
  qryBuscaSIAT.SQL.Add('where ecoatv.codeco=eco.codeco and lsv.numlsv is not null)>0                 ');
  qryBuscaSIAT.SQL.Add('and                       ');
  qryBuscaSIAT.SQL.Add('(select count(*) from     ');
  qryBuscaSIAT.SQL.Add('(select aaa.* from siatthe.tblecoatv aaa where aaa.situac=''A'' and (aaa.datfim is null or aaa.datfim>sysdate)) ecoatv                  ');
  qryBuscaSIAT.SQL.Add('left join                 ');
  qryBuscaSIAT.SQL.Add('(select bbb.* from siatthe.tblatv bbb where bbb.situac=''A'' and bbb.ultniv=''S'') atv on atv.codatv=ecoatv.codatv                      ');
  qryBuscaSIAT.SQL.Add('left join                 ');
  qryBuscaSIAT.SQL.Add('(select max(fff.codatvcfg) as codatvcfg,fff.codatv from siatthe.tblatvcfg fff group by fff.codatv) maxatvcfg on maxatvcfg.codatv=atv.codatv ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblatvcfg atvcfg on atvcfg.codatvcfg=maxatvcfg.codatvcfg  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbllsv lsv on lsv.codlsv=atvcfg.codlsv                    ');
  qryBuscaSIAT.SQL.Add('where ecoatv.codeco=eco.codeco and lsv.numlsv is null)>0                    ');
  qryBuscaSIAT.SQL.Add('then ''M''                   ');
  qryBuscaSIAT.SQL.Add('else ''O''                   ');
  qryBuscaSIAT.SQL.Add('end as tipo_prestador,       ');
  qryBuscaSIAT.SQL.Add('case                         ');
  qryBuscaSIAT.SQL.Add('when (select count(*) from   ');
  qryBuscaSIAT.SQL.Add('(select ccc.* from siatthe.tblecoeqditm ccc where (ccc.datfim is null or ccc.datfim>sysdate) and ccc.situac=''A'') ecoeqditm ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbleqditm eqditm on eqditm.codeqditm=ecoeqditm.codeqditm  ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tbleqd eqd on eqd.codeqd=eqditm.codeqd              ');
  qryBuscaSIAT.SQL.Add('where ecoeqditm.codeco=eco.codeco and eqditm.codeqditm in (5,7,21))>0 ');
  qryBuscaSIAT.SQL.Add('then ''S''                                                            ');
  qryBuscaSIAT.SQL.Add('else ''N''                                                            ');
  qryBuscaSIAT.SQL.Add('end as simples_nacional                                               ');
  qryBuscaSIAT.SQL.Add('from siatthe.tbleco eco                                               ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblnjr njr on njr.codnjr=eco.codnjr                 ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblecoelo ecoelo on ecoelo.codeco=eco.codeco        ');
  qryBuscaSIAT.SQL.Add('where siatthe.recuperaSituacaoSiat(eco.codEco,''S'',''TIPSIT'')=''A'' ');
//  qryBuscaSIAT.SQL.Add(' where eco.cpfcnpj in (''93419083000139'',''81223445667749'')      ');
  qryBuscaSIAT.open;



  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;



  vCont := 0; vAtu := 0;

  while not qryBuscaSIAT.eof do
    begin
    vIDPessoa := 0;

    vIDPessoa := TrazIDPessoa_Sist_Siat(qryBuscaSIAT.fieldbyname('cpfcnpj').AsString);

    if vIDPessoa = 0 then
      begin
      vCont := vCont + 1;

      Status.Panels[1].Text := 'Inserido: ' + IntToStr(vCont)+' Cod.Eco: '+qryBuscaSIAT.fieldbyname('codeco').AsString;
      Application.ProcessMessages;


      qryDestino.Close;
      qryDestino.SQL.Clear;
      qryDestino.SQL.Add('INSERT INTO pessoa_sistema_siat(id, codeco, tippes,             ');
      qryDestino.SQL.Add(' insmun, cpf_cnpj, nomeraz, nomerazres, tipo_prestador,tipo_imovel,simples_nacional)         ');
      qryDestino.SQL.Add('VALUES (nextval(''pessoa_sistema_siat_id_seq''), :codeco, :tippes,     ');
      qryDestino.SQL.Add('  :insmun, :cpf_cnpj, :nomeraz, :nomerazres,:tipo_prestador,:tipo_imovel,:simples_nacional ) ');

      qryDestino.ParamByName('codeco').Value     := qryBuscaSIAT.fieldbyname('codeco').AsInteger;
      qryDestino.ParamByName('tippes').Value     := Trim(qryBuscaSIAT.fieldbyname('tipo_pessoa').AsString);
      qryDestino.ParamByName('insmun').Value     := Trim(qryBuscaSIAT.fieldbyname('insmun').AsString);
      qryDestino.ParamByName('cpf_cnpj').Value   := Trim(qryBuscaSIAT.fieldbyname('cpfcnpj').AsString);
      qryDestino.ParamByName('nomeraz').Value    := Trim(qryBuscaSIAT.fieldbyname('razao_social').AsString);
      qryDestino.ParamByName('nomerazres').Value := Trim(qryBuscaSIAT.fieldbyname('razao_social').AsString);
      qryDestino.ParamByName('tipo_prestador').Value   := Trim(qryBuscaSIAT.fieldbyname('tipo_prestador').AsString);
      qryDestino.ParamByName('tipo_imovel').Value      := Trim(qryBuscaSIAT.fieldbyname('tipo_imovel').AsString);
      qryDestino.ParamByName('simples_nacional').Value := Trim(qryBuscaSIAT.fieldbyname('simples_nacional').AsString);
      qryDestino.ExecSQL;
      end
    else//Caso encontre a pessoa, atualizar os campos principais. //02/02/2016
      begin
      vAtu := vAtu + 1;

      Status.Panels[1].Text := 'Atualizado: ' + IntToStr(vAtu)+' Cod.Eco: '+qryBuscaSIAT.fieldbyname('codeco').AsString;
      Application.ProcessMessages;

      qryImportacao.close;
      qryImportacao.sql.Clear;
      qryImportacao.sql.add('UPDATE pessoa_sistema_siat                                                                 ');
      qryImportacao.sql.add(' SET codeco=:codeco, tippes=:tippes, insmun=:insmun, nomeraz=:nomeraz, ');
      qryImportacao.sql.add(' nomerazres=:nomerazres, tipo_prestador=:tipo_prestador, tipo_imovel=:tipo_imovel,         ');
      qryImportacao.sql.add(' simples_nacional=:simples_nacional    ');
      qryImportacao.sql.add('where id =:id                          ');
      qryImportacao.ParamByName('id').Value         := vIDPessoa;
      qryImportacao.ParamByName('codeco').Value     := qryBuscaSIAT.fieldbyname('codeco').AsInteger;
      qryImportacao.ParamByName('tippes').Value     := Trim(qryBuscaSIAT.fieldbyname('tipo_pessoa').AsString);
      qryImportacao.ParamByName('insmun').Value     := Trim(qryBuscaSIAT.fieldbyname('insmun').AsString);
      qryImportacao.ParamByName('nomeraz').Value    := Trim(qryBuscaSIAT.fieldbyname('razao_social').AsString);
      qryImportacao.ParamByName('nomerazres').Value := Trim(qryBuscaSIAT.fieldbyname('razao_social').AsString);
      qryImportacao.ParamByName('tipo_prestador').Value   := Trim(qryBuscaSIAT.fieldbyname('tipo_prestador').AsString);
      qryImportacao.ParamByName('tipo_imovel').Value      := Trim(qryBuscaSIAT.fieldbyname('tipo_imovel').AsString);
      qryImportacao.ParamByName('simples_nacional').Value := Trim(qryBuscaSIAT.fieldbyname('simples_nacional').AsString);
      qryImportacao.ExecSQL;


      end;


    qryBuscaSIAT.next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

procedure TfrmImportaICMS.Insere_TabelaEcoAtv;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select codecoatv, codeco, codatv, fiscal, tipatv  ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.tblecoatv order by codeco            ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' Cod.Eco: '+qryBuscaSIAT.fieldbyname('codeco').AsString;
    Application.ProcessMessages;


    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO tabela_ecoatv_siat(id, codecoatv, codeco, codatv, fiscal, tipatv) ');
    qryRegistro.SQL.Add('VALUES (nextval(''tabela_ecoatv_siat_id_seq''), :codecoatv, :codeco, :codatv, :fiscal, :tipatv) ');
    qryRegistro.ParamByName('codecoatv').Value := qryBuscaSIAT.fieldbyname('codecoatv').AsInteger;
    qryRegistro.ParamByName('codeco').Value    := qryBuscaSIAT.fieldbyname('codeco').AsInteger;
    qryRegistro.ParamByName('codatv').Value    := Trim(qryBuscaSIAT.fieldbyname('codatv').AsString);
    qryRegistro.ParamByName('fiscal').Value    := Trim(qryBuscaSIAT.fieldbyname('fiscal').AsString);
    qryRegistro.ParamByName('tipatv').Value    := Trim(qryBuscaSIAT.fieldbyname('tipatv').AsString);
    qryRegistro.ExecSQL;

    qryBuscaSIAT.next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

procedure TfrmImportaICMS.Insere_TabelaAtv;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select codatv, codatvsup, numatv, titulo, descom  ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.tblatv order by codatv               ');
  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' Cod.Atv: '+qryBuscaSIAT.fieldbyname('codatv').AsString;
    Application.ProcessMessages;


    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO tabela_atv_siat(id, codatv, codatvsup, numatv, titulo, descom)              ');
    qryRegistro.SQL.Add('VALUES (nextval(''tabela_atv_siat_id_seq''), :codatv, :codatvsup, :numatv, :titulo, :descom) ');

    qryRegistro.ParamByName('codatv').Value    := Trim(qryBuscaSIAT.fieldbyname('codatv').AsString);
    qryRegistro.ParamByName('codatvsup').Value := Trim(qryBuscaSIAT.fieldbyname('codatvsup').AsString);
    qryRegistro.ParamByName('numatv').Value    := Trim(qryBuscaSIAT.fieldbyname('numatv').AsString);
    qryRegistro.ParamByName('titulo').Value    := Trim(qryBuscaSIAT.fieldbyname('titulo').AsString);
    qryRegistro.ParamByName('descom').Value    := Trim(qryBuscaSIAT.fieldbyname('descom').AsString);
    qryRegistro.ExecSQL;

    qryBuscaSIAT.next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

procedure TfrmImportaICMS.Insere_ISSPago;
begin
  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select cad.codpes, x.valorTotal as valortotal,                                 ');
  qryBuscaSIAT.SQL.Add('x.ValAtu as valorAtualizado, x.vlrLancado, x.codlnc, x.codtrb, x.codcad        ');

  qryBuscaSIAT.SQL.Add('	from (select cad.codcad codcad, cad.tipcad tipcad,t.codtrb,dpp.codlnc,                    ');
  qryBuscaSIAT.SQL.Add('	t.desmin as descricao, m.DATMVA as dtMov,  count(*) QtdParcelas,                          ');
  qryBuscaSIAT.SQL.Add('	dpp.identi as identi, sum(dpp.vallanmoe) vlrLancado,                                      ');
  qryBuscaSIAT.SQL.Add('	sum(dpp.vallanmoe) + sum(case when da.codtrb is not null then d.valdoc else 0 end) ValLan,');
  qryBuscaSIAT.SQL.Add('	sum(dpp.vallanmoe + dpp.atumon + dpp.jurfin) ValAtu,                                      ');
  qryBuscaSIAT.SQL.Add('	sum(dpp.jurmor) Juros, sum(dpp.mulmor) Multa, sum(dpp.descon) Desconto,                   ');
  qryBuscaSIAT.SQL.Add('	sum(dp.valemodca) Emolumento, sum(dpp.valpago) valorTotal                                 ');
  qryBuscaSIAT.SQL.Add('	from SIATTHE.TBLMVA m                                                                     ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLMVALTA ml                                                           ');
  qryBuscaSIAT.SQL.Add('	    on m.codmva = ml.codmva                                                               ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLDCM d                                                               ');
  qryBuscaSIAT.SQL.Add('	    on ml.codmvalta = d.codmvalta                                                         ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMPAG dp                                                            ');
  qryBuscaSIAT.SQL.Add('	on d.coddcm = dp.coddcm                                                                   ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMPAGPAR dpp                                                        ');
  qryBuscaSIAT.SQL.Add('	    on dp.coddcmpag = dpp.coddcmpag                                                       ');
  qryBuscaSIAT.SQL.Add('	left join siatthe.tblcad cad                                                              ');
  qryBuscaSIAT.SQL.Add('	    on dpp.codcad = cad.codcad                                                            ');
  qryBuscaSIAT.SQL.Add('	left join SIATTHE.TBLDCMAJS da                                                            ');
  qryBuscaSIAT.SQL.Add('	    on (d.coddcm = da.coddcm)                                                             ');
  qryBuscaSIAT.SQL.Add('	inner join SIATTHE.TBLTRB t                                                               ');
  qryBuscaSIAT.SQL.Add('	    on (t.codtrb = dpp.codtrb or t.codtrb = da.codtrb)                                    ');

  qryBuscaSIAT.SQL.Add('      inner join SIATTHE.tbltrbgtr tg                        ');
  qryBuscaSIAT.SQL.Add('    on t.codtrb = tg.codtrb                                  ');
  qryBuscaSIAT.SQL.Add('  inner join SIATTHE.tblgtr g                                ');
  qryBuscaSIAT.SQL.Add('    on g.codgtr = tg.codgtr where g.codgtr = 36              ');


  qryBuscaSIAT.SQL.Add('	group by t.codtrb, t.desmin,  cad.tipcad, cad.codcad, m.DATMVA, dpp.identi, dpp.codlnc        ');
  qryBuscaSIAT.SQL.Add('	) x left join siatthe.tblcad cad on cad.codcad = x.codcad                                     ');


  qryBuscaSIAT.SQL.Add(' order by cad.codpes                                ');

  qryBuscaSIAT.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;
  vDataHora := now;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    vIDPessoa := TrazIDPessoa_Sistema_Siat(qryBuscaSIAT.fieldbyname('codpes').AsInteger);

    if vIDPessoa > 0 then
      begin
      qryRegistro.Close;
      qryRegistro.SQL.Clear;

      qryRegistro.SQL.Add('INSERT INTO iss_pago(id, tributo_id, data_vencimento, data_pagamento, valor_lancado,  ');
      qryRegistro.SQL.Add(' valor_atualizado, valor_total, datahora_processamento, codtrb,                       ');
      qryRegistro.SQL.Add(' codcad, pessoa_sistema_siat_id, data_lancamento)                                     ');

      qryRegistro.SQL.Add('VALUES (nextval(''iss_pago_id_seq''), :tributo_id, :data_vencimento, :data_pagamento, :valor_lancado,  ');
      qryRegistro.SQL.Add(' :valor_atualizado, :valor_total, :datahora_processamento, :codtrb,                       ');
      qryRegistro.SQL.Add(' :codcad, :pessoa, :data_lancamento)                                                      ');

      qryRegistro.ParamByName('tributo_id').Value    := TrazID_Tributo(qryBuscaSIAT.fieldbyname('codtrb').AsString);

      TrazDatas; //Traz data lançamento pelo codlnc na tabela siatthe.tbllnc

      if qryDados.fieldbyname('datven').Asstring <> '' then
        qryRegistro.ParamByName('data_vencimento').Value     := qryDados.fieldbyname('datven').AsDateTime
      else
        qryRegistro.ParamByName('data_vencimento').Value     := null;

      if qryDados.fieldbyname('datpag').Asstring <> '' then
        qryRegistro.ParamByName('data_pagamento').Value    := qryDados.fieldbyname('datpag').AsDateTime
      else
        qryRegistro.ParamByName('data_pagamento').Value    := null;

      qryRegistro.ParamByName('valor_lancado').AsFloat        := qryBuscaSIAT.fieldbyname('vlrLancado').AsFloat;
      qryRegistro.ParamByName('valor_atualizado').AsFloat     := qryBuscaSIAT.fieldbyname('valorAtualizado').AsFloat;
      qryRegistro.ParamByName('valor_total').AsFloat          := qryBuscaSIAT.fieldbyname('valortotal').AsFloat;
      qryRegistro.ParamByName('datahora_processamento').Value := vDataHora;
      qryRegistro.ParamByName('codtrb').Value                 := Trim(qryBuscaSIAT.fieldbyname('codtrb').AsString);
      qryRegistro.ParamByName('codcad').Value                 := Trim(qryBuscaSIAT.fieldbyname('codcad').AsString);
      qryRegistro.ParamByName('pessoa').Value                 := vIDPessoa;

      if qryDados.fieldbyname('datbla').AsString <> '' then
        qryRegistro.ParamByName('data_lancamento').Value   := qryDados.fieldbyname('datbla').AsDateTime
      else
        qryRegistro.ParamByName('data_lancamento').Value   := null;



      Try
        Try
        qryRegistro.ExecSQL;
        Except

        End;
      Finally
      End;

      end;//if TrazIDPessoa_Sistema_Siat() > 0 then

    qryBuscaSIAT.Next;

    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

function TfrmImportaICMS.TrazIDPessoa_Sistema_Siat(
  codpes: integer): integer;
begin
  Result := 0;

  qryTributo.Close;
  qryTributo.SQL.Clear;
  qryTributo.SQL.Add('select id from pessoa_sistema_siat   ');
  qryTributo.SQL.Add('where codpes=:cod     ');
  qryTributo.ParamByName('cod').Value := codpes;
  qryTributo.open;

  if qryTributo.RecordCount > 0 then
    Result := qryTributo.fieldbyname('id').AsInteger;



end;

procedure TfrmImportaICMS.Atualiza_IDPessoaNF;

begin
  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select id, pessoa_sistema_siat_id, cnpj_cpf  ');
  qryDestino.SQL.Add('from nota_fiscal where id > 8000000 and id <= 16000000 order by id ');
  qryDestino.open;

  vCont := 0;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryDestino.recordcount);
  Application.ProcessMessages;

  while not qryDestino.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := IntToStr(vCont)+' - NF.ID.: '+IntToStr(qryDestino.fieldbyname('id').Value);
    Application.ProcessMessages;

    qryImportacao.close;
    qryImportacao.sql.Clear;
    qryImportacao.sql.add('update nota_fiscal set pessoa_sistema_siat_id =:pid  ');
    qryImportacao.sql.add('where id =:id     ');
    qryImportacao.ParamByName('id').Value   := qryDestino.fieldbyname('id').Asinteger;
    qryImportacao.ParamByName('pid').Value  := TrazIDPessoa_Sist_Siat(qryDestino.fieldbyname('cnpj_cpf').AsString);
    qryImportacao.ExecSQL;


    qryDestino.Next;
    end;
//=========== ATUALIZAR TABELA registro65

  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select id, pessoa_sistema_siat_id, cnpj_mf from registro65 where id > 8000000 and id <= 16000000 order by id ');
  qryDestino.open;

  vCont := 0;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryDestino.recordcount);
  Application.ProcessMessages;

  while not qryDestino.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := IntToStr(vCont)+' - r65.ID.: '+IntToStr(qryDestino.fieldbyname('id').Value);
    Application.ProcessMessages;

    qryImportacao.close;
    qryImportacao.sql.Clear;
    qryImportacao.sql.add('update registro65 set pessoa_sistema_siat_id =:pid  ');
    qryImportacao.sql.add('where id =:id     ');
    qryImportacao.ParamByName('id').Value   := qryDestino.fieldbyname('id').Asinteger;
    qryImportacao.ParamByName('pid').Value  := TrazIDPessoa_Sist_Siat(qryDestino.fieldbyname('cnpj_mf').AsString);
    qryImportacao.ExecSQL;


    qryDestino.Next;
    end;



  showmessage('Acabou...');

end;

function TfrmImportaICMS.TrazIDPessoa_Sist_Siat(cpfcnpj: string): integer;
begin
  Result := 0;

  qryVerifica.Close;
  qryVerifica.SQL.Clear;
  qryVerifica.SQL.Add('select id from pessoa_sistema_siat   ');
  qryVerifica.SQL.Add('where cpf_cnpj=:cpf     ');
  qryVerifica.ParamByName('cpf').Value := Trim(cpfcnpj);
  qryVerifica.open;

  if qryVerifica.RecordCount > 0 then
    Result := qryVerifica.fieldbyname('id').AsInteger;

end;

procedure TfrmImportaICMS.BitBtn5Click(Sender: TObject);
begin
     // 2016
  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select nome, cnpj from pessoa ');
  qryDestino.open;

  vCont := 0;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryDestino.recordcount);
  Application.ProcessMessages;

  while not qryDestino.eof do
    begin

    qryVerifica.Close;
    qryVerifica.SQL.Clear;
    qryVerifica.SQL.Add('select id from pessoa_sistema_siat where cpf_cnpj =:cnpj ');
    qryVerifica.ParamByName('cnpj').Value   := qryDestino.fieldbyname('cnpj').Value;
    qryVerifica.open;

    if qryVerifica.RecordCount = 0 then
      begin

      vCont := vCont + 1;

      Status.Panels[1].Text := IntToStr(vCont)+' - ID.: '+qryDestino.fieldbyname('cnpj').Asstring;
      Application.ProcessMessages;

      //95496 -> último id de pessoa_sistema_siat
      qryImportacao.close;
      qryImportacao.sql.Clear;
      qryImportacao.SQL.Add('INSERT INTO pessoa_sistema_siat(id, cpf_cnpj, nomeraz, nomerazres)                 ');
      qryImportacao.SQL.Add('VALUES (nextval(''pessoa_sistema_siat_id_seq''), :cpf_cnpj, :nomeraz, :nomerazres) ');
      qryImportacao.ParamByName('cpf_cnpj').Value   := qryDestino.fieldbyname('cnpj').Value;
      qryImportacao.ParamByName('nomeraz').Value    := qryDestino.fieldbyname('nome').AsString;
      qryImportacao.ParamByName('nomerazres').Value := qryDestino.fieldbyname('nome').AsString;
      qryImportacao.ExecSQL;
      end;

    qryDestino.Next;
    end;

  showmessage('Acabou...');



//  btnConfirma.Enabled := False;
//  RetiraNotaFiscalDuplicada;
//  Insere_Pessoa_GSRF;
//  Atualiza_IDPessoaNF;
end;

procedure TfrmImportaICMS.Agrupa_ISS_Pago;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('SELECT pes.id, pes.cpf_cnpj,extract(YEAR from data_vencimento) as ano, extract(MONTH from data_vencimento) as mes,  ');
  qryImportacao.SQL.Add('sum(issp.valor_lancado) as valor_lancado, sum(issp.valor_atualizado) as valor_atualizado,                           ');
  qryImportacao.SQL.Add('  sum(issp.valor_total) as valor_total                                                                              ');
  qryImportacao.SQL.Add('FROM iss_pago issp, pessoa_sistema_siat pes                                                                         ');
  qryImportacao.SQL.Add('WHERE issp.pessoa_sistema_siat_id = pes.id                                                                          ');
  //qryImportacao.SQL.Add('  and data_vencimento is not null                                                                                   ');
  qryImportacao.SQL.Add('  and tippes = ''PJ''                                                                                               ');
  qryImportacao.SQL.Add('group by pes.id,pes.cpf_cnpj, ano, mes                                                                              ');
  qryImportacao.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryImportacao.recordcount);
  Application.ProcessMessages;
  vDataHora := now;
  while not qryImportacao.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;


    qryDestino.Close;
    qryDestino.SQL.Clear;

    qryDestino.SQL.Add('INSERT INTO agrupa_iss_pago(id, pessoa_sistema_siat_id, cnpj, valor_lancado,  ');
    qryDestino.SQL.Add(' valor_atualizado, valor_total, datahora_processamento,ano, mes, data)       ');
    qryDestino.SQL.Add('VALUES (nextval(''agrupa_iss_pago_id_seq''), :pessoa_sistema_siat_id, :cnpj, :valor_lancado, ');
    qryDestino.SQL.Add(' :valor_atualizado, :valor_total, :datahora_processamento,:ano, :mes, :data)   ');

    qryDestino.ParamByName('pessoa_sistema_siat_id').Value := qryImportacao.fieldbyname('id').Value;
    qryDestino.ParamByName('cnpj').Value                   := Trim(qryImportacao.fieldbyname('cpf_cnpj').Value);
    qryDestino.ParamByName('valor_lancado').Value    := qryImportacao.fieldbyname('valor_lancado').Value;
    qryDestino.ParamByName('valor_atualizado').Value := qryImportacao.fieldbyname('valor_atualizado').Value;
    qryDestino.ParamByName('valor_total').Value      := qryImportacao.fieldbyname('valor_total').Value;

    qryDestino.ParamByName('datahora_processamento').Value := vDataHora;
    qryDestino.ParamByName('ano').Value              := qryImportacao.fieldbyname('ano').Value;
    qryDestino.ParamByName('mes').Value              := qryImportacao.fieldbyname('mes').Value;
    qryDestino.ParamByName('data').Value             := qryImportacao.fieldbyname('ano').Asstring+'-'+
                                                        qryImportacao.fieldbyname('mes').AsString+'-01';

    Try
      Try
      qryDestino.ExecSQL;
      Except
      vArquivoTexto := 'C:\SEMF\AGRUPA_ISSPAGO.TXT';
      GravaArquivoTexto(qryImportacao.fieldbyname('cpf_cnpj').Value);

      End;
    Finally
    End;

    qryImportacao.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);


end;

procedure TfrmImportaICMS.con_gsrfBeforeConnect(Sender: TObject);
begin
  con_gsrf.Properties.Add('cliente_enconding=latin1');
  con_gsrf.Properties.Add('codepage=latin1');

end;

procedure TfrmImportaICMS.Insere_Pessoa_GSRF;
begin
  qryGSRF.Close;
  qryGSRF.SQL.Clear;
  qryGSRF.SQL.Add('select *  ');
  qryGSRF.SQL.Add('from carlito              ');
  qryGSRF.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryGSRF.recordcount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryGSRF.eof do
    begin

    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    //=====Endereço==========
    vGen1 := UpperCase(Trim(qryGSRF.fieldbyname('cidade').AsString));
    vGen1 := TrocaCaracterEspecial(vGen1,true);

    qryGSRF2.Close;
    qryGSRF2.SQL.Clear;
    qryGSRF2.SQL.Add('INSERT INTO endereco(id, bairro, cep, complemento, logradouro, numero, ');
    qryGSRF2.SQL.Add(' pontoreferencia, municipio_id)                                        ');
    qryGSRF2.SQL.Add('VALUES(nextval(''endereco_id_seq''), :bairro, :cep, :complemento, :logradouro, :numero, ');
    qryGSRF2.SQL.Add(' :ponto, :municipio_id)                                      ');
    qryGSRF2.ParamByName('bairro').Value       := '';
    qryGSRF2.ParamByName('cep').Value          := Trim(qryGSRF.fieldbyname('cep').AsString);
    qryGSRF2.ParamByName('complemento').Value  := '';
    qryGSRF2.ParamByName('logradouro').Value   := UpperCase(Trim(qryGSRF.fieldbyname('endereco').AsString));
    qryGSRF2.ParamByName('numero').Value       := Trim(qryGSRF.fieldbyname('numero').AsString);
    qryGSRF2.ParamByName('ponto').Value        := '';
    qryGSRF2.ParamByName('municipio_id').Value := TrazID_Municipio(vGen1);
    qryGSRF2.ExecSQL;

    qryGSRF2.Close;
    qryGSRF2.SQL.Clear;
    qryGSRF2.SQL.Add('select Max(id) as ultimo from endereco');
    qryGSRF2.open;

    vID := qryGSRF2.fieldbyname('ultimo').AsInteger;

    //===== Endereço Final ==========

    qryGSRF2.Close;
    qryGSRF2.SQL.Clear;

    qryGSRF2.SQL.Add('INSERT INTO pessoa(tipo, id, email, nome, telefonecomercial,   ');
    qryGSRF2.SQL.Add(' cpf, cnpj, endereco_id)                                ');
    qryGSRF2.SQL.Add('VALUES (:tipo, nextval(''pessoa_id_seq''), :email, :nome, :telefonecomercial,  ');
    qryGSRF2.SQL.Add(' :cpf, :cnpj, :endereco_id)                         ');

    if Length(Trim(qryGSRF.fieldbyname('cnpj_cpf').AsString)) = 11 then
      qryGSRF2.ParamByName('tipo').Value    := 'PF'
    else
      qryGSRF2.ParamByName('tipo').Value    := 'PJ';

    qryGSRF2.ParamByName('email').Value    := Trim(qryGSRF.fieldbyname('email').AsString);

    qryGSRF2.ParamByName('nome').Value     := UpperCase(Trim(qryGSRF.fieldbyname('razao_social').AsString));
    qryGSRF2.ParamByName('telefonecomercial').Value    := Trim(qryGSRF.fieldbyname('telefone').AsString);

    if Length(Trim(qryGSRF.fieldbyname('cnpj_cpf').AsString)) = 11 then
      qryGSRF2.ParamByName('cpf').Value    := Trim(qryGSRF.fieldbyname('cnpj_cpf').AsString)
    else
      qryGSRF2.ParamByName('cnpj').Value    := Trim(qryGSRF.fieldbyname('cnpj_cpf').AsString);

    qryGSRF2.ParamByName('endereco_id').Value := vID;
    qryGSRF2.ExecSQL;

    qryGSRF.next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);

end;

function TfrmImportaICMS.TrazID_Municipio(cidade: string): integer;
begin
  Result := 0;

  qryGSRF3.Close;
  qryGSRF3.SQL.Clear;
  qryGSRF3.SQL.Add('select id from municipio   ');
  qryGSRF3.SQL.Add('where descricao=:descr     ');
  qryGSRF3.ParamByName('descr').Value := Trim(cidade);
  qryGSRF3.open;

  if qryGSRF3.RecordCount > 0 then
    Result := qryGSRF3.fieldbyname('id').AsInteger;

end;

procedure TfrmImportaICMS.RetiraNotaFiscalDuplicada;
var
 vNota : string;
begin
  lblInicio.Caption := 'INÍCIO: '+timetostr(now)+' -> '+datetostr(date);
  BitBtn5.Enabled   := False;

  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select id,numero_notafiscal, cnpj_cpf                             ');
  qryDestino.SQL.Add('from nota_fiscal               ');
  qryDestino.SQL.Add('where situacao = ''N''              ');
//  qryDestino.SQL.Add('  and cnpj_cpf = ''06259075000178''              ');
  qryDestino.SQL.Add('  and extract(YEAR from data_emissao) = 2013    ');
//  qryDestino.SQL.Add('  and extract(MONTH from data_emissao) = 11                         ');
  qryDestino.SQL.Add('  and numero_notafiscal in                                          ');

  qryDestino.SQL.Add('(select numero_notafiscal                                         ');
  qryDestino.SQL.Add('from nota_fiscal                                                  ');
  qryDestino.SQL.Add('where situacao = ''N''                               ');
//  qryDestino.SQL.Add('  and cnpj_cpf = ''06259075000178''                               ');
  qryDestino.SQL.Add('  and extract(YEAR from data_emissao) = 2013  ');
//  qryDestino.SQL.Add('  and extract(MONTH from data_emissao) = 11                       ');
  qryDestino.SQL.Add('group by numero_notafiscal                                        ');
  qryDestino.SQL.Add('having count(*) > 1)                                              ');
  qryDestino.SQL.Add('order by cnpj_cpf, numero_notafiscal                              ');

  qryDestino.open;

  vNota := '';
  vCont := 0;
  Status.Panels[0].Text := 'Registros: '+IntToStr(qryDestino.RecordCount);
  Application.ProcessMessages;

  while not qryDestino.eof do
    begin

      vCont := vCont + 1;

      Status.Panels[1].Text := IntToStr(vCont);
      Application.ProcessMessages;


    if Trim(vNota) = Trim(qryDestino.fieldbyname('numero_notafiscal').AsString) then
      begin

      qryImportacao.close;
      qryImportacao.sql.Clear;
      qryImportacao.sql.add('update nota_fiscal set situacao =:sit  ');
      qryImportacao.sql.add('where id =:id     ');
      qryImportacao.ParamByName('id').Value  := qryDestino.fieldbyname('id').Asinteger;
      qryImportacao.ParamByName('sit').Value := 'X';
      qryImportacao.ExecSQL;

      end;


     vNota := Trim(qryDestino.fieldbyname('numero_notafiscal').AsString);

    qryDestino.next;
    end;
   lblFim.Caption := 'FINAL: '+timetostr(now)+' -> '+datetostr(date);

   showmessage('Final da marcação dos duplicados...');

end;

procedure TfrmImportaICMS.Processa_ArrecadacaoGrupoLocal;
begin

  qryBuscaSIAT.Close;
  qryBuscaSIAT.SQL.Clear;
  qryBuscaSIAT.SQL.Add('select m.datmva, ur.coduor,  upper(ur.codorg) "local", dca.codusualt as "login",u.codusu, ');
  qryBuscaSIAT.SQL.Add('u.nomcom as "usuario",g.codgtr,g.codgrp as "codigoGrupo",                                 ');
  qryBuscaSIAT.SQL.Add('upper(g.descri) as "descricaoGrupo",t.codtrb,t.desres,t.desmin,                           ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe "valLan",                                                                   ');
  qryBuscaSIAT.SQL.Add('dpp.vallanmoe + dpp.atumon + dpp.jurfin "valAtu",                                         ');
  qryBuscaSIAT.SQL.Add('dpp.jurmor "juros", dpp.mulmor "multa",                                                   ');
  qryBuscaSIAT.SQL.Add('dpp.descon "desconto",  dp.valemodca "emolumento", dpp.valpago "valorTotal"               ');
  qryBuscaSIAT.SQL.Add('from SIATTHE.TBLMVA m inner join SIATTHE.TBLMVALTA ml on m.codmva = ml.codmva             ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLDCM d on ml.codmvalta = d.codmvalta                                 ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAG dp on d.coddcm = dp.coddcm                                    ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCA dca on dp.coddca = dca.coddca                                    ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMPAGPAR dpp on dp.coddcmpag = dpp.coddcmpag                        ');
  qryBuscaSIAT.SQL.Add('left join siatthe.tblcad cad on dpp.codcad = cad.codcad                                   ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLDCMAJS da on (d.coddcm = da.coddcm)                                  ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.TBLTRB t on (t.codtrb = dpp.codtrb or t.codtrb = da.codtrb)            ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.tbltrbgtr tg  on t.codtrb = tg.codtrb                                  ');
  qryBuscaSIAT.SQL.Add('inner join SIATTHE.tblgtr g on g.codgtr = tg.codgtr                                       ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLUSU u on dca.codusualt = u.login                                     ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLFUN fun on u.codusu = fun.codusu                                     ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLLTC ltc on fun.codfun = ltc.codfun                                   ');
  qryBuscaSIAT.SQL.Add('left join SIATTHE.TBLUOR ur on ltc.coduor = ur.coduor                                     ');
  qryBuscaSIAT.SQL.Add('where g.codgtr in (34,35,36,37,38,39,40)                                                  ');
  qryBuscaSIAT.SQL.Add('  and m.datmva between  to_date(''01/01/2014'', ''DD/MM/YYYY'') and to_date(''31/01/2014'', ''DD/MM/YYYY'')  ');
  qryBuscaSIAT.SQL.Add('  order by 1,3,4,5                                                                        ');
  qryBuscaSIAT.open;
  Status.Panels[0].Text := 'Total: ' + IntToStr(qryBuscaSIAT.recordcount);
  Application.ProcessMessages;

  vCont := 0;
  vDataHora := now;

  while not qryBuscaSIAT.eof do
    begin

    vCont := vCont + 1;
    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO arrecadacao_grupo_local(id, cod_local, local, login, cod_usuario, usuario, codgtr, codigo_grupo, ');
    qryRegistro.SQL.Add(' descricao_grupo,codtrb, desres, desmin, data_movimento, valor_lancado, valor_atualizado,                                    ');
    qryRegistro.SQL.Add(' juros, multa, desconto, emolumento, valor_total, datahora_processamento)                                    ');
    qryRegistro.SQL.Add('VALUES (nextval(''arrecadacao_grupo_local_id_seq''),:cod_local, :local, :login, :cod_usuario, :usuario, :codgtr, :codigo_grupo,   ');
    qryRegistro.SQL.Add(' :descricao_grupo,:codtrb, :desres, :desmin, :datamov, :valor_lancado, :valor_atualizado,                              ');
    qryRegistro.SQL.Add(' :juros, :multa, :desconto, :emolumento, :valor_total, :dt_processamento)                              ');

    qryRegistro.ParamByName('cod_local').AsInteger      := qryBuscaSIAT.fieldbyname('coduor').AsInteger;

    if Trim(qryBuscaSIAT.fieldbyname('local').AsString) <> '' then
      qryRegistro.ParamByName('local').AsString           := Trim(qryBuscaSIAT.fieldbyname('local').AsString)
    else
      qryRegistro.ParamByName('local').AsString           := 'SEM LOCAL';


    if Trim(qryBuscaSIAT.fieldbyname('login').AsString) <> '' then
      qryRegistro.ParamByName('login').AsString           := Trim(qryBuscaSIAT.fieldbyname('login').AsString)
    else
      qryRegistro.ParamByName('login').AsString           := 'SEM LOGIN';


    qryRegistro.ParamByName('cod_usuario').AsInteger    := qryBuscaSIAT.fieldbyname('codusu').AsInteger;


    if Trim(qryBuscaSIAT.fieldbyname('usuario').AsString) <> '' then
      qryRegistro.ParamByName('usuario').AsString         := Trim(qryBuscaSIAT.fieldbyname('usuario').AsString)
    else
      qryRegistro.ParamByName('usuario').AsString         := 'SEM USUARIO';


    qryRegistro.ParamByName('codgtr').AsInteger         := qryBuscaSIAT.fieldbyname('codgtr').AsInteger;
    qryRegistro.ParamByName('codigo_grupo').AsString    := Trim(qryBuscaSIAT.fieldbyname('codigoGrupo').AsString);
    qryRegistro.ParamByName('descricao_grupo').AsString := Trim(qryBuscaSIAT.fieldbyname('descricaoGrupo').AsString);
    qryRegistro.ParamByName('codtrb').AsString          := Trim(qryBuscaSIAT.fieldbyname('codtrb').AsString);
    qryRegistro.ParamByName('desres').AsString          := Trim(qryBuscaSIAT.fieldbyname('desres').AsString);
    qryRegistro.ParamByName('desmin').AsString          := Trim(qryBuscaSIAT.fieldbyname('desmin').AsString);
    qryRegistro.ParamByName('datamov').Value            := qryBuscaSIAT.fieldbyname('datmva').AsDateTime;
    qryRegistro.ParamByName('valor_lancado').AsFloat    := qryBuscaSIAT.fieldbyname('valLan').AsFloat;
    qryRegistro.ParamByName('valor_atualizado').AsFloat := qryBuscaSIAT.fieldbyname('valatu').AsFloat;
    qryRegistro.ParamByName('juros').AsFloat            := qryBuscaSIAT.fieldbyname('juros').AsFloat;
    qryRegistro.ParamByName('multa').AsFloat            := qryBuscaSIAT.fieldbyname('multa').AsFloat;
    qryRegistro.ParamByName('desconto').AsFloat         := qryBuscaSIAT.fieldbyname('desconto').AsFloat;
    qryRegistro.ParamByName('emolumento').AsFloat       := qryBuscaSIAT.fieldbyname('emolumento').AsFloat;
    qryRegistro.ParamByName('valor_total').AsFloat      := qryBuscaSIAT.fieldbyname('valorTotal').AsFloat;
    qryRegistro.ParamByName('dt_processamento').Value   := vDataHora;
    qryRegistro.ExecSQL;


    qryBuscaSIAT.Next;
    end;//while

  lblFim.Caption := 'FINAL: '+timetostr(now);


end;

procedure TfrmImportaICMS.Endereco_Tomadores;
begin

  qryTomador.Close;
  qryTomador.SQL.Clear;
  qryTomador.SQL.Add('select * from tomador_base order by cpfcnpj, ultimanota ');
  qryTomador.open;

  vCont := 0;
  vDataHora := now;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryTomador.recordcount);
  Application.ProcessMessages;

  while not qryTomador.eof do
    begin

    vCont := vCont + 1;
    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont);
    Status.Panels[2].Text := Trim(qryTomador.fieldbyname('cpfcnpj').AsString);
    Application.ProcessMessages;

    vIDPessoa := TrazIDPessoaExterna(Trim(qryTomador.fieldbyname('cpfcnpj').AsString));

    if vIDPessoa > 0 then
      begin

      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('update pessoa_externa set tomador =:tomador, email=:email, telefone=:telefone  ');
      qryRegistro.sql.add('where id =:id     ');
      qryRegistro.ParamByName('id').Value       := vIDPessoa;
      qryRegistro.ParamByName('tomador').Value  := True;
      qryRegistro.ParamByName('email').Value    := Trim(qryTomador.fieldbyname('email').AsString);
      qryRegistro.ParamByName('telefone').Value := Trim(qryTomador.fieldbyname('telefone').AsString);

      qryRegistro.ExecSQL;


      end// if vIDPessoa > 0 then
    else
      begin

      if Length(Trim(qryTomador.fieldbyname('cpfcnpj').AsString)) <= 11 then
        vTipo := 'PF'
      else
        vTipo := 'PJ';

      qryRegistro.Close;
      qryRegistro.SQL.Clear;
      qryRegistro.SQL.Add('INSERT INTO pessoa_externa(id, nome, tipo, cpf_cnpj, cepisa, receita, situacao_receita,tomador,email,telefone)            ');
      qryRegistro.SQL.Add('VALUES (nextval(''pessoa_externa_id_seq''), :nome, :tipo, :cpf_cnpj, :cepisa, :receita, :situacao,:tomador,:email,:telefone) ');

      qryRegistro.ParamByName('nome').Value     := Trim(qryTomador.fieldbyname('nomerazao').AsString);
      qryRegistro.ParamByName('tipo').Value     := Trim(vTipo);
      qryRegistro.ParamByName('cpf_cnpj').Value := Trim(qryTomador.fieldbyname('cpfcnpj').AsString);
      qryRegistro.ParamByName('cepisa').Value   := False;
      qryRegistro.ParamByName('receita').Value  := False;
      qryRegistro.ParamByName('situacao').Value := '';
      qryRegistro.ParamByName('tomador').Value  := True;
      qryRegistro.ParamByName('email').Value    := Trim(qryTomador.fieldbyname('email').AsString);
      qryRegistro.ParamByName('telefone').Value := Trim(qryTomador.fieldbyname('telefone').AsString);

      qryRegistro.ExecSQL;

      //pesquisando
      qryTributo.Close;
      qryTributo.SQL.Clear;
      qryTributo.SQL.Add('select Max(id) as Ultimo from pessoa_externa');
      qryTributo.open;

      vIDPessoa := qryTributo.fieldbyname('ultimo').AsInteger;

      end;

    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO endereco_tomador(id, logradouro, numero, complemento, bairro, cep, ');
    qryRegistro.SQL.Add(' municipio, uf, datahora_processamento, pessoa_externa_id)                     ');
    qryRegistro.SQL.Add('VALUES (nextval(''endereco_tomador_id_seq''), :logradouro, :numero, :complemento, :bairro, :cep, ');
    qryRegistro.SQL.Add(' :municipio, :uf, :datahora, :pessoa_externa_id)                     ');
    qryRegistro.ParamByName('logradouro').Value  := Trim(qryTomador.fieldbyname('logradouro').AsString);
    qryRegistro.ParamByName('numero').Value      := Trim(qryTomador.fieldbyname('numero').AsString);
    qryRegistro.ParamByName('complemento').Value := Trim(qryTomador.fieldbyname('complemento').AsString);
    qryRegistro.ParamByName('bairro').Value      := Trim(qryTomador.fieldbyname('bairro').AsString);
    qryRegistro.ParamByName('cep').Value         := Trim(qryTomador.fieldbyname('cep').AsString);
    qryRegistro.ParamByName('municipio').Value   := Trim(qryTomador.fieldbyname('cidade').AsString);
    qryRegistro.ParamByName('uf').Value          := Trim(qryTomador.fieldbyname('uf').AsString);
    qryRegistro.ParamByName('datahora').Value          := vDataHora;
    qryRegistro.ParamByName('pessoa_externa_id').Value := vIDPessoa;


    Try
      Try
      qryRegistro.ExecSQL;
      Except

      End;
      Finally
    End;

    qryTomador.Next;
    
    end;
    ShowMessage('Final da Importação - Tomadores.');

end;

procedure TfrmImportaICMS.Processa_Rendimentos_Autonomos;
begin
  XlsToStringGrid(strgDados,'D:\SEMF\Receita-Rendimentos\Rendimentos_Profissionais_Liberais_Teresina.xls');


  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(x);
  Application.ProcessMessages;

  for i := 1 to strgDados.rowcount -1 do
    begin
    if trim(strgDados.cells[0,i]) <> '' then
      begin
      vCont := vCont + 1;
      vCPF  := Trim(strgDados.cells[0,i]);
      vCPF  := StringReplace( vCPF, '.'  , '' , [rfReplaceAll]);
      vCPF  := StringReplace( vCPF, '-'  , '' , [rfReplaceAll]);

      vCPF  := StrZeroString(vCPF,11);


      Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' - CPF: '+Trim(vCPF);
      Application.ProcessMessages;


      if not PesquisaAutonomoCadastrado(vCPF) then
        begin
        qryRegistro.Close;
        qryRegistro.SQL.Clear;
        qryRegistro.SQL.Add('INSERT INTO autonomo_receita(id, nome, cpf)  ');
        qryRegistro.SQL.Add('VALUES (nextval(''autonomo_receita_id_seq''), :nome, :cpf) ');
        qryRegistro.ParamByName('nome').Asstring := Trim(strgDados.cells[1,i]);
        qryRegistro.ParamByName('cpf').Asstring  := vCPF;
        qryRegistro.ExecSQL;

        qryRegistro.Close; qryRegistro.SQL.Clear;
        qryRegistro.SQL.Add('select Max(id) as Ultimo from autonomo_receita');
        qryRegistro.open;

        vIDPessoa := qryRegistro.fieldbyname('ultimo').AsInteger;

        end;// if PesquisaAutonomoCadastrado(vCPF)


        qryRegistro.Close;
        qryRegistro.SQL.Clear;
        qryRegistro.SQL.Add('INSERT INTO rendimentos_autonomo(id, natureza_ocupacao, ocupacao_atual, ano, municipio, ');
        qryRegistro.SQL.Add(' valorrecebido_pj, valorrecebido_pf, valorrecebido_acumulado, valorrecebido_exterior,               ');
        qryRegistro.SQL.Add(' rendimento_isento_naotrib, rendimento_sujeito_atrib, rendimento_ir_suspensa,     ');
        qryRegistro.SQL.Add(' autonomo_receita_id)                                                             ');
        qryRegistro.SQL.Add('VALUES (nextval(''rendimentos_autonomo_id_seq''), :natureza_ocupacao, :ocupacao_atual, :ano, :municipio, ');
        qryRegistro.SQL.Add(' :valorrecebido_pj, :valorrecebido_pf, :valorrecebido_acumulado, :valorrecebido_exterior,               ');
        qryRegistro.SQL.Add(' :rendimento_isento_naotrib, :rendimento_sujeito_atrib, :rendimento_ir_suspensa,     ');
        qryRegistro.SQL.Add(' :autonomo_receita_id)                                                               ');

        qryRegistro.ParamByName('natureza_ocupacao').Asstring      := Trim(strgDados.cells[2,i]);
        qryRegistro.ParamByName('ocupacao_atual').Asstring         := Trim(strgDados.cells[3,i]);
        qryRegistro.ParamByName('ano').AsInteger                   := strtoint(strgDados.cells[4,i]);
        qryRegistro.ParamByName('municipio').Asstring              := UpperCase(strgDados.cells[5,i]);

        qryRegistro.ParamByName('valorrecebido_pj').Value          := strtofloat(StringReplace( strgDados.cells[6,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('valorrecebido_pf').Value          := strtofloat(StringReplace( strgDados.cells[7,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('valorrecebido_acumulado').Value   := strtofloat(StringReplace( strgDados.cells[8,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('valorrecebido_exterior').Value    := strtofloat(StringReplace( strgDados.cells[9,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('rendimento_isento_naotrib').Value := strtofloat(StringReplace( strgDados.cells[10,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('rendimento_sujeito_atrib').Value  := strtofloat(StringReplace( strgDados.cells[11,i], '.'  , '' , [rfReplaceAll]));
        qryRegistro.ParamByName('rendimento_ir_suspensa').Value    := strtofloat(StringReplace( strgDados.cells[12,i], '.'  , '' , [rfReplaceAll]));

        qryRegistro.ParamByName('autonomo_receita_id').Value       := vIDPessoa;


        qryRegistro.ExecSQL;

        
      end;
    end;// for i := 1 ...

  showmessage('Final da leitura do GRID...');

end;

function TfrmImportaICMS.PesquisaAutonomoCadastrado(cpf: string): boolean;
begin
  Result := False;

  qryAutonomo.Close;
  qryAutonomo.SQL.Clear;
  qryAutonomo.SQL.Add('select id from autonomo_receita      ');
  qryAutonomo.SQL.Add('where cpf =:cpf           ');
  qryAutonomo.ParamByName('cpf').Value := cpf;
  qryAutonomo.open;

  if qryAutonomo.RecordCount > 0 then
    begin
    Result := True;
    vIDPessoa := qryAutonomo.fieldbyname('id').AsInteger;

    end;

end;

procedure TfrmImportaICMS.Atualiza_NFE_Autonomos;
begin

  qryNFSE.Close;
  qryNFSE.SQL.Clear;
  qryNFSE.SQL.Add('SELECT PREST_CPF_CNPJ,                          ');
  qryNFSE.SQL.Add('	extract(year from data_hora_emissao) as ano,   ');
  qryNFSE.SQL.Add('	count(prest_cpf_cnpj) as quantidade,           ');
  qryNFSE.SQL.Add('	SUM (valor_nota) AS valor,                     ');
  qryNFSE.SQL.Add('	SUM (valor_iss) AS iss                         ');
  qryNFSE.SQL.Add('FROM NOTA_FISCAL_AVULSA                         ');
//  qryNFSE.SQL.Add('where  extract(year from data_hora_emissao) = 2011  ');
  qryNFSE.SQL.Add('GROUP BY PREST_CPF_CNPJ, extract(year from data_hora_emissao) ');
  qryNFSE.SQL.Add('order by PREST_CPF_CNPJ, extract(year from data_hora_emissao) ');
  qryNFSE.open;

  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(qryNFSE.RecordCount);
  Application.ProcessMessages;

  while not qryNFSE.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' - CPF: '+Trim(qryNFSE.fieldbyname('PREST_CPF_CNPJ').AsString);
    Application.ProcessMessages;


    if PesquisaAutonomoCadastrado(qryNFSE.fieldbyname('PREST_CPF_CNPJ').AsString) then
      begin

      qryRegistro.Close;
      qryRegistro.SQL.Clear;
      qryRegistro.SQL.Add('INSERT INTO notasfiscais_autonomo(id, ano, quantidade, valor_nota, valor_iss, autonomo_receita_id)      ');
      qryRegistro.SQL.Add('VALUES (nextval(''notasfiscais_autonomo_id_seq''), :ano, :quantidade, :valor_nota, :valor_iss, :pessoa) ');
      qryRegistro.ParamByName('ano').Value        := qryNFSE.fieldbyname('ano').AsInteger;
      qryRegistro.ParamByName('quantidade').Value := qryNFSE.fieldbyname('quantidade').AsInteger;
      qryRegistro.ParamByName('valor_nota').Value := qryNFSE.fieldbyname('valor').AsFloat;
      qryRegistro.ParamByName('valor_iss').Value  := qryNFSE.fieldbyname('iss').AsFloat;
      qryRegistro.ParamByName('pessoa').Value     := vIDPessoa;
      qryRegistro.ExecSQL;

      end;

    qryNFSE.Next;
    end;//while ...
  showmessage('Fim.');

end;

procedure TfrmImportaICMS.Atualiza_CMC_Autonomos;
begin
  qryAutonomo.Close;
  qryAutonomo.SQL.Clear;
  qryAutonomo.SQL.Add('select id, cpf from autonomo_receita     ');
  qryAutonomo.open;

  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(qryAutonomo.RecordCount);
  Application.ProcessMessages;

  while not qryAutonomo.eof do
    begin
    vCont := vCont + 1;

    Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' - CPF: '+Trim(qryAutonomo.fieldbyname('cpf').AsString);
    Application.ProcessMessages;

    qryDados.Close;
    qryDados.SQL.Clear;
    qryDados.SQL.Add('select eco.insmun, enq.situac, item.descom, eco.cpfcnpj, enq.datini, enq.datfim ');
    qryDados.SQL.Add('from siatthe.tbleco eco                                              ');
    qryDados.SQL.Add('inner join siatthe.tblecoeqditm enq on enq.codeco = eco.codeco       ');
    qryDados.SQL.Add('inner join siatthe.tbleqditm item on item.codeqditm = enq.codeqditm  ');
    qryDados.SQL.Add('where eco.cpfcnpj =:cpf                                              ');

    qryDados.Parameters.ParamByName('cpf').Value := qryAutonomo.fieldbyname('cpf').asstring;
    qryDados.open;

    if qryDados.RecordCount > 0 then
      begin
      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('update autonomo_receita set insmun =:insmun, descom =:descom, situac =:situac, data_ini =:ini, data_fim =:fim ');
      qryRegistro.sql.add('where id =:id     ');
      qryRegistro.ParamByName('id').Value     := qryAutonomo.fieldbyname('id').asinteger;
      qryRegistro.ParamByName('insmun').Value := qryDados.fieldbyname('insmun').AsString;
      qryRegistro.ParamByName('descom').Value := qryDados.fieldbyname('descom').AsString;
      qryRegistro.ParamByName('situac').Value := qryDados.fieldbyname('situac').AsString;

      if qryDados.fieldbyname('datini').AsString <> '' then
        qryRegistro.ParamByName('ini').Value    := qryDados.fieldbyname('datini').AsDateTime
      else
        qryRegistro.ParamByName('ini').Value    := null;


      if qryDados.fieldbyname('datfim').AsString <> '' then
        qryRegistro.ParamByName('fim').Value    := qryDados.fieldbyname('datfim').AsDateTime
      else
        qryRegistro.ParamByName('fim').Value    := null;

      qryRegistro.ExecSQL;


      end;


    qryAutonomo.Next;
    end;

end;

procedure TfrmImportaICMS.Importa_Planilha_NotasFiscaisSAT;
begin
  Status.Panels[0].Text := 'Lendo Planilha...';
  Application.ProcessMessages;
  
  XlsToStringGrid(strgDados,'C:\SEMF\Atualiza_Avulso\Notas Fiscais SAT new.xlsx');


  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(x);
  Application.ProcessMessages;



  for i := 1 to strgDados.rowcount -1 do
    begin
    if trim(strgDados.cells[0,i]) <> '' then
      begin

      vCont := vCont + 1;
      vCPF  := Trim(strgDados.cells[3,i]);
      vCPF  := StringReplace( vCPF, '.'  , '' , [rfReplaceAll]);
      vCPF  := StringReplace( vCPF, '-'  , '' , [rfReplaceAll]);
      vCPF  := StringReplace( vCPF, '/'  , '' , [rfReplaceAll]);

      //vCPF  := StrZeroString(vCPF,11);


      Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' - CPF: '+Trim(vCPF);
      Application.ProcessMessages;

      qryTeste.Close;
      qryTeste.SQL.Clear;
      qryTeste.SQL.Add('INSERT INTO notasfiscais_sat(data_nf, num_nf, serie_nf, cpf_cnpj_prestador, nome_prestador, ');
      qryTeste.SQL.Add('logradouro_prestador, bairro_prestador, cep_prestador, municipio_prestador,   ');
      qryTeste.SQL.Add('estado_prestador, cpf_cnpj_tomador, nome_tomador, logradouro_tomador,         ');
      qryTeste.SQL.Add('bairro_tomador, cep_tomador, municipio_tomador, estado_tomador,               ');
      qryTeste.SQL.Add('valor_nf, aliquota, valor_tributo, cod_atv_fiscal_sat, cod_atv_fiscal_smt,    ');
      qryTeste.SQL.Add('descr_atv_sat, descr_atv_smt)                                                 ');
      qryTeste.SQL.Add('VALUES (:data_nf, :num_nf, :serie_nf, :cpf_cnpj_prestador, :nome_prestador,       ');
      qryTeste.SQL.Add(':logradouro_prestador, :bairro_prestador, :cep_prestador, :municipio_prestador,   ');
      qryTeste.SQL.Add(':estado_prestador, :cpf_cnpj_tomador, :nome_tomador, :logradouro_tomador,         ');
      qryTeste.SQL.Add(':bairro_tomador, :cep_tomador, :municipio_tomador, :estado_tomador,               ');
      qryTeste.SQL.Add(':valor_nf, :aliquota, :valor_tributo, :cod_atv_fiscal_sat, :cod_atv_fiscal_smt,   ');
      qryTeste.SQL.Add(':descr_atv_sat, :descr_atv_smt)                                                   ');

      qryTeste.ParamByName('data_nf').Asstring  := trim(strgDados.cells[0,i]);
      qryTeste.ParamByName('num_nf').Asstring  := trim(strgDados.cells[1,i]);
      qryTeste.ParamByName('serie_nf').Asstring  := trim(strgDados.cells[2,i]);
      qryTeste.ParamByName('cpf_cnpj_prestador').Asstring  := vCPF;
      qryTeste.ParamByName('nome_prestador').Asstring  := trim(strgDados.cells[4,i]);
      qryTeste.ParamByName('logradouro_prestador').Asstring  := trim(strgDados.cells[5,i]);
      qryTeste.ParamByName('bairro_prestador').Asstring  := trim(strgDados.cells[6,i]);
      qryTeste.ParamByName('cep_prestador').Asstring  := trim(strgDados.cells[7,i]);
      qryTeste.ParamByName('municipio_prestador').Asstring  := trim(strgDados.cells[8,i]);
      qryTeste.ParamByName('estado_prestador').Asstring  := trim(strgDados.cells[9,i]);
      qryTeste.ParamByName('cpf_cnpj_tomador').Asstring  := trim(strgDados.cells[10,i]);
      qryTeste.ParamByName('nome_tomador').Asstring  := trim(strgDados.cells[11,i]);
      qryTeste.ParamByName('logradouro_tomador').Asstring  := trim(strgDados.cells[12,i]);
      qryTeste.ParamByName('bairro_tomador').Asstring  := trim(strgDados.cells[13,i]);
      qryTeste.ParamByName('cep_tomador').Asstring  := trim(strgDados.cells[14,i]);
      qryTeste.ParamByName('municipio_tomador').Asstring  := trim(strgDados.cells[15,i]);
      qryTeste.ParamByName('estado_tomador').Asstring  := trim(strgDados.cells[16,i]);
      qryTeste.ParamByName('valor_nf').Asstring  := trim(strgDados.cells[17,i]);
      qryTeste.ParamByName('aliquota').Asstring  := trim(strgDados.cells[18,i]);
      qryTeste.ParamByName('valor_tributo').Asstring  := trim(strgDados.cells[19,i]);
      qryTeste.ParamByName('cod_atv_fiscal_sat').Asstring  := trim(strgDados.cells[20,i]);
      qryTeste.ParamByName('cod_atv_fiscal_smt').Asstring  := trim(strgDados.cells[21,i]);
      qryTeste.ParamByName('descr_atv_sat').Asstring  := trim(strgDados.cells[22,i]);
      qryTeste.ParamByName('descr_atv_smt').Asstring  := trim(strgDados.cells[23,i]);


      Try
        Try
        qryTeste.ExecSQL;
        Except

        End;
      Finally
      End;




      end; // if trim(strgDados.cells[0,i]) <> '' ...


    end; // for i := 1 to strgDados.rowcount -1 do ...

   showmessage('Final da leitura do GRID...');
end;

procedure TfrmImportaICMS.conecta_150BeforeConnect(Sender: TObject);
begin
  conecta_150.Properties.Add('cliente_enconding=latin1');
  conecta_150.Properties.Add('codepage=latin1');

end;

procedure TfrmImportaICMS.ProcessaMalha_SimplesNacional;
begin
  Linha   := 0;
  Entrada := '';  vNomeArquivo := '';
  vCaminhoDestino := 'C:\SEMF\Simples Nacional\Processados\';


  vCont := 0;
  vID   := 0;

  lblDestino.Caption := 'Caminho Destino: '+vCaminhoDestino;

  FileSearch(edtCaminho.Text,'*.txt', false);

  for i := 0 to ListBox1.Items.Count-1 do
    begin

    vNomeArquivo := copy(listbox1.Items[i],35,50);
    Linha := 0;

    If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
      begin
      AssignFile(ArqTexto,edtCaminho.Text + Trim(vNomeArquivo));
      Reset(ArqTexto);

      while not Eof(ArqTexto) do
        begin
        Linha := Linha + 1;
        Readln(ArqTexto,Entrada);

        vTipo  := Copy(Entrada,1,5);
        vCont := vCont + 1;

        Status.Panels[0].Text := 'Registros: ' + IntToStr(Linha)+' Tipo: '+vTipo;
        Application.ProcessMessages;

        if vTipo = '00000' then
          begin
          vMunicipio := '';
          vCNPJ_Matriz := ''; vPA := '';

          Separador_RegistrosSN;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;

          qryApuracao.SQL.Add('INSERT INTO dados_apuracao(id, reg, id_declaracao, num_recibo, num_autenticacao, dt_transmissao, ');
          qryApuracao.SQL.Add(' data_transmissao, versao, cnpj_matriz, nome, cod_tom, optante, abertura, data_abertura,                           ');
          qryApuracao.SQL.Add(' pa, rpa, rzfs, im, operacao, regime, rpac, rpa_int, rpa_ext, nome_arquivo)                                    ');
          qryApuracao.SQL.Add('VALUES (nextval(''dados_apuracao_id_seq''), :reg, :id_declaracao, :num_recibo, :num_autenticacao, :dt_transmissao, ');
          qryApuracao.SQL.Add(' :data_transmissao, :versao, :cnpj_matriz, :nome, :cod_tom, :optante, :abertura, :data_abertura,                     ');
          qryApuracao.SQL.Add(' :pa, :rpa, :rzfs, :im, :operacao, :regime, :rpac, :rpa_int, :rpa_ext, :nome_arquivo)                         ');

          qryApuracao.ParamByName('reg').Value              := vUm;
          qryApuracao.ParamByName('id_declaracao').Value    := vDois;
          qryApuracao.ParamByName('num_recibo').Value       := vTres;
          qryApuracao.ParamByName('num_autenticacao').Value := vQuatro;
          qryApuracao.ParamByName('dt_transmissao').Value   := vCinco;

          if Trim(vCinco) <> '' then
            qryApuracao.ParamByName('data_transmissao').Value := Copy(vCinco,1,4)+'-'+Copy(vCinco,5,2)+'-'+Copy(vCinco,7,2)
          else
            qryApuracao.ParamByName('data_transmissao').Value := null;


          qryApuracao.ParamByName('versao').Value      := vSeis;
          qryApuracao.ParamByName('cnpj_matriz').Value := vSete;
          qryApuracao.ParamByName('nome').Value        := vOito;
          qryApuracao.ParamByName('cod_tom').Value     := vNove;
          qryApuracao.ParamByName('optante').Value     := vDez;
          qryApuracao.ParamByName('abertura').Value    := vOnze;//AAAAMMDD

          if Trim(vOnze) <> '' then
            qryApuracao.ParamByName('data_abertura').Value := Copy(vOnze,1,4)+'-'+Copy(vOnze,5,2)+'-'+Copy(vOnze,7,2)
          else
            qryApuracao.ParamByName('data_abertura').Value := null;

          if Trim(vTreze) = '' then vTreze := '0';
          if Trim(vQuatorze) = '' then vQuatorze := '0';
          if Trim(vQuinze) = '' then vQuinze := '0';

          qryApuracao.ParamByName('pa').Value       := vDoze;
          qryApuracao.ParamByName('rpa').Value      := strtofloat(vTreze);
          qryApuracao.ParamByName('rzfs').Value     := strtofloat(vQuatorze);
          qryApuracao.ParamByName('im').Value       := strtofloat(vQuinze);
          qryApuracao.ParamByName('operacao').Value := vDezesseis;
          qryApuracao.ParamByName('regime').Value   := vDezessete;

          if Trim(vDezoito) = '' then vDezoito := '0';   if Trim(vDezenove) = '' then vDezenove := '0';
           if Trim(vVinte) = '' then vVinte := '0';

          qryApuracao.ParamByName('rpac').Value     := strtofloat(vDezoito);
          qryApuracao.ParamByName('rpa_int').Value  := strtofloat(vDezenove);
          qryApuracao.ParamByName('rpa_ext').Value  := strtofloat(vVinte);
          qryApuracao.ParamByName('nome_arquivo').Value  := Trim(vNomeArquivo);
          qryApuracao.ExecSQL;

          vMunicipio := vNove; vCNPJ_Matriz := vSete; vPA := vDoze;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;
          qryApuracao.SQL.Add('select Max(id) as ultimo from dados_apuracao');
          qryApuracao.open;

          vID := qryApuracao.fieldbyname('ultimo').AsInteger;


          //Incluir pessoa.
          qryApuracao.Close;
          qryApuracao.SQL.Clear;
          qryApuracao.SQL.Add('INSERT INTO pessoa(id, nome, cpf_cnpj)  ');
          qryApuracao.SQL.Add('VALUES (nextval(''pessoa_id_seq''), :nome, :cpf_cnpj)  ');
          qryApuracao.ParamByName('nome').Value     := vOito;
          qryApuracao.ParamByName('cpf_cnpj').Value := vSete;

          Try
            Try
              qryApuracao.ExecSQL;
            Except

            End;
          Finally
          End;


          end;

        if vTipo = '03000' then
          begin
          vCNPJ_Filial := '';

          Separador_RegistrosSN;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;
          qryApuracao.SQL.Add('INSERT INTO estabelecimento_filial(id, reg, cnpj, uf, cod_tom, vltotal, ime,      ');
          qryApuracao.SQL.Add(' limite, lim_ultrapassado_pa, prex, prex2, dados_apuracao_id)                     ');
          qryApuracao.SQL.Add('VALUES (nextval(''estabelecimento_filial_id_seq''), :reg, :cnpj, :uf, :cod_tom,   ');
          qryApuracao.SQL.Add(' :vltotal, :ime, :limite, :lim_ultrapassado_pa, :prex, :prex2, :dados_apuracao_id)');
          qryApuracao.ParamByName('reg').Value      := vUm;
          qryApuracao.ParamByName('cnpj').Value     := vDois;
          qryApuracao.ParamByName('uf').Value       := vTres;
          qryApuracao.ParamByName('cod_tom').Value  := vQuatro;

          if Trim(vCinco) = '' then vCinco := '0';
          if Trim(vSeis) = '' then vSeis := '0';
          if Trim(vSete) = '' then vSete := '0';
          if Trim(vNove) = '' then vNove := '0';
          if Trim(vDez) = '' then vDez := '0';

          qryApuracao.ParamByName('vltotal').Value  := strtofloat(vCinco);
          qryApuracao.ParamByName('ime').Value      := strtofloat(vSeis);
          qryApuracao.ParamByName('limite').Value   := strtofloat(vSete);
          qryApuracao.ParamByName('lim_ultrapassado_pa').Value := vOito;
          qryApuracao.ParamByName('prex').Value     := strtofloat(vNove);
          qryApuracao.ParamByName('prex2').Value    := strtofloat(vDez);
          qryApuracao.ParamByName('dados_apuracao_id').Value := vID;
          qryApuracao.ExecSQL;

          vCNPJ_Filial := vDois;

          end; //if vTipo = '03000'

        if vTipo = '03100' then
          begin
          vTP_Atividade := '';
          Separador_RegistrosSN;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;
          qryApuracao.SQL.Add('INSERT INTO atividade_estabelecimento(id, reg, tipo, vltotal, dados_apuracao_id) ');
          qryApuracao.SQL.Add('VALUES (nextval(''atividade_estabelecimento_id_seq''), :reg, :tipo, :vltotal, :dados_apuracao_id) ');
          qryApuracao.ParamByName('reg').Value      := vUm;
          qryApuracao.ParamByName('tipo').Value     := vDois;
          if Trim(vTres) = '' then vTres := '0';
          qryApuracao.ParamByName('vltotal').Value  := strtofloat(vTres);
          qryApuracao.ParamByName('dados_apuracao_id').Value := vID;
          qryApuracao.ExecSQL;
          vTP_Atividade := vDois;
          end; //if vTipo = '03100'

        if vTipo = '03110' then
          begin

          Separador_RegistrosSN;

  


          qryApuracao.Close;
          qryApuracao.SQL.Clear;
          qryApuracao.SQL.Add('INSERT INTO receita_atividade_a(id, reg, uf, cod_tom, valor, cofins, csll, icms, ');
          qryApuracao.SQL.Add(' inss, ipi, irpj, iss, pis, aliqapur, vlimposto, aliquota_cofins, valor_cofins,  ');
          qryApuracao.SQL.Add(' aliquota_csll, valor_csll, aliquota_icms, valor_icms, aliquota_inss,            ');
          qryApuracao.SQL.Add(' valor_inss, aliquota_ipi, valor_ipi, aliquota_irpj, valor_irpj,                 ');
          qryApuracao.SQL.Add(' aliquota_iss, valor_iss, aliquota_pis, valor_pis, diferenca,                    ');
          qryApuracao.SQL.Add(' maiortributo, dados_apuracao_id, cnpj_matriz,tipo, pa, municipio, cnpj_filial)  ');

          qryApuracao.SQL.Add('VALUES (nextval(''receita_atividade_a_id_seq''), :reg, :uf, :cod_tom, :valor, :cofins, :csll, :icms, ');
          qryApuracao.SQL.Add(' :inss, :ipi, :irpj, :iss, :pis, :aliqapur, :vlimposto, :aliquota_cofins, :valor_cofins,            ');
          qryApuracao.SQL.Add(' :aliquota_csll, :valor_csll, :aliquota_icms, :valor_icms, :aliquota_inss,         ');
          qryApuracao.SQL.Add(' :valor_inss, :aliquota_ipi, :valor_ipi, :aliquota_irpj, :valor_irpj,              ');
          qryApuracao.SQL.Add(' :aliquota_iss, :valor_iss, :aliquota_pis, :valor_pis, :diferenca,                 ');
          qryApuracao.SQL.Add(' :maiortributo, :dados_apuracao_id, :cnpj_matriz,:tipo, :pa, :municipio, :cnpj_filial) ');

          if Trim(vQuatro) = '' then vQuatro := '0';

          qryApuracao.ParamByName('reg').Value              := vUm;
          qryApuracao.ParamByName('uf').Value               := vDois;

          qryApuracao.ParamByName('cod_tom').Value          := vTres;
          qryApuracao.ParamByName('valor').Value            := StrtoFloat(vQuatro);
          qryApuracao.ParamByName('cofins').Value           := vCinco;
          qryApuracao.ParamByName('csll').Value             := vSeis;
          qryApuracao.ParamByName('icms').Value             := vSete;
          qryApuracao.ParamByName('inss').Value             := vOito;
          qryApuracao.ParamByName('ipi').Value              := vNove;
          qryApuracao.ParamByName('irpj').Value             := vDez;
          qryApuracao.ParamByName('iss').Value              := vOnze;
          qryApuracao.ParamByName('pis').Value              := vDoze;

          if Trim(vTreze) = '' then vTreze := '0';         if Trim(vQuatorze) = '' then vQuatorze := '0';
          if Trim(vQuinze) = '' then vQuinze := '0';       if Trim(vDezesseis) = '' then vDezesseis := '0';
          if Trim(vDezessete) = '' then vDezessete := '0'; if Trim(vDezoito) = '' then vDezoito := '0';
          if Trim(vDezenove) = '' then vDezenove := '0';   if Trim(vVinte) = '' then vVinte := '0';
          if Trim(vVinteUm) = '' then vVinteUm := '0';     if Trim(vVinteDois) = '' then vVinteDois := '0';
          if Trim(vVinteTres) = '' then vVinteTres := '0'; if Trim(vVinteQuatro) = '' then vVinteQuatro := '0';
          if Trim(vVinteCinco) = '' then vVinteCinco := '0';
          if Trim(vVinteSeis) = '' then vVinteSeis := '0';
          if Trim(vVinteSete) = '' then vVinteSete := '0';
          if Trim(vVinteOito) = '' then vVinteOito := '0';
          if Trim(vVinteNove) = '' then vVinteNove := '0';
          if Trim(vTrinta) = '' then vTrinta := '0';
          if Trim(vTrintaUm) = '' then vTrintaUm := '0';

          qryApuracao.ParamByName('aliqapur').Value         := StrtoFloat(vTreze);
          qryApuracao.ParamByName('vlimposto').Value        := StrtoFloat(vQuatorze);
          qryApuracao.ParamByName('aliquota_cofins').Value  := StrtoFloat(vQuinze);
          qryApuracao.ParamByName('valor_cofins').Value     := StrtoFloat(vDezesseis);
          qryApuracao.ParamByName('aliquota_csll').Value    := StrtoFloat(vDezessete);
          qryApuracao.ParamByName('valor_csll').Value       := StrtoFloat(vDezoito);
          qryApuracao.ParamByName('aliquota_icms').Value    := StrtoFloat(vDezenove);
          qryApuracao.ParamByName('valor_icms').Value       := StrtoFloat(vVinte);
          qryApuracao.ParamByName('aliquota_inss').Value    := StrtoFloat(vVinteUm);
          qryApuracao.ParamByName('valor_inss').Value       := StrtoFloat(vVinteDois);
          qryApuracao.ParamByName('aliquota_ipi').Value     := StrtoFloat(vVinteTres);
          qryApuracao.ParamByName('valor_ipi').Value        := StrtoFloat(vVinteQuatro);
          qryApuracao.ParamByName('aliquota_irpj').Value    := StrtoFloat(vVinteCinco);
          qryApuracao.ParamByName('valor_irpj').Value       := StrtoFloat(vVinteSeis);
          qryApuracao.ParamByName('aliquota_iss').Value     := StrtoFloat(vVinteSete);
          qryApuracao.ParamByName('valor_iss').Value        := StrtoFloat(vVinteOito);
          qryApuracao.ParamByName('aliquota_pis').Value     := StrtoFloat(vVinteNove);
          qryApuracao.ParamByName('valor_pis').Value        := StrtoFloat(vTrinta);
          qryApuracao.ParamByName('diferenca').Value        := StrtoFloat(vTrintaUm);
          qryApuracao.ParamByName('maiortributo').Value     := vTrintaDois;
          qryApuracao.ParamByName('dados_apuracao_id').Value := vID;
          qryApuracao.ParamByName('cnpj_matriz').Value := vCNPJ_Matriz;
          qryApuracao.ParamByName('tipo').Value        := vTP_Atividade;
          qryApuracao.ParamByName('pa').Value          := vPA;
          qryApuracao.ParamByName('municipio').Value   := vMunicipio;
          qryApuracao.ParamByName('cnpj_filial').Value := vCNPJ_Filial;


          qryApuracao.ExecSQL;

          end; //if vTipo = '03110'

        if vTipo = '03120' then
          begin

          Separador_RegistrosSN;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;

          qryApuracao.SQL.Add('INSERT INTO receita_atividade_b(id, reg, aliqapur, aliquota_cofins, valor_cofins, ');
          qryApuracao.SQL.Add(' aliquota_csll, valor_csll, aliquota_icms, valor_icms, aliquota_inss, valor_inss, ');
          qryApuracao.SQL.Add(' aliquota_ipi, valor_ipi, aliquota_irpj, valor_irpj, aliquota_iss,                ');
          qryApuracao.SQL.Add(' valor_iss, aliquota_pis, valor_pis, dados_apuracao_id)                           ');

          qryApuracao.SQL.Add('VALUES (nextval(''receita_atividade_b_id_seq''), :reg, :aliqapur, :aliquota_cofins, :valor_cofins, ');
          qryApuracao.SQL.Add(' :aliquota_csll, :valor_csll, :aliquota_icms, :valor_icms, :aliquota_inss, :valor_inss, ');
          qryApuracao.SQL.Add(' :aliquota_ipi, :valor_ipi, :aliquota_irpj, :valor_irpj, :aliquota_iss,                ');
          qryApuracao.SQL.Add(' :valor_iss, :aliquota_pis, :valor_pis, :dados_apuracao_id)                           ');


          if Trim(vDois) = '' then vDois := '0';       if Trim(vTres) = '' then vTres := '0';
          if Trim(vQuatro) = '' then vQuatro := '0';   if Trim(vCinco) = '' then vCinco := '0';
          if Trim(vSeis) = '' then vSeis := '0';       if Trim(vSete) = '' then vSete := '0';
          if Trim(vOito) = '' then vOito := '0';       if Trim(vNove) = '' then vNove := '0';
          if Trim(vDez) = '' then vDez := '0';         if Trim(vOnze) = '' then vOnze := '0';
          if Trim(vDoze) = '' then vDoze := '0';       if Trim(vTreze) = '' then vTreze := '0';
          if Trim(vQuatorze) = '' then vQuatorze := '0'; if Trim(vQuinze) = '' then vQuinze := '0';
          if Trim(vDezesseis) = '' then vDezesseis := '0';
          if Trim(vDezessete) = '' then vDezessete := '0';
          if Trim(vDezoito) = '' then vDezoito := '0';

          qryApuracao.ParamByName('reg').Value              := vUm;
          qryApuracao.ParamByName('aliqapur').Value         := StrtoFloat(vDois);
          qryApuracao.ParamByName('aliquota_cofins').Value  := StrtoFloat(vTres);
          qryApuracao.ParamByName('valor_cofins').Value     := StrtoFloat(vQuatro);
          qryApuracao.ParamByName('aliquota_csll').Value    := StrtoFloat(vCinco);
          qryApuracao.ParamByName('valor_csll').Value       := StrtoFloat(vSeis);
          qryApuracao.ParamByName('aliquota_icms').Value    := StrtoFloat(vSete);
          qryApuracao.ParamByName('valor_icms').Value       := StrtoFloat(vOito);
          qryApuracao.ParamByName('aliquota_inss').Value    := StrtoFloat(vNove);
          qryApuracao.ParamByName('valor_inss').Value       := StrtoFloat(vDez);
          qryApuracao.ParamByName('aliquota_ipi').Value     := StrtoFloat(vOnze);
          qryApuracao.ParamByName('valor_ipi').Value        := StrtoFloat(vDoze);
          qryApuracao.ParamByName('aliquota_irpj').Value    := StrtoFloat(vTreze);
          qryApuracao.ParamByName('valor_irpj').Value       := StrtoFloat(vQuatorze);
          qryApuracao.ParamByName('aliquota_iss').Value     := StrtoFloat(vQuinze);
          qryApuracao.ParamByName('valor_iss').Value        := StrtoFloat(vDezesseis);
          qryApuracao.ParamByName('aliquota_pis').Value     := StrtoFloat(vDezessete);
          qryApuracao.ParamByName('valor_pis').Value        := StrtoFloat(vDezoito);

          qryApuracao.ParamByName('dados_apuracao_id').Value := vID;
          qryApuracao.ExecSQL;

          end; //if vTipo = '03120'

        if vTipo = '03130' then
          begin

          Separador_RegistrosSN;

          qryApuracao.Close;
          qryApuracao.SQL.Clear;

          qryApuracao.SQL.Add('INSERT INTO receita_atividade_c(id, reg, aliqapur, aliquota_cofins, valor_cofins, ');
          qryApuracao.SQL.Add(' aliquota_csll, valor_csll, aliquota_icms, valor_icms, aliquota_inss, valor_inss, ');
          qryApuracao.SQL.Add(' aliquota_ipi, valor_ipi, aliquota_irpj, valor_irpj, aliquota_iss,                ');
          qryApuracao.SQL.Add(' valor_iss, aliquota_pis, valor_pis, dados_apuracao_id)                           ');

          qryApuracao.SQL.Add('VALUES (nextval(''receita_atividade_c_id_seq''), :reg, :aliqapur, :aliquota_cofins, :valor_cofins, ');
          qryApuracao.SQL.Add(' :aliquota_csll, :valor_csll, :aliquota_icms, :valor_icms, :aliquota_inss, :valor_inss, ');
          qryApuracao.SQL.Add(' :aliquota_ipi, :valor_ipi, :aliquota_irpj, :valor_irpj, :aliquota_iss,                ');
          qryApuracao.SQL.Add(' :valor_iss, :aliquota_pis, :valor_pis, :dados_apuracao_id)                           ');


          if Trim(vDois) = '' then vDois := '0';       if Trim(vTres) = '' then vTres := '0';
          if Trim(vQuatro) = '' then vQuatro := '0';   if Trim(vCinco) = '' then vCinco := '0';
          if Trim(vSeis) = '' then vSeis := '0';       if Trim(vSete) = '' then vSete := '0';
          if Trim(vOito) = '' then vOito := '0';       if Trim(vNove) = '' then vNove := '0';
          if Trim(vDez) = '' then vDez := '0';         if Trim(vOnze) = '' then vOnze := '0';
          if Trim(vDoze) = '' then vDoze := '0';       if Trim(vTreze) = '' then vTreze := '0';
          if Trim(vQuatorze) = '' then vQuatorze := '0'; if Trim(vQuinze) = '' then vQuinze := '0';
          if Trim(vDezesseis) = '' then vDezesseis := '0';
          if Trim(vDezessete) = '' then vDezessete := '0';
          if Trim(vDezoito) = '' then vDezoito := '0';

          qryApuracao.ParamByName('reg').Value              := vUm;
          qryApuracao.ParamByName('aliqapur').Value         := StrtoFloat(vDois);
          qryApuracao.ParamByName('aliquota_cofins').Value  := StrtoFloat(vTres);
          qryApuracao.ParamByName('valor_cofins').Value     := StrtoFloat(vQuatro);
          qryApuracao.ParamByName('aliquota_csll').Value    := StrtoFloat(vCinco);
          qryApuracao.ParamByName('valor_csll').Value       := StrtoFloat(vSeis);
          qryApuracao.ParamByName('aliquota_icms').Value    := StrtoFloat(vSete);
          qryApuracao.ParamByName('valor_icms').Value       := StrtoFloat(vOito);
          qryApuracao.ParamByName('aliquota_inss').Value    := StrtoFloat(vNove);
          qryApuracao.ParamByName('valor_inss').Value       := StrtoFloat(vDez);
          qryApuracao.ParamByName('aliquota_ipi').Value     := StrtoFloat(vOnze);
          qryApuracao.ParamByName('valor_ipi').Value        := StrtoFloat(vDoze);
          qryApuracao.ParamByName('aliquota_irpj').Value    := StrtoFloat(vTreze);
          qryApuracao.ParamByName('valor_irpj').Value       := StrtoFloat(vQuatorze);
          qryApuracao.ParamByName('aliquota_iss').Value     := StrtoFloat(vQuinze);
          qryApuracao.ParamByName('valor_iss').Value        := StrtoFloat(vDezesseis);
          qryApuracao.ParamByName('aliquota_pis').Value     := StrtoFloat(vDezessete);
          qryApuracao.ParamByName('valor_pis').Value        := StrtoFloat(vDezoito);

          qryApuracao.ParamByName('dados_apuracao_id').Value := vID;
          qryApuracao.ExecSQL;

          end; //if vTipo = '03130'

        Status.Panels[1].Text := IntToStr(vCont)+' Arq.: '+Trim(vNomeArquivo);

        Application.ProcessMessages;

        end;//while


      end; //If FileExists

      CloseFile(ArqTexto);

      //Move arquivo lido
      If FileExists(edtCaminho.Text + Trim(vNomeArquivo)) then
        MoveFile(Pchar(edtCaminho.Text+Trim(vNomeArquivo)),Pchar(vCaminhoDestino+Trim(vNomeArquivo)));


    end;//for ...

end;

procedure TfrmImportaICMS.Separador_RegistrosSN;
begin
  //Limpando as variáveis
  vUm := ''; vDois := ''; vTres := ''; vQuatro := ''; vCinco := ''; vSeis := ''; vSete := ''; vOito := ''; vNove := ''; vDez := ''; vOnze := ''; vDoze := '';
  vTreze := ''; vQuatorze := ''; vQuinze := ''; vDezesseis := ''; vDezessete := ''; vDezoito := ''; vDezenove := ''; vVinte := '';
  vVinteUm := ''; vVinteDois := ''; vVinteTres := ''; vVinteQuatro := ''; vVinteCinco := ''; vVinteSeis := '';vVinteSete := ''; vVinteOito := '';
  vVinteNove := ''; vTrinta := ''; vTrintaUm := ''; vTrintaDois := '';



  If PosEx('|', Entrada) <> 0 then
    Item := LeftStr(Entrada, PosEx('|', Entrada) - 1);
  vUm := Item;


  If Pos('|', Entrada) <> 0 then
    Item := Copy(Entrada, Pos('|', Entrada)+1, (Length(Entrada)-Pos('|',Entrada)));
  vDois := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vTres := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vQuatro := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vCinco := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vSeis := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vSete := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vOito := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vNove := LeftStr(Item, PosEx('|', Item) - 1);

  //10
  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  if vUm = '03000' then vDez := Item
  else vDez := LeftStr(Item, PosEx('|', Item) - 1);


  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vOnze := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vDoze := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vTreze := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vQuatorze := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vQuinze := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vDezesseis := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vDezessete := LeftStr(Item, PosEx('|', Item) - 1);

  //18
  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vDezoito := LeftStr(Item, PosEx('|', Item) - 1);

  if (vUm = '03120') or (vUm = '03130')  then vDezoito := Item
  else vDezoito := LeftStr(Item, PosEx('|', Item) - 1);


  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vDezenove := LeftStr(Item, PosEx('|', Item) - 1);

  //20
  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  if vUm = '00000' then vVinte := Item
  else vVinte := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteUm := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteDois := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteTres := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteQuatro := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteCinco := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteSeis := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteSete := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteOito := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vVinteNove := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vTrinta := LeftStr(Item, PosEx('|', Item) - 1);
  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vTrintaUm := LeftStr(Item, PosEx('|', Item) - 1);

  If Pos('|', Item) <> 0 then
    Item := Copy(Item, Pos('|', Item)+1, (Length(Item)-Pos('|',Item)));
  vTrintaDois := Item;


end;

procedure TfrmImportaICMS.AlimentaNotaFiscalMensal;
begin
  //senha: d$f2014$% (ACESSAR NFE NO ORACLE)

  //Procedure criada em 15/01/2016 - Novo processamento de notas fiscais
  lblInicio.Caption := 'INÍCIO: '+timetostr(now);
  vCont := 0;


  qryNFSE.Close;
  qryNFSE.SQL.Clear;
  qryNFSE.SQL.Add('select                                                                                    ');
  qryNFSE.SQL.Add('nfse.MES_COMPETENCIA,                                                                     ');
  qryNFSE.SQL.Add('SUBSTR(nfse.MES_COMPETENCIA,1,2) AS MES,                                                  ');
  qryNFSE.SQL.Add('SUBSTR(nfse.MES_COMPETENCIA,3,4) AS ANO,                                                  ');
  qryNFSE.SQL.Add('sum(nfse.valor_nota) as total_valor_nota,                                                 ');
  qryNFSE.SQL.Add('sum(nfse.valor_servico) as total_valor_servico,                                           ');
  qryNFSE.SQL.Add('sum(nfse.VALOR_ISS) as total_valor_iss,                                                   ');
  qryNFSE.SQL.Add('nfse.prest_cpf_cnpj,                                                                      ');
  qryNFSE.SQL.Add('nfse.PREST_INSCRICAO_MUNICIPAL                                                            ');
  qryNFSE.SQL.Add('from nfse.nota_fiscal nfse                                                                ');
  qryNFSE.SQL.Add('left join NFSE.nota_fiscal_avulsa nfsa on nfsa.ID_NOTA_FISCAL_AVULSA=nfse.ID_NOTA_AVULSA  ');
  qryNFSE.SQL.Add('where                                                                                     ');
  qryNFSE.SQL.Add('  nfse.situacao_nf=''1'' and length(nfse.prest_cpf_cnpj)=14 and                             ');
  qryNFSE.SQL.Add('  (nfse.id_nota_avulsa is null or                                                          ');
  qryNFSE.SQL.Add('  length(nfse.prest_inscricao_municipal)=7 and nfsa.nf_contumaz=''S'')                      ');

  //qryNFSE.SQL.Add('and nfse.prest_cpf_cnpj = ''04691807000179''                                              ');
  qryNFSE.SQL.Add('and SUBSTR(nfse.MES_COMPETENCIA,3,4) = ''2016''     ');
  qryNFSE.SQL.Add('and SUBSTR(nfse.MES_COMPETENCIA,1,2) = ''12''    ');
  qryNFSE.SQL.Add('group by                                                                                  ');
  qryNFSE.SQL.Add('nfse.MES_COMPETENCIA,                                                                     ');
  qryNFSE.SQL.Add('nfse.prest_cpf_cnpj,                                                                      ');
  qryNFSE.SQL.Add('nfse.PREST_INSCRICAO_MUNICIPAL                                                            ');
  qryNFSE.SQL.Add('order by nfse.MES_COMPETENCIA, nfse.prest_cpf_cnpj                                        ');


 {
  qryNFSE.SQL.Add('select MES_COMPETENCIA,                      ');
  qryNFSE.SQL.Add('SUBSTR(MES_COMPETENCIA,1,2) AS MES,          ');
  qryNFSE.SQL.Add('SUBSTR(MES_COMPETENCIA,3,4) AS ANO,          ');
  qryNFSE.SQL.Add('sum(valor_nota) as total_valor_nota,         ');
  qryNFSE.SQL.Add('sum(valor_servico) as total_valor_servico,   ');
  qryNFSE.SQL.Add('sum(VALOR_ISS) as total_valor_iss,           ');
  qryNFSE.SQL.Add('prest_cpf_cnpj,                              ');
  qryNFSE.SQL.Add('PREST_INSCRICAO_MUNICIPAL                    ');
  qryNFSE.SQL.Add('from nota_fiscal                             ');
  qryNFSE.SQL.Add('where situacao_nf=''1'' and id_nota_avulsa is null and length(prest_cpf_cnpj)=14 ');
  //qryNFSE.SQL.Add('AND SUBSTR(MES_COMPETENCIA,3,4)=2009         ');


//  qryNFSE.SQL.Add('and prest_cpf_cnpj in (''93419083000139'',''81223445667749'')    ');

  qryNFSE.SQL.Add('group by MES_COMPETENCIA, prest_cpf_cnpj,    ');
  qryNFSE.SQL.Add('PREST_INSCRICAO_MUNICIPAL                    ');
  qryNFSE.SQL.Add('order by MES_COMPETENCIA, prest_cpf_cnpj     ');    }
  qryNFSE.open;

  Status.Panels[0].Text := 'Total: ' + IntToStr(qryNFSE.recordcount);
  Application.ProcessMessages;

  vDataHora := now;

  while not qryNFSE.eof do
    begin

    vNomeMes := '';
    case qryNFSE.fieldbyname('MES').Value of
      1:  vNomeMes := 'JANEIRO';
      2:  vNomeMes := 'FEVEREIRO';
      3:  vNomeMes := 'MARCO';
      4:  vNomeMes := 'ABRIL';
      5:  vNomeMes := 'MAIO';
      6:  vNomeMes := 'JUNHO';
      7:  vNomeMes := 'JULHO';
      8:  vNomeMes := 'AGOSTO';
      9:  vNomeMes := 'SETEMBRO';
      10: vNomeMes := 'OUTUBRO';
      11: vNomeMes := 'NOVEMBRO';
      12: vNomeMes := 'DEZEMBRO';
    end;

    vCont := vCont + 1;

    Status.Panels[1].Text := '1-Registros: ' + IntToStr(vCont);
    Application.ProcessMessages;


    qryDestino.Close;
    qryDestino.SQL.Clear;
    qryDestino.SQL.Add('INSERT INTO notafiscal_mensal(id, cnpj_cpf, inscricao_municipal, valor_total_nota, valor_total_servico, ');
    qryDestino.SQL.Add(' valor_total_iss, ano, mes, nome_mes, data, competencia, datahora_processamento, pessoa_sistema_siat_id)  ');

    qryDestino.SQL.Add('VALUES (nextval(''notafiscal_mensal_id_seq''), :cnpj_cpf, :inscricao_municipal, :valor_total_nota, :valor_total_servico, ');
    qryDestino.SQL.Add(' :valor_total_iss, :ano, :mes, :nome_mes, :data, :competencia, :datahora, :pessoa_sistema_siat_id)  ');

    qryDestino.ParamByName('cnpj_cpf').Value              := qryNFSE.fieldbyname('prest_cpf_cnpj').Value;
    qryDestino.ParamByName('inscricao_municipal').Value   := qryNFSE.fieldbyname('PREST_INSCRICAO_MUNICIPAL').Value;
    qryDestino.ParamByName('valor_total_nota').AsFloat    := qryNFSE.fieldbyname('total_valor_nota').AsFloat;
    qryDestino.ParamByName('valor_total_servico').AsFloat := qryNFSE.fieldbyname('total_valor_servico').AsFloat;
    qryDestino.ParamByName('valor_total_iss').AsFloat     := qryNFSE.fieldbyname('total_valor_iss').AsFloat;

    qryDestino.ParamByName('ano').Value              := qryNFSE.fieldbyname('ANO').Value;
    qryDestino.ParamByName('mes').Value              := qryNFSE.fieldbyname('MES').Value;
    qryDestino.ParamByName('nome_mes').Value         := vNomeMes;
    qryDestino.ParamByName('data').Value             := qryNFSE.fieldbyname('ANO').Asstring+'-'+
                                                        qryNFSE.fieldbyname('MES').AsString+'-01';
    qryDestino.ParamByName('competencia').Value      := qryNFSE.fieldbyname('MES_COMPETENCIA').Value;
    qryDestino.ParamByName('datahora').Value         := vDataHora;
    qryDestino.ParamByName('pessoa_sistema_siat_id').Value := TrazIDPessoa_Sist_Siat(qryNFSE.fieldbyname('prest_cpf_cnpj').AsString);



    Try
      Try
      qryDestino.ExecSQL;
      Except
//      vArquivoTexto := 'C:\SEMF\AGRUPAMENTOS.TXT';
//      GravaArquivoTexto(qryNFSE.fieldbyname('cnpj_cpf').Value);

      End;
    Finally
    End;


    qryNFSE.Next;
    end; //while...

  lblFim.Caption := 'FINAL: '+timetostr(now);
  Showmessage('FINAL');

end;

procedure TfrmImportaICMS.conecta_siscon_auxBeforeConnect(Sender: TObject);
begin
  conecta_siscon_aux.Properties.Add('cliente_enconding=latin1');
  conecta_siscon_aux.Properties.Add('codepage=latin1');

end;

procedure TfrmImportaICMS.BitBtn6Click(Sender: TObject);
begin
  BitBtn6.Enabled := False;
//  AtualizaEnderecoReceita;//20/09/2016

// Cruzamento_NFE_Cartao;

  Insere_PessoaSistemaSiat;

  showmessage('FINAL DE PESSOA SIAT.');
end;

procedure TfrmImportaICMS.Cruzamento_NFE_Cartao;
begin
//Nesta funcionalidade faço um levantamento de notafiscal_mensal com cartao_mensal, verificando o que
//tem em NF e não tem em cartão, dando inserts em cartão com valor "0.0". 31/03/2016

  qryDestino.Close;
  qryDestino.SQL.Clear;
  qryDestino.SQL.Add('select cnpj_cpf, ano, mes, inscricao_municipal, pessoa_sistema_siat_id,data from notafiscal_mensal   ');
//  qryDestino.SQL.Add('where cnpj_cpf = ''01239648000188'' and ano < 2016 ');
  qryDestino.SQL.Add('where ano = 2016 ');
  qryDestino.SQL.Add('order by ano, mes                                  ');
  qryDestino.open;

  Status.Panels[0].Text := 'Registros: '+IntToStr(qryDestino.RecordCount);
  Application.ProcessMessages;

  vCont := 0;

  while not qryDestino.eof do
    begin


    if not Verifica_CartaoMensal(qryDestino.fieldbyname('cnpj_cpf').Value,qryDestino.fieldbyname('ano').Value,qryDestino.fieldbyname('mes').Value) then
      begin
      vCont := vCont + 1;

      qryImportacao.close;
      qryImportacao.sql.Clear;
      qryImportacao.sql.add('INSERT INTO cartao_mensal(id, nome_credenciado, tipo, cnpj_mf, insc_estadual, ');
      qryImportacao.sql.add('  mes, ano, valor_operacao, pessoa_sistema_siat_id, data)                     ');
      qryImportacao.sql.add('VALUES (nextval(''cartao_mensal_id_seq''), :nome_credenciado, :tipo, :cnpj_mf, ');
      qryImportacao.sql.add(' :insc_estadual, :mes, :ano, :valor_operacao, :pessoa_sistema_siat_id, :data)  ');
      qryImportacao.ParamByName('nome_credenciado').Value := '';
      qryImportacao.ParamByName('tipo').Value             := '65';
      qryImportacao.ParamByName('cnpj_mf').Value          := qryDestino.fieldbyname('cnpj_cpf').Value;
      qryImportacao.ParamByName('insc_estadual').Value    := qryDestino.fieldbyname('inscricao_municipal').Value;
      qryImportacao.ParamByName('mes').Value              := qryDestino.fieldbyname('mes').Asinteger;
      qryImportacao.ParamByName('ano').Value              := qryDestino.fieldbyname('ano').Asinteger;
      qryImportacao.ParamByName('valor_operacao').Value   := 0;
      qryImportacao.ParamByName('pessoa_sistema_siat_id').Value   := qryDestino.fieldbyname('pessoa_sistema_siat_id').Value;
      qryImportacao.ParamByName('data').Value   := qryDestino.fieldbyname('data').Value;
      qryImportacao.ExecSQL;
      end;

    Status.Panels[1].Text := IntToStr(vCont)+' - CNPJ.: '+qryDestino.fieldbyname('cnpj_cpf').AsString;
    Application.ProcessMessages;


    qryDestino.Next;
    end;

  showmessage('Acabou...');




end;

function TfrmImportaICMS.Verifica_CartaoMensal(cnpj: string; ano,
  mes: integer): boolean;
begin
  Result := False;

  qryImportacao.Close;
  qryImportacao.SQL.Clear;
  qryImportacao.SQL.Add('select id from cartao_mensal where cnpj_mf=:cnpj_mf and ano =:ano and mes =:mes    ');
  qryImportacao.ParamByName('cnpj_mf').Value := cnpj;
  qryImportacao.ParamByName('ano').Value := ano;
  qryImportacao.ParamByName('mes').Value := mes;
  qryImportacao.open;

  if qryImportacao.RecordCount > 0 then
    Result := True;


end;

procedure TfrmImportaICMS.AtualizaEnderecoReceita;
begin

  XlsToStringGrid(strgDados,'D:\SEMF\Contribuintes_Teresina\Extracao_CPF_Contribuintes_Teresina.xlsx');

  vDataHora := now;
  vSituacao := '';

  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(x);
  Application.ProcessMessages;

  for i := 1 to strgDados.rowcount -1 do
    begin
    if trim(strgDados.cells[0,i]) <> '' then
      begin
      vCont := vCont + 1;
      vCPF  := Trim(strgDados.cells[0,i]);
      vCPF  := StringReplace( vCPF, '.'  , '' , [rfReplaceAll]);
      vCPF  := StringReplace( vCPF, '-'  , '' , [rfReplaceAll]);
      vCPF  := StringReplace( vCPF, '/'  , '' , [rfReplaceAll]);

      vCPF  := StrZeroString(vCPF,11);


      Status.Panels[1].Text := 'Registros: ' + IntToStr(vCont)+' - CPF: '+Trim(vCPF);
      Application.ProcessMessages;

    if Length(vCPF) <= 11 then
      vTipo := 'PF'
    else
      vTipo := 'PJ';

    vCEP  := Trim(strgDados.cells[12,i]);
    vCEP  := StringReplace( vCEP, '-'  , '' , [rfReplaceAll]);

    vSituacao := Trim(strgDados.cells[13,i]);

    vIDPessoa := TrazIDPessoaExterna(vCPF);

    if vIDPessoa > 0 then
      begin

      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('update pessoa_externa set situacao =:situacao  ');
      qryRegistro.sql.add('where id =:id     ');
      qryRegistro.ParamByName('id').Value       := vIDPessoa;
      qryRegistro.ParamByName('situacao').Value := vSituacao;
      qryRegistro.ExecSQL;

      qryRegistro.close;
      qryRegistro.sql.Clear;
      qryRegistro.sql.add('UPDATE endereco_receita SET tipo_logradouro=:tipo, logradouro=:logradouro, numero=:numero,     ');
      qryRegistro.sql.add('  complemento=:complemento, bairro=:bairro, cep=:cep, municipio=:municipio, uf=:uf, datahora_processamento=:dthora  ');
      qryRegistro.sql.add('where pessoa_externa_id =:id     ');
      qryRegistro.ParamByName('id').Value          := vIDPessoa;
      qryRegistro.ParamByName('tipo').Value        := Trim(strgDados.cells[5,i]);
      qryRegistro.ParamByName('logradouro').Value  := Trim(strgDados.cells[6,i]);
      qryRegistro.ParamByName('numero').Value      := Trim(strgDados.cells[7,i]);
      qryRegistro.ParamByName('complemento').Value := Trim(strgDados.cells[8,i]);
      qryRegistro.ParamByName('bairro').Value      := Trim(strgDados.cells[9,i]);
      qryRegistro.ParamByName('cep').Value         := Trim(vCEP);
      qryRegistro.ParamByName('municipio').Value   := 'Teresina';
      qryRegistro.ParamByName('uf').Value          := 'PI';
      qryRegistro.ParamByName('dthora').Value      := vDataHora;
      qryRegistro.ExecSQL;



      end// if vIDPessoa > 0 then
    else
      begin

      qryRegistro.Close;
      qryRegistro.SQL.Clear;
      qryRegistro.SQL.Add('INSERT INTO pessoa_externa(id, nome, tipo, cpf_cnpj, cepisa, receita, situacao_receita)            ');
      qryRegistro.SQL.Add('VALUES (nextval(''pessoa_externa_id_seq''), :nome, :tipo, :cpf_cnpj, :cepisa, :receita, :situacao) ');

      qryRegistro.ParamByName('nome').Value     := Trim(strgDados.cells[1,i]);
      qryRegistro.ParamByName('tipo').Value     := Trim(vTipo);
      qryRegistro.ParamByName('cpf_cnpj').Value := Trim(vCPF);
      qryRegistro.ParamByName('cepisa').Value   := False;
      qryRegistro.ParamByName('receita').Value  := True;
      qryRegistro.ParamByName('situacao').Value := vSituacao;
      qryRegistro.ExecSQL;

      //pesquisando
      qryTributo.Close;
      qryTributo.SQL.Clear;
      qryTributo.SQL.Add('select Max(id) as Ultimo from pessoa_externa');
      qryTributo.open;

      vIDPessoa := qryTributo.fieldbyname('ultimo').AsInteger;



    qryRegistro.Close;
    qryRegistro.SQL.Clear;
    qryRegistro.SQL.Add('INSERT INTO endereco_receita(id, tipo_logradouro, logradouro, numero,                       ');
    qryRegistro.SQL.Add(' complemento, bairro, cep, municipio, uf, datahora_processamento, pessoa_externa_id)        ');

    qryRegistro.SQL.Add('VALUES (nextval(''endereco_receita_id_seq''), :tipo_logradouro, :logradouro, :numero,       ');
    qryRegistro.SQL.Add(' :complemento, :bairro, :cep, :municipio, :uf, :datahora, :pessoa)                          ');

    qryRegistro.ParamByName('tipo_logradouro').Value := Trim(strgDados.cells[5,i]);
    qryRegistro.ParamByName('logradouro').Value      := Trim(strgDados.cells[6,i]);
    qryRegistro.ParamByName('numero').Value          := Trim(strgDados.cells[7,i]);
    qryRegistro.ParamByName('complemento').Value     := Trim(strgDados.cells[8,i]);
    qryRegistro.ParamByName('bairro').Value          := Trim(strgDados.cells[9,i]);
    qryRegistro.ParamByName('cep').Value             := Trim(vCEP);
    qryRegistro.ParamByName('municipio').Value       := 'Teresina';
    qryRegistro.ParamByName('uf').Value              := 'PI';
    qryRegistro.ParamByName('datahora').Value        := vDataHora;
    qryRegistro.ParamByName('pessoa').Value          := vIDPessoa;


    Try
      Try
      qryRegistro.ExecSQL;
      Except

      End;
      Finally
    End;

    end;//if vIDPessoa > 0 then

    end;
    end;// for i := 1 ...

  ShowMessage('Final da Atualização dos Endereços Receita');

end;

procedure TfrmImportaICMS.Brasil_GrauRisco;
begin
  XlsToStringGrid(strgDados,'C:\Temp\Grau de Risco.xlsx');


  vCont := 0;
  Status.Panels[0].Text := 'Total: ' + IntToStr(x);
  Application.ProcessMessages;

  for i := 1 to strgDados.rowcount -1 do
    begin
    if trim(strgDados.cells[0,i]) <> '' then
      begin
      vCont := vCont + 1;

      Status.Panels[1].Text := 'Registros: ' + Trim(strgDados.cells[0,i]);
      Application.ProcessMessages;

      qryBrasil.Close;
      qryBrasil.SQL.Clear;
      qryBrasil.SQL.Add('INSERT INTO atividade(id, codigo, descricao, grauriscomeioambiente,  ');
      qryBrasil.SQL.Add(' grauriscovigilanciasanitaria,grupoatividades)                       ');
      qryBrasil.SQL.Add('VALUES(:id, :codigo, :descricao, :grauriscomeioambiente,             ');
      qryBrasil.SQL.Add(' :grauriscovigilanciasanitaria,:grupoatividades)                     ');
      qryBrasil.ParamByName('id').Value := vCont;
      qryBrasil.ParamByName('codigo').Asstring                       := Trim(strgDados.cells[0,i]);
      qryBrasil.ParamByName('descricao').Asstring                    := Trim(strgDados.cells[1,i]);
      qryBrasil.ParamByName('grauriscomeioambiente').Asstring        := Trim(strgDados.cells[2,i]);
      qryBrasil.ParamByName('grauriscovigilanciasanitaria').Asstring := Trim(strgDados.cells[3,i]);
      qryBrasil.ParamByName('grupoatividades').Asstring              := Trim(strgDados.cells[4,i]);
      qryBrasil.ExecSQL;

      end;
    end;// for i := 1 ...

  showmessage('Final da leitura do GRID...');

end;

procedure TfrmImportaICMS.BitBtn7Click(Sender: TObject);
begin
  Brasil_GrauRisco;
end;

procedure TfrmImportaICMS.conecta_brasilBeforeConnect(Sender: TObject);
begin
  conecta_brasil.Properties.Add('cliente_enconding=latin1');
  conecta_brasil.Properties.Add('codepage=latin1');

end;

procedure TfrmImportaICMS.conecta_local_sisconBeforeConnect(
  Sender: TObject);
begin
  conecta_local_siscon.Properties.Add('cliente_enconding=latin1');
  conecta_local_siscon.Properties.Add('codepage=latin1');

end;

end.

User: ceti
DatabaseName: SIATTHE
Password: s3mf!c3T1
ServerName: 10.10.8.6
PortNumber: 1521
DriverType: thin
ServiceName: the

======================

atividade_estabelecimento
atividade_pessoa	       
cruzamento
dados_apuracao	         
estabelecimento_filial	 
pessoa	                 
receita_atividade_a	     
receita_atividade_b
receita_atividade_c




insert into cartao_mensal  

SELECT nextval('cartao_mensal_id_seq'), r65.nome_credenciado, r65.tipo, r65.cnpj_mf, r65.insc_estadual, EXTRACT(MONTH FROM r65.data) as mes, EXTRACT(YEAR FROM r65.data) as ano, 
       sum(r65.valor_operacao) as valor_operacao, r65.pessoa_sistema_siat_id
from registro10 r10 
inner join registro65 r65 on r10.id = r65.registro10_id 

where r65.cnpj_mf = '00943161000119' and EXTRACT(YEAR FROM r65.data) = 2015 and EXTRACT(MONTH FROM r65.data) = 11 and r65.totalizado = 'f'

--where EXTRACT(YEAR FROM data) = 2015 AND pessoa_sistema_siat_id = 3598

group by r65.nome_credenciado, r65.tipo, r65.cnpj_mf, r65.insc_estadual, EXTRACT(MONTH FROM r65.data), EXTRACT(YEAR FROM r65.data), 
       r65.pessoa_sistema_siat_id

