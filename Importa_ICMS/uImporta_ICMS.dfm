object frmImportaICMS: TfrmImportaICMS
  Left = 549
  Top = 54
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Importa ICMS-Server 100.100.100.203 (12-2018)'
  ClientHeight = 281
  ClientWidth = 509
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = [fsBold]
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object lblInforme: TLabel
    Left = 67
    Top = 14
    Width = 209
    Height = 13
    Caption = 'Informe o caminho dos arquivos TXT'
  end
  object lblInicio: TLabel
    Left = 24
    Top = 64
    Width = 38
    Height = 13
    Caption = 'In'#237'cio:'
  end
  object lblFim: TLabel
    Left = 24
    Top = 97
    Width = 24
    Height = 13
    Caption = 'Fim:'
  end
  object lblDestino: TLabel
    Left = 5
    Top = 209
    Width = 100
    Height = 13
    Caption = 'Caminho Destino:'
  end
  object Label1: TLabel
    Left = 152
    Top = 246
    Width = 282
    Height = 13
    Caption = 'Registro65 est'#225' sendo atualizado em 10.10.8.163'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clGreen
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Status: TStatusBar
    Left = 0
    Top = 262
    Width = 509
    Height = 19
    Panels = <
      item
        Width = 150
      end
      item
        Width = 280
      end
      item
        Width = 50
      end>
  end
  object btnConfirma: TBitBtn
    Left = 80
    Top = 168
    Width = 75
    Height = 25
    Caption = '&Confirma'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = btnConfirmaClick
  end
  object BitBtn1: TBitBtn
    Left = 307
    Top = 168
    Width = 75
    Height = 25
    Caption = '&Fechar'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = BitBtn1Click
  end
  object edtCaminho: TFilenameEdit
    Left = 66
    Top = 30
    Width = 376
    Height = 21
    Filter = 'All files (*.txt)|*.txt'
    NumGlyphs = 1
    TabOrder = 3
    Text = 'C:\SEMF\Cartao\Arquivos\'
  end
  object ListBox1: TListBox
    Left = 584
    Top = 384
    Width = 489
    Height = 305
    ItemHeight = 13
    TabOrder = 4
    Visible = False
  end
  object BitBtn2: TBitBtn
    Left = 712
    Top = 520
    Width = 201
    Height = 25
    Caption = 'Atualiza Nome Credenciado'
    Enabled = False
    TabOrder = 5
    Visible = False
    OnClick = BitBtn2Click
  end
  object BitBtn4: TBitBtn
    Left = 960
    Top = 520
    Width = 75
    Height = 25
    Caption = 'BitBtn4'
    Enabled = False
    TabOrder = 6
    Visible = False
    OnClick = BitBtn4Click
  end
  object rgOpcoes: TRadioGroup
    Left = 23
    Top = 412
    Width = 506
    Height = 155
    Caption = 'Op'#231#245'es'
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'Processa Cart'#227'o Cr'#233'dito'
      'Processa N.F. Eletr'#244'nica(2016)'
      'Agrupar Nota Fiscal'
      'Processa Registro Pagamento'
      'Pagamentos / Faixa'
      'Processa Relat'#243'rio Arrecada'#231#227'o'
      'Processa Relat'#243'rio Credito Geral'
      'Processa Endere'#231'o Receita'
      'Agrupar Cart'#227'o de Credito'
      'Processa Endere'#231'o Cepisa'
      'Insere ISS Pago'
      'Agrupa ISS Pago'
      'Processa Arrecada'#231#227'o por Grupo Local'
      'Processa Endere'#231'o Tomador'
      'Processa Rendimentos Aut'#244'nomos'
      'Atualiza NFE Aut'#244'nomos'
      'Atualiza CMC Aut'#244'nomos'
      'Processa Malha Simples Nacional')
    TabOrder = 7
    OnClick = rgOpcoesClick
  end
  object BitBtn3: TBitBtn
    Left = 736
    Top = 560
    Width = 113
    Height = 25
    Caption = 'Atualiza Pessoa'
    Enabled = False
    TabOrder = 8
    OnClick = BitBtn3Click
  end
  object BitBtn5: TBitBtn
    Left = 1168
    Top = 552
    Width = 121
    Height = 25
    Caption = 'Atualiza'#231#245'es'
    Enabled = False
    TabOrder = 9
    Visible = False
    OnClick = BitBtn5Click
  end
  object Edit1: TEdit
    Left = 26
    Top = 118
    Width = 455
    Height = 21
    TabOrder = 10
  end
  object strgDados: TStringGrid
    Left = 600
    Top = 40
    Width = 657
    Height = 289
    ColCount = 2
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    ParentFont = False
    TabOrder = 11
    ColWidths = (
      82
      109)
  end
  object BitBtn6: TBitBtn
    Left = 232
    Top = 376
    Width = 209
    Height = 25
    Caption = 'Insere/Atualiza Pessoa SIAT'
    Enabled = False
    TabOrder = 12
    OnClick = BitBtn6Click
  end
  object BitBtn7: TBitBtn
    Left = 448
    Top = 328
    Width = 75
    Height = 25
    Caption = 'BitBtn7'
    Enabled = False
    TabOrder = 13
    OnClick = BitBtn7Click
  end
  object conecta_siscon_aux: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '100.100.100.203'
    Port = 5432
    Database = 'siscon_aux'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = conecta_siscon_auxBeforeConnect
    Left = 32
    Top = 48
  end
  object qryDestino: TZQuery
    Connection = conecta_siscon_aux
    Params = <>
    Left = 341
    Top = 65
  end
  object qryImportacao: TZQuery
    Connection = conecta_siscon_aux
    Params = <>
    Left = 93
    Top = 57
  end
  object Conecta_SIAT: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=d$f123;User ID=ceti;Data Source="(DE' +
      'SCRIPTION=(ADDRESS=(PROTOCOL = TCP)(HOST = 10.10.8.6)(PORT = 152' +
      '1))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = THE)))";Pe' +
      'rsist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 648
    Top = 96
  end
  object qryBuscaSIAT: TADOQuery
    Connection = Conecta_SIAT
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      'select * from SIATTHE.tblpes where cpfcnpj = '#39'06626253019090'#39)
    Left = 648
    Top = 144
  end
  object qryNome: TZQuery
    Connection = conecta_siscon_aux
    Params = <>
    Left = 197
    Top = 57
  end
  object qryVerifica: TZQuery
    Connection = conecta_siscon_aux
    Params = <>
    Left = 261
    Top = 57
  end
  object con_siscon: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '100.100.100.203'
    Port = 5432
    Database = 'siscon'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = con_sisconBeforeConnect
    Left = 832
    Top = 368
  end
  object qryTributo: TZQuery
    Connection = con_siscon
    Params = <>
    Left = 973
    Top = 369
  end
  object qryRegistro: TZQuery
    Connection = con_siscon
    Params = <>
    Left = 893
    Top = 425
  end
  object ppDBPipeline1: TppDBPipeline
    DataSource = DataSource1
    UserName = 'DBPipeline1'
    Left = 1184
    Top = 432
    object ppDBPipeline1ppField1: TppField
      FieldAlias = 'tributo'
      FieldName = 'tributo'
      FieldLength = 0
      DataType = dtMemo
      DisplayWidth = 0
      Position = 0
      Searchable = False
      Sortable = False
    end
    object ppDBPipeline1ppField2: TppField
      Alignment = taRightJustify
      FieldAlias = 'ano'
      FieldName = 'ano'
      FieldLength = 0
      DataType = dtInteger
      DisplayWidth = 10
      Position = 1
    end
    object ppDBPipeline1ppField3: TppField
      Alignment = taRightJustify
      FieldAlias = 'janeiro'
      FieldName = 'janeiro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 2
    end
    object ppDBPipeline1ppField4: TppField
      Alignment = taRightJustify
      FieldAlias = 'fevereiro'
      FieldName = 'fevereiro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 3
    end
    object ppDBPipeline1ppField5: TppField
      Alignment = taRightJustify
      FieldAlias = 'marco'
      FieldName = 'marco'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 4
    end
    object ppDBPipeline1ppField6: TppField
      Alignment = taRightJustify
      FieldAlias = 'abril'
      FieldName = 'abril'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 5
    end
    object ppDBPipeline1ppField7: TppField
      Alignment = taRightJustify
      FieldAlias = 'maio'
      FieldName = 'maio'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 6
    end
    object ppDBPipeline1ppField8: TppField
      Alignment = taRightJustify
      FieldAlias = 'junho'
      FieldName = 'junho'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 7
    end
    object ppDBPipeline1ppField9: TppField
      Alignment = taRightJustify
      FieldAlias = 'julho'
      FieldName = 'julho'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 8
    end
    object ppDBPipeline1ppField10: TppField
      Alignment = taRightJustify
      FieldAlias = 'agosto'
      FieldName = 'agosto'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 9
    end
    object ppDBPipeline1ppField11: TppField
      Alignment = taRightJustify
      FieldAlias = 'setembro'
      FieldName = 'setembro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 10
    end
    object ppDBPipeline1ppField12: TppField
      Alignment = taRightJustify
      FieldAlias = 'outubro'
      FieldName = 'outubro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 11
    end
    object ppDBPipeline1ppField13: TppField
      Alignment = taRightJustify
      FieldAlias = 'novembro'
      FieldName = 'novembro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 12
    end
    object ppDBPipeline1ppField14: TppField
      Alignment = taRightJustify
      FieldAlias = 'dezembro'
      FieldName = 'dezembro'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 13
    end
    object ppDBPipeline1ppField15: TppField
      Alignment = taRightJustify
      FieldAlias = 'total'
      FieldName = 'total'
      FieldLength = 0
      DataType = dtDouble
      DisplayWidth = 10
      Position = 14
    end
  end
  object ppReport1: TppReport
    AutoStop = False
    DataPipeline = ppDBPipeline1
    PrinterSetup.BinName = 'Default'
    PrinterSetup.DocumentName = 'Report'
    PrinterSetup.Orientation = poLandscape
    PrinterSetup.PaperName = 'A4'
    PrinterSetup.PrinterName = 'Default'
    PrinterSetup.mmMarginBottom = 6350
    PrinterSetup.mmMarginLeft = 6350
    PrinterSetup.mmMarginRight = 6350
    PrinterSetup.mmMarginTop = 6350
    PrinterSetup.mmPaperHeight = 210000
    PrinterSetup.mmPaperWidth = 297000
    PrinterSetup.PaperSize = 9
    DeviceType = 'Screen'
    EmailSettings.ReportFormat = 'PDF'
    OutlineSettings.CreateNode = True
    OutlineSettings.CreatePageNodes = True
    OutlineSettings.Enabled = True
    OutlineSettings.Visible = True
    TextSearchSettings.DefaultString = '<FindText>'
    TextSearchSettings.Enabled = True
    Left = 208
    Top = 296
    Version = '10.02'
    mmColumnWidth = 0
    DataPipelineName = 'ppDBPipeline1'
    object ppHeaderBand1: TppHeaderBand
      mmBottomOffset = 0
      mmHeight = 13229
      mmPrintPosition = 0
    end
    object ppDetailBand1: TppDetailBand
      mmBottomOffset = 0
      mmHeight = 15346
      mmPrintPosition = 0
      object ppDBText1: TppDBText
        UserName = 'DBText1'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'janeiro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 19844
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText2: TppDBText
        UserName = 'DBText2'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'fevereiro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 38894
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText3: TppDBText
        UserName = 'DBText3'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'marco'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 59002
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText4: TppDBText
        UserName = 'DBText4'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'abril'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 79904
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText5: TppDBText
        UserName = 'DBText5'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'maio'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 103452
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText6: TppDBText
        UserName = 'DBText6'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'junho'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 123825
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText7: TppDBText
        UserName = 'DBText7'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'ano'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 6879
        mmTop = 1059
        mmWidth = 12171
        BandType = 4
      end
      object ppDBText8: TppDBText
        UserName = 'DBText8'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'julho'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 145786
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText9: TppDBText
        UserName = 'DBText9'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'agosto'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 164836
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText10: TppDBText
        UserName = 'DBText10'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'setembro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 184415
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText11: TppDBText
        UserName = 'DBText11'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'outubro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 203465
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText12: TppDBText
        UserName = 'DBText12'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'novembro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 223309
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText13: TppDBText
        UserName = 'DBText13'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'dezembro'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 243946
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
      object ppDBText14: TppDBText
        UserName = 'DBText14'
        Border.BorderPositions = []
        Border.Color = clBlack
        Border.Style = psSolid
        Border.Visible = False
        Border.Weight = 1.000000000000000000
        DataField = 'total'
        DataPipeline = ppDBPipeline1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Name = 'Arial'
        Font.Size = 9
        Font.Style = []
        TextAlignment = taRightJustified
        Transparent = True
        DataPipelineName = 'ppDBPipeline1'
        mmHeight = 3598
        mmLeft = 262996
        mmTop = 1059
        mmWidth = 17727
        BandType = 4
      end
    end
    object ppFooterBand1: TppFooterBand
      mmBottomOffset = 0
      mmHeight = 13229
      mmPrintPosition = 0
    end
    object ppGroup1: TppGroup
      BreakName = 'tributo'
      DataPipeline = ppDBPipeline1
      KeepTogether = True
      OutlineSettings.CreateNode = True
      UserName = 'Group1'
      mmNewColumnThreshold = 0
      mmNewPageThreshold = 0
      DataPipelineName = 'ppDBPipeline1'
      object ppGroupHeaderBand1: TppGroupHeaderBand
        mmBottomOffset = 0
        mmHeight = 15875
        mmPrintPosition = 0
        object ppLabel1: TppLabel
          UserName = 'Label1'
          AutoSize = False
          Border.BorderPositions = []
          Border.Color = clBlack
          Border.Style = psSolid
          Border.Visible = False
          Border.Weight = 1.000000000000000000
          Caption = 'tributo'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Name = 'Arial'
          Font.Size = 12
          Font.Style = []
          Transparent = True
          mmHeight = 4763
          mmLeft = 3440
          mmTop = 529
          mmWidth = 12435
          BandType = 3
          GroupNo = 0
        end
        object ppDBMemo1: TppDBMemo
          UserName = 'DBMemo1'
          Border.BorderPositions = []
          Border.Color = clBlack
          Border.Style = psSolid
          Border.Visible = False
          Border.Weight = 1.000000000000000000
          CharWrap = False
          DataField = 'tributo'
          DataPipeline = ppDBPipeline1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Name = 'Arial'
          Font.Size = 12
          Font.Style = []
          Transparent = True
          DataPipelineName = 'ppDBPipeline1'
          mmHeight = 5821
          mmLeft = 16669
          mmTop = 265
          mmWidth = 62177
          BandType = 3
          GroupNo = 0
          mmBottomOffset = 0
          mmOverFlowOffset = 0
          mmStopPosition = 0
          mmLeading = 0
        end
        object ppLabel2: TppLabel
          UserName = 'Label2'
          AutoSize = False
          Border.BorderPositions = []
          Border.Color = clBlack
          Border.Style = psSolid
          Border.Visible = False
          Border.Weight = 1.000000000000000000
          Caption = 'janeiro'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Name = 'Arial'
          Font.Size = 12
          Font.Style = []
          Transparent = True
          mmHeight = 4763
          mmLeft = 20638
          mmTop = 10848
          mmWidth = 13494
          BandType = 3
          GroupNo = 0
        end
      end
      object ppGroupFooterBand1: TppGroupFooterBand
        mmBottomOffset = 0
        mmHeight = 16404
        mmPrintPosition = 0
        object ppDBCalc1: TppDBCalc
          UserName = 'DBCalc1'
          Border.BorderPositions = []
          Border.Color = clBlack
          Border.Style = psSolid
          Border.Visible = False
          Border.Weight = 1.000000000000000000
          DataField = 'janeiro'
          DataPipeline = ppDBPipeline1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Name = 'Arial'
          Font.Size = 9
          Font.Style = []
          ResetGroup = ppGroup1
          TextAlignment = taRightJustified
          Transparent = True
          DBCalcType = dcMaximum
          DataPipelineName = 'ppDBPipeline1'
          mmHeight = 3704
          mmLeft = 5027
          mmTop = 1588
          mmWidth = 33338
          BandType = 5
          GroupNo = 0
        end
        object ppDBCalc2: TppDBCalc
          UserName = 'DBCalc2'
          Border.BorderPositions = []
          Border.Color = clBlack
          Border.Style = psSolid
          Border.Visible = False
          Border.Weight = 1.000000000000000000
          DataField = 'janeiro'
          DataPipeline = ppDBPipeline1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Name = 'Arial'
          Font.Size = 9
          Font.Style = []
          ResetGroup = ppGroup1
          TextAlignment = taRightJustified
          Transparent = True
          DBCalcType = dcMinimum
          DataPipelineName = 'ppDBPipeline1'
          mmHeight = 3704
          mmLeft = 39688
          mmTop = 1588
          mmWidth = 47361
          BandType = 5
          GroupNo = 0
        end
      end
    end
  end
  object DataSource1: TDataSource
    DataSet = ZQuery1
    Left = 1208
    Top = 376
  end
  object ZQuery1: TZQuery
    Connection = con_siscon
    SQL.Strings = (
      'select tributo.codtrb||'#39'-'#39'||tributo.desmin as tributo,ano,'
      'SUM(CASE WHEN mes = 1 THEN valor_pago ELSE 0  END) as JANEIRO,'
      'SUM(CASE WHEN mes = 2 THEN valor_pago ELSE 0  END) as FEVEREIRO,'
      'SUM(CASE WHEN mes = 3 THEN valor_pago ELSE 0  END) as MARCO,'
      'SUM(CASE WHEN mes = 4 THEN valor_pago ELSE 0  END) as ABRIL,'
      'SUM(CASE WHEN mes = 5 THEN valor_pago ELSE 0  END) as MAIO,'
      'SUM(CASE WHEN mes = 6 THEN valor_pago ELSE 0  END) as JUNHO,'
      'SUM(CASE WHEN mes = 7 THEN valor_pago ELSE 0  END) as JULHO,'
      'SUM(CASE WHEN mes = 8 THEN valor_pago ELSE 0  END) as AGOSTO,'
      'SUM(CASE WHEN mes = 9 THEN valor_pago ELSE 0  END) as SETEMBRO,'
      'SUM(CASE WHEN mes = 10 THEN valor_pago ELSE 0  END) as OUTUBRO,'
      'SUM(CASE WHEN mes = 11 THEN valor_pago ELSE 0  END) as NOVEMBRO,'
      'SUM(CASE WHEN mes = 12 THEN valor_pago ELSE 0  END) as DEZEMBRO,'
      'SUM(valor_pago) AS TOTAL'
      ''
      'from registro_pagamento, tributo '
      'where tributo_id in (103,104) and ano=2011'
      '  and tributo_id = tributo.id'
      'group by tributo.codtrb||'#39'-'#39'||tributo.desmin, ano '
      ''
      'UNION'
      ''
      'select tributo.codtrb||'#39'-'#39'||tributo.desmin as tributo,ano,'
      'SUM(CASE WHEN mes = 1 THEN valor_pago ELSE 0 END) as JANEIRO,'
      'SUM(CASE WHEN mes = 2 THEN valor_pago ELSE 0  END) as FEVEREIRO,'
      'SUM(CASE WHEN mes = 3 THEN valor_pago ELSE 0  END) as MARCO,'
      'SUM(CASE WHEN mes = 4 THEN valor_pago ELSE 0  END) as ABRIL,'
      'SUM(CASE WHEN mes = 5 THEN valor_pago ELSE 0  END) as MAIO,'
      'SUM(CASE WHEN mes = 6 THEN valor_pago ELSE 0  END) as JUNHO,'
      'SUM(CASE WHEN mes = 7 THEN valor_pago ELSE 0  END) as JULHO,'
      'SUM(CASE WHEN mes = 8 THEN valor_pago ELSE 0  END) as AGOSTO,'
      'SUM(CASE WHEN mes = 9 THEN valor_pago ELSE 0  END) as SETEMBRO,'
      'SUM(CASE WHEN mes = 10 THEN valor_pago ELSE 0  END) as OUTUBRO,'
      'SUM(CASE WHEN mes = 11 THEN valor_pago ELSE 0  END) as NOVEMBRO,'
      'SUM(CASE WHEN mes = 12 THEN valor_pago ELSE 0  END) as DEZEMBRO,'
      'SUM(valor_pago) AS TOTAL'
      ''
      'from registro_pagamento, tributo'
      'where tributo_id in (103,104) and ano=2012'
      '  and tributo_id = tributo.id'
      'group by tributo.codtrb||'#39'-'#39'||tributo.desmin, ano '
      'order by tributo, ano'
      '')
    Params = <>
    Left = 1149
    Top = 377
  end
  object qryDados: TADOQuery
    Connection = Conecta_SIAT
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      'select * from SIATTHE.tblpes where cpfcnpj = '#39'06626253019090'#39)
    Left = 752
    Top = 96
  end
  object con_gsrf: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '100.100.100.203'
    Port = 5432
    Database = 'gsrf'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = con_gsrfBeforeConnect
    Left = 456
    Top = 560
  end
  object qryGSRF: TZQuery
    Connection = con_gsrf
    Params = <>
    Left = 517
    Top = 561
  end
  object qryGSRF2: TZQuery
    Connection = con_gsrf
    Params = <>
    Left = 517
    Top = 617
  end
  object qryGSRF3: TZQuery
    Connection = con_gsrf
    Params = <>
    Left = 445
    Top = 617
  end
  object qryTomador: TZQuery
    Connection = con_siscon
    Params = <>
    Left = 965
    Top = 433
  end
  object qryAutonomo: TZQuery
    Connection = con_siscon
    Params = <>
    Left = 917
    Top = 361
  end
  object Conecta_NFSE: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=d$f2014$%;User ID=CONSULTA;Data Sour' +
      'ce="(DESCRIPTION=(ADDRESS=(PROTOCOL = TCP)(HOST = 10.10.8.10)(PO' +
      'RT = 1522))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = NF' +
      'SETHEP)))";Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 856
    Top = 184
  end
  object qryNFSE: TADOQuery
    Connection = Conecta_NFSE
    Parameters = <>
    Prepared = True
    Left = 856
    Top = 248
  end
  object conecta_150: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '10.10.8.150'
    Port = 5432
    Database = 'teste'
    User = 'postgres'
    Password = '123456'
    BeforeConnect = conecta_150BeforeConnect
    Left = 232
    Top = 464
  end
  object qryTeste: TZQuery
    Connection = conecta_150
    Params = <>
    Left = 293
    Top = 465
  end
  object conecta_simplesnacional: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '100.100.100.203'
    Port = 5432
    Database = 'simples_nacional'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = con_sisconBeforeConnect
    Left = 1152
    Top = 8
  end
  object qryApuracao: TZQuery
    Connection = conecta_simplesnacional
    Params = <>
    Left = 1157
    Top = 113
  end
  object qryPessoaSN: TZQuery
    Connection = conecta_simplesnacional
    Params = <>
    Left = 1157
    Top = 57
  end
  object conecta_brasil: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '100.100.100.203'
    Port = 5432
    Database = 'smaj8'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = conecta_brasilBeforeConnect
    Left = 96
    Top = 520
  end
  object qryBrasil: TZQuery
    Connection = conecta_brasil
    Params = <>
    Left = 93
    Top = 569
  end
  object conecta_local_siscon: TZConnection
    Protocol = 'postgresql-7.4'
    HostName = '10.10.8.163'
    Port = 5432
    Database = 'siscon_aux'
    User = 'postgres'
    Password = 'sysadm'
    BeforeConnect = conecta_local_sisconBeforeConnect
    Left = 416
    Top = 144
  end
  object qryRegistro65_local: TZQuery
    Connection = conecta_local_siscon
    Params = <>
    Left = 477
    Top = 65
  end
end
