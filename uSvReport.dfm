object svReport: TsvReport
  DisplayName = 'svReport'
  AfterInstall = ServiceAfterInstall
  AfterUninstall = ServiceAfterUninstall
  OnExecute = ServiceExecute
  OnStart = ServiceStart
  Height = 900
  Width = 1200
  PixelsPerInch = 144
  object con1: TMyConnection
    Database = 'sevenths_web_manage2'
    Port = 3838
    Username = 'Fatra'
    Server = 'abdi-svr02.ddns.net'
    Left = 19
    Top = 19
    EncryptedPassword = 'C8FFCCFF99FF9EFF91FF98FF99FF9EFF91FF98FF'
  end
  object qReportWeb: TwwMyQuery
    SQLUpdate.Strings = (
      'UPDATE API_REPORT'
      'SET'
      
        '  GENERATE_FLAG = :GENERATE_FLAG, ERROR = :ERROR, ERROR_USER = :' +
        'ERROR_USER, OUTPUT_FILE = :OUTPUT_FILE, PATH_FILE = :PATH_FILE, ' +
        'PDF_FLAG = :PDF_FLAG'
      'WHERE'
      '  API_REPORT_ID = :Old_API_REPORT_ID')
    Connection = con1
    SQL.Strings = (
      'SELECT '
      'a.*'
      'FROM API_REPORT a'
      'WHERE GENERATE_FLAG = 0 AND IFNULL(ERROR,'#39#39') = '#39#39
      '&KONDISI'
      'ORDER BY RAND()'
      'LIMIT 50')
    Left = 96
    Top = 19
    MacroData = <
      item
        Name = 'KONDISI'
      end>
    object qReportWebAPI_REPORT_ID: TIntegerField
      FieldName = 'API_REPORT_ID'
      Origin = 'API_REPORT.API_REPORT_ID'
    end
    object qReportWebJSON: TMemoField
      FieldName = 'JSON'
      Origin = 'API_REPORT.JSON'
      BlobType = ftMemo
    end
    object qReportWebSERVER_ID: TIntegerField
      FieldName = 'SERVER_ID'
      Origin = 'API_REPORT.SERVER_ID'
    end
    object qReportWebERP_RPT_ID: TIntegerField
      FieldName = 'ERP_RPT_ID'
      Origin = 'API_REPORT.ERP_RPT_ID'
    end
    object qReportWebGENERATE_FLAG: TIntegerField
      FieldName = 'GENERATE_FLAG'
      Origin = 'API_REPORT.GENERATE_FLAG'
    end
    object qReportWebPDF_FLAG: TIntegerField
      FieldName = 'PDF_FLAG'
      Origin = 'API_REPORT.PDF_FLAG'
    end
    object qReportWebREAD_FLAG: TBooleanField
      FieldName = 'READ_FLAG'
      Origin = 'API_REPORT.READ_FLAG'
    end
    object qReportWebNOTIF_FLAG: TBooleanField
      FieldName = 'NOTIF_FLAG'
      Origin = 'API_REPORT.NOTIF_FLAG'
    end
    object qReportWebERROR: TMemoField
      FieldName = 'ERROR'
      Origin = 'API_REPORT.ERROR'
      BlobType = ftMemo
    end
    object qReportWebERROR_USER: TMemoField
      FieldName = 'ERROR_USER'
      Origin = 'API_REPORT.ERROR_USER'
      BlobType = ftMemo
    end
    object qReportWebPATH_FILE: TStringField
      FieldName = 'PATH_FILE'
      Origin = 'API_REPORT.PATH_FILE'
      Size = 500
    end
    object qReportWebOUTPUT_FILE: TStringField
      FieldName = 'OUTPUT_FILE'
      Origin = 'API_REPORT.OUTPUT_FILE'
      Size = 500
    end
    object qReportWebCREATED_DATE: TDateTimeField
      FieldName = 'CREATED_DATE'
      Origin = 'API_REPORT.CREATED_DATE'
    end
    object qReportWebLAST_UPDATED_DATE: TDateTimeField
      FieldName = 'LAST_UPDATED_DATE'
      Origin = 'API_REPORT.LAST_UPDATED_DATE'
    end
    object qReportWebERP_USER_ID: TIntegerField
      FieldName = 'ERP_USER_ID'
      Origin = 'API_REPORT.ERP_USER_ID'
    end
    object qReportWebHOSTNAME: TStringField
      FieldName = 'HOSTNAME'
      Origin = 'API_REPORT.HOSTNAME'
      Size = 100
    end
    object qReportWebPORT: TStringField
      FieldName = 'PORT'
      Origin = 'API_REPORT.PORT'
      Size = 50
    end
    object qReportWebDATABASE_NAME: TStringField
      FieldName = 'DATABASE_NAME'
      Origin = 'API_REPORT.DATABASE_NAME'
      Size = 60
    end
    object qReportWebTEMPLATE_FILE: TStringField
      FieldName = 'TEMPLATE_FILE'
      Origin = 'API_REPORT.TEMPLATE_FILE'
      Size = 100
    end
  end
  object qErpDetail: TwwMyQuery
    Connection = con1
    SQL.Strings = (
      'SELECT * FROM('
      'SELECT'
      #39#39' ERP_DETAIL_ID,'
      #39#39' WHERE_COLUMN,'
      #39#39' WHERE_VALUE'
      ')parameter'
      'WHERE WHERE_COLUMN <> '#39#39)
    CachedUpdates = True
    AfterScroll = qErpDetailAfterScroll
    Left = 192
    Top = 19
    object qErpDetailERP_DETAIL_ID: TStringField
      FieldName = 'ERP_DETAIL_ID'
      FixedChar = True
      Size = 65530
    end
    object qErpDetailWHERE_COLUMN: TStringField
      FieldName = 'WHERE_COLUMN'
      FixedChar = True
      Size = 65530
    end
    object qErpDetailWHERE_VALUE: TStringField
      FieldName = 'WHERE_VALUE'
      FixedChar = True
      Size = 65530
    end
  end
  object Query2: TwwMyQuery
    Connection = con1
    Left = 278
    Top = 19
  end
  object qDB: TwwMyQuery
    Connection = con1
    SQL.Strings = (
      'SELECT schema_name FROM information_schema.schemata'
      
        'WHERE schema_name <> '#39'information_schema'#39' AND  schema_name <> '#39'm' +
        'ysql'#39' and schema_name <> '#39'performance_schema'#39
      '&KONDISI')
    Left = 365
    Top = 19
    MacroData = <
      item
        Name = 'KONDISI'
      end>
    object qDBschema_name: TStringField
      FieldName = 'schema_name'
      Size = 64
    end
  end
  object qReportWebDel: TwwMyQuery
    Connection = con1
    SQL.Strings = (
      'SELECT '
      'a.*'
      'FROM API_REPORT a'
      'WHERE GENERATE_FLAG = 0 AND IFNULL(ERROR,'#39#39') = '#39#39
      '&KONDISI'
      'LIMIT 50')
    Left = 442
    Top = 19
    MacroData = <
      item
        Name = 'KONDISI'
      end>
    object qReportWebDelAPI_REPORT_ID: TIntegerField
      FieldName = 'API_REPORT_ID'
    end
    object qReportWebDelJSON: TMemoField
      FieldName = 'JSON'
      BlobType = ftMemo
    end
    object qReportWebDelSERVER_ID: TIntegerField
      FieldName = 'SERVER_ID'
    end
    object qReportWebDelERP_RPT_ID: TIntegerField
      FieldName = 'ERP_RPT_ID'
    end
    object qReportWebDelGENERATE_FLAG: TIntegerField
      FieldName = 'GENERATE_FLAG'
    end
    object qReportWebDelPDF_FLAG: TIntegerField
      FieldName = 'PDF_FLAG'
    end
    object qReportWebDelREAD_FLAG: TBooleanField
      FieldName = 'READ_FLAG'
    end
    object qReportWebDelNOTIF_FLAG: TBooleanField
      FieldName = 'NOTIF_FLAG'
    end
    object qReportWebDelERROR: TMemoField
      FieldName = 'ERROR'
      BlobType = ftMemo
    end
    object qReportWebDelERROR_USER: TMemoField
      FieldName = 'ERROR_USER'
      BlobType = ftMemo
    end
    object qReportWebDelPATH_FILE: TStringField
      FieldName = 'PATH_FILE'
      Size = 500
    end
    object qReportWebDelOUTPUT_FILE: TStringField
      FieldName = 'OUTPUT_FILE'
      Size = 500
    end
    object qReportWebDelCREATED_DATE: TDateTimeField
      FieldName = 'CREATED_DATE'
    end
    object qReportWebDelLAST_UPDATED_DATE: TDateTimeField
      FieldName = 'LAST_UPDATED_DATE'
    end
    object qReportWebDelERP_USER_ID: TIntegerField
      FieldName = 'ERP_USER_ID'
    end
    object qReportWebDelHOSTNAME: TStringField
      FieldName = 'HOSTNAME'
      Size = 100
    end
    object qReportWebDelPORT: TStringField
      FieldName = 'PORT'
      Size = 50
    end
    object qReportWebDelDATABASE_NAME: TStringField
      FieldName = 'DATABASE_NAME'
      Size = 60
    end
    object qReportWebDelTEMPLATE_FILE: TStringField
      FieldName = 'TEMPLATE_FILE'
      Size = 100
    end
  end
  object db1: TMyConnection
    Username = 'Fatra'
    AfterConnect = db1AfterConnect
    Left = 19
    Top = 125
    EncryptedPassword = 'C8FFCCFF99FF9EFF91FF98FF99FF9EFF91FF98FF'
  end
  object Query1: TwwMyQuery
    Connection = db1
    Left = 86
    Top = 125
  end
  object sql1: TMyScript
    Debug = True
    OnError = sql1Error
    Connection = db1
    Left = 163
    Top = 125
  end
  object wwMyQuery1: TwwMyQuery
    Connection = db1
    Left = 1075
    Top = 115
  end
  object qClient: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '   &select'
      'FROM'
      '  &From'
      'WHERE'
      '  a.active_flag = '#39'Y'#39'  '
      'ORDER BY'
      '  &Order'
      '')
    Left = 365
    Top = 125
    MacroData = <
      item
        Name = 'select'
      end
      item
        Name = 'From'
      end
      item
        Name = 'Order'
      end>
  end
  object qNo_Doc: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '  a.DOCUMENT_NO,'
      '  a.DOCUMENT_DATE,'
      '  a.NOTE'
      'FROM'
      '  &KONDISI'
      'order by a.DOCUMENT_NO')
    Left = 442
    Top = 125
    MacroData = <
      item
        Name = 'KONDISI'
      end>
  end
  object qUser: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT * '
      'FROM ERP_USER'
      'WHERE ERP_USER_ID=:ERP_USER_ID')
    Left = 643
    Top = 58
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ERP_USER_ID'
        Value = nil
      end>
  end
  object qERP_SQL: TwwMyQuery
    Connection = db1
    Left = 730
    Top = 58
  end
  object q1: TwwMyQuery
    Connection = db1
    Left = 19
    Top = 230
  end
  object ds1: TwwDataSource
    DataSet = q1
    Left = 19
    Top = 307
  end
  object q2: TwwMyQuery
    Connection = db1
    Left = 70
    Top = 230
  end
  object ds2: TwwDataSource
    DataSet = q2
    Left = 70
    Top = 307
  end
  object q3: TwwMyQuery
    Connection = db1
    Left = 118
    Top = 230
  end
  object ds3: TwwDataSource
    DataSet = q3
    Left = 118
    Top = 307
  end
  object q4: TwwMyQuery
    Connection = db1
    Left = 166
    Top = 230
  end
  object ds4: TwwDataSource
    DataSet = q4
    Left = 166
    Top = 307
  end
  object q5: TwwMyQuery
    Connection = db1
    Left = 214
    Top = 230
  end
  object ds5: TwwDataSource
    DataSet = q5
    Left = 214
    Top = 307
  end
  object q6: TwwMyQuery
    Connection = db1
    Left = 262
    Top = 230
  end
  object ds6: TwwDataSource
    DataSet = q6
    Left = 262
    Top = 307
  end
  object q7: TwwMyQuery
    Connection = db1
    Left = 310
    Top = 230
  end
  object ds7: TwwDataSource
    DataSet = q7
    Left = 310
    Top = 307
  end
  object q8: TwwMyQuery
    Connection = db1
    Left = 358
    Top = 230
  end
  object ds8: TwwDataSource
    DataSet = q8
    Left = 358
    Top = 307
  end
  object q9: TwwMyQuery
    Connection = db1
    Left = 406
    Top = 230
  end
  object ds9: TwwDataSource
    DataSet = q9
    Left = 406
    Top = 307
  end
  object q10: TwwMyQuery
    Connection = db1
    Left = 463
    Top = 230
  end
  object ds10: TwwDataSource
    DataSet = q10
    Left = 463
    Top = 307
  end
  object q11: TwwMyQuery
    Connection = db1
    Left = 19
    Top = 394
  end
  object ds11: TwwDataSource
    DataSet = q11
    Left = 19
    Top = 470
  end
  object q12: TwwMyQuery
    Connection = db1
    Left = 70
    Top = 394
  end
  object ds12: TwwDataSource
    DataSet = q12
    Left = 70
    Top = 470
  end
  object q13: TwwMyQuery
    Connection = db1
    Left = 118
    Top = 394
  end
  object ds13: TwwDataSource
    DataSet = q13
    Left = 118
    Top = 470
  end
  object q14: TwwMyQuery
    Connection = db1
    Left = 166
    Top = 394
  end
  object ds14: TwwDataSource
    DataSet = q14
    Left = 166
    Top = 470
  end
  object q15: TwwMyQuery
    Connection = db1
    Left = 214
    Top = 394
  end
  object ds15: TwwDataSource
    DataSet = q15
    Left = 214
    Top = 470
  end
  object q16: TwwMyQuery
    Connection = db1
    Left = 262
    Top = 394
  end
  object ds16: TwwDataSource
    DataSet = q16
    Left = 262
    Top = 470
  end
  object q17: TwwMyQuery
    Connection = db1
    Left = 310
    Top = 394
  end
  object ds17: TwwDataSource
    DataSet = q17
    Left = 310
    Top = 470
  end
  object q18: TwwMyQuery
    Connection = db1
    Left = 358
    Top = 394
  end
  object ds18: TwwDataSource
    DataSet = q18
    Left = 358
    Top = 470
  end
  object q19: TwwMyQuery
    Connection = db1
    Left = 406
    Top = 394
  end
  object ds19: TwwDataSource
    DataSet = q19
    Left = 406
    Top = 470
  end
  object q20: TwwMyQuery
    Connection = db1
    Left = 463
    Top = 394
  end
  object ds20: TwwDataSource
    DataSet = q20
    Left = 463
    Top = 470
  end
  object qHeader: TwwMyQuery
    Connection = db1
    Left = 636
    Top = 144
  end
  object dsHeader: TwwDataSource
    DataSet = qHeader
    Left = 636
    Top = 221
  end
  object qDetail: TwwMyQuery
    Connection = db1
    Left = 684
    Top = 144
  end
  object dsDetail: TwwDataSource
    DataSet = qDetail
    Left = 684
    Top = 221
  end
  object qDetail1: TwwMyQuery
    Connection = db1
    Left = 732
    Top = 144
  end
  object dsDetail1: TwwDataSource
    DataSet = qDetail1
    Left = 732
    Top = 221
  end
  object qDetail2: TwwMyQuery
    Connection = db1
    Left = 780
    Top = 144
  end
  object dsDetail2: TwwDataSource
    DataSet = qDetail2
    Left = 780
    Top = 221
  end
  object qDetail3: TwwMyQuery
    Connection = db1
    Left = 838
    Top = 144
  end
  object dsDetail3: TwwDataSource
    DataSet = qDetail3
    Left = 838
    Top = 221
  end
  object qDetail4: TwwMyQuery
    Connection = db1
    Left = 895
    Top = 144
  end
  object dsDetail4: TwwDataSource
    DataSet = qDetail4
    Left = 895
    Top = 221
  end
  object qNo_Form: TwwMyQuery
    Connection = db1
    Left = 636
    Top = 307
  end
  object dsNo_Form: TwwDataSource
    DataSet = qNo_Form
    Left = 636
    Top = 384
  end
  object qSql: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT a.*, b.*, b.master_source'
      'FROM'
      '  erp_rpt a'
      '  INNER JOIN ERP_RPT_SQL b ON (a.ERP_RPT_ID = b.ERP_RPT_ID)'
      'WHERE'
      '  a.ERP_RPT_ID = :ERP_RPT_ID'
      'ORDER BY'
      '  b.no_urut')
    Left = 684
    Top = 307
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ERP_RPT_ID'
        Value = nil
      end>
  end
  object dsSql: TwwDataSource
    DataSet = qSql
    Left = 684
    Top = 384
  end
  object qGroup: TwwMyQuery
    Connection = db1
    Left = 732
    Top = 307
  end
  object dsGroup: TwwDataSource
    DataSet = qGroup
    Left = 732
    Top = 384
  end
  object qErp_rpt_sql_name: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '* FROM erp_rpt_sql_name'
      'WHERE SQL_NAME <> '#39'Tidak Ada'#39
      'ORDER BY SEQ')
    Left = 624
    Top = 480
  end
  object qSetup: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT DISTINCT *'
      'FROM'
      '  setup a'
      'left join Address b on a.address_id=b.address_id'
      '  ')
    Left = 29
    Top = 576
  end
  object QListEditor: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      'b.ERP_RPT_DETAIL_ID,'
      'l.*,'
      'b.PARAM_TABLE1, '
      'URUT'
      'FROM ERP_RPT a '
      'INNER JOIN ERP_RPT_DETAIL b ON (a.ERP_RPT_ID = b.ERP_RPT_ID)'
      
        'INNER JOIN ERP_RPT_PARAM l ON (b.ERP_RPT_PARAM_ID = l.ERP_RPT_PA' +
        'RAM_ID)'
      'WHERE TRUE'
      'AND a.ERP_RPT_ID = :ERP_RPT_ID'
      'ORDER BY URUT')
    Left = 115
    Top = 576
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ERP_RPT_ID'
        Value = nil
      end>
  end
  object qPeriod: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT'
      '  A.PERIOD_NAME,'
      '  PERIOD_DATE,'
      '  A.START_DATE,'
      '  DATE_FORMAT(A.PERIOD_DATE,'#39'%Y-%m-%d'#39') PERIOD_STR,'
      '  DATE_ADD(A.PERIOD_DATE, INTERVAL 1 MONTH) TGL_BARU,'
      '  A.YEARMONTH,'
      '  A.YEAR,'
      '  A.MONTH,'
      '  A.OPEN_FLAG,'
      '   PERIOD_AWAL'
      'FROM'
      '  period A'
      '&Kondisi')
    Left = 202
    Top = 576
    MacroData = <
      item
        Name = 'Kondisi'
      end>
    object qPeriodPERIOD_NAME: TIntegerField
      FieldName = 'PERIOD_NAME'
    end
    object qPeriodPERIOD_DATE: TDateTimeField
      FieldName = 'PERIOD_DATE'
    end
    object qPeriodSTART_DATE: TDateTimeField
      FieldName = 'START_DATE'
    end
    object qPeriodPERIOD_STR: TStringField
      FieldName = 'PERIOD_STR'
      Size = 10
    end
    object qPeriodTGL_BARU: TDateTimeField
      FieldName = 'TGL_BARU'
    end
    object qPeriodYEARMONTH: TIntegerField
      FieldName = 'YEARMONTH'
    end
    object qPeriodYEAR: TSmallintField
      FieldName = 'YEAR'
    end
    object qPeriodMONTH: TByteField
      FieldName = 'MONTH'
    end
    object qPeriodOPEN_FLAG: TStringField
      FieldName = 'OPEN_FLAG'
      FixedChar = True
      Size = 1
    end
    object qPeriodPERIOD_AWAL: TStringField
      FieldName = 'PERIOD_AWAL'
      Size = 10
    end
  end
  object CBank: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      
        'select    a.COA_ID, a.COA_NAME,COA_CODE,ifnull(b.MATA_UANG_ID,1)' +
        ' MATA_UANG_ID,'
      
        '(CASE when b.ACCOUNT_TYPE_ID=@ACC_GIRO then 1 else 0 end) as TIP' +
        'E,'
      'm.MATA_UANG_NAME,a.COA_NAME as BANK_NAME,b.ACCOUNT_NAME2'
      'FROM'
      '  COA a'
      'INNER JOIN ACCOUNT b ON (a.ACCOUNT_ID=b.ACCOUNT_ID)'
      'LEFT JOIN MATA_UANG m ON (m.MATA_UANG_ID=b.MATA_UANG_ID)'
      'LEFT JOIN erp_user_d_akun eud ON (a.COA_ID =eud.COA_ID)'
      ''
      
        'WHERE a.ACTIVE_FLAG = '#39'Y'#39' and (b.ACCOUNT_TYPE_ID=@ACC_KASBANK or' +
        ' b.ACCOUNT_TYPE_ID=@ACC_GIRO)'
      ' &kondisi'
      ' &kondisi2'
      '&kondisiUser'
      'GROUP BY A.COA_ID '
      'order by b.PRIMARY_FLAG desc,COA_CODE')
    Left = 29
    Top = 662
    MacroData = <
      item
        Name = 'kondisi'
      end
      item
        Name = 'kondisi2'
      end
      item
        Name = 'kondisiUser'
      end>
  end
  object CGudang: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '  a.WAREHOUSE_ID,'
      '  a.ADDRESS_ID,'
      '  a.PRIMARY_FLAG,'
      '  a.WAREHOUSE_NAME '
      'FROM'
      '  warehouse a'
      
        '  LEFT  JOIN Erp_warehouse g on a.WAREHOUSE_ID = g.WAREHOUSE_ID ' +
        'and erp_user_id = :USER '
      'WHERE   active_flag = '#39'Y'#39
      '&Kondisi'
      '&KondisiGudang'
      'GROUP BY a.WAREHOUSE_ID'
      'ORDER BY'
      '  IFNULL(g.PRIMARY_FLAG,a.PRIMARY_FLAG) DESC,'
      '  a.WAREHOUSE_NAME')
    Left = 115
    Top = 661
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'USER'
        Value = nil
      end>
    MacroData = <
      item
        Name = 'kondisi'
      end
      item
        Name = 'KondisiGudang'
      end>
  end
  object cSales: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT'
      'K.KARYAWAN_ID,'
      'K.FIRST_NAME,'
      'K.LAST_NAME, K.kata_depan,'
      'K.DESCRIPTION,'
      'b.DISPLAY_NAME BAGIAN,'
      'dd.DISPLAY_NAME DIVISI,'
      'bb.BRANCH_NAME CABANG,'
      'k.DIVISI_ID'
      'FROM KARYAWAN K'
      
        'LEFT JOIN ERP_LOOKUP_VALUE b ON (k.DEPT_ID = b.ERP_LOOKUP_VALUE_' +
        'ID)'
      
        'LEFT JOIN ERP_LOOKUP_VALUE dd ON k.DIVISI_ID=dd.ERP_LOOKUP_VALUE' +
        '_ID'
      
        'LEFT JOIN PERSON_SALES ps ON k.KARYAWAN_ID = ps.KARYAWAN_ID AND ' +
        'ps.PERSON_ID = :PERSON_ID'
      
        'LEFT JOIN Erp_USER_sales g on k.KARYAWAN_ID = g.KARYAWAN_ID and ' +
        'g.erp_user_id = :USER   '
      'LEFT JOIN BRANCH bb ON k.BRANCH_ID=bb.BRANCH_ID'
      
        'WHERE (k.dept_id=@SALES OR k.dept_id=@TELEMARKETING) and (k.ACTI' +
        'VE_FLAG ='#39'Y'#39') and '
      '((K.END_DATE = 0 ) or    (K.END_DATE IS NULL)) '
      '&kondisi'
      '&kondisi1'
      '&kondisiUser'
      'GROUP BY k.KARYAWAN_ID'
      'ORDER BY'
      ''
      'K.FIRST_NAME')
    Left = 202
    Top = 661
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'PERSON_ID'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'USER'
        Value = nil
      end>
    MacroData = <
      item
        Name = 'kondisi'
      end
      item
        Name = 'kondisi1'
      end
      item
        Name = 'kondisiUser'
      end>
  end
  object MEREK: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '  b.DISPLAY_NAME,'
      '  b.ERP_LOOKUP_VALUE_ID,'
      '  b.PRIMARY_FLAG'
      'FROM'
      '  erp_lookup_set a'
      
        '  INNER JOIN erp_lookup_value b ON (a.ERP_LOOKUP_SET_ID = b.ERP_' +
        'LOOKUP_SET_ID)'
      
        '  LEFT JOIN Erp_group_merek g on b.ERP_LOOKUP_VALUE_ID = g.MEREK' +
        '_ID and g.erp_user_id = :USER '
      'WHERE'
      '  PROGRAM_CODE = '#39'MEREK'#39
      'AND ACTIVE_FLAG  ='#39'Y'#39
      '&KONDISI'
      '&kondisiUser'
      'GROUP BY b.ERP_LOOKUP_VALUE_ID'
      ''
      'ORDER BY b.Primary_Flag DESC,b.DISPLAY_NAME')
    Left = 29
    Top = 749
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'USER'
        Value = nil
      end>
    MacroData = <
      item
        Name = 'kondisi'
      end
      item
        Name = 'kondisiUser'
      end>
  end
  object qUpdate_Coa_Balance: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      ' UPDATE COA_BALANCE SET '
      '  COA_SALDO =  COA_SALDO + :COA_SALDO,'
      '  COA_DEBET =  COA_DEBET + :COA_DEBET,'
      '  COA_CREDIT = COA_CREDIT + :COA_CREDIT'
      'WHERE  (COA_ID = :COA_ID) AND '
      '       (PERIOD_NAME = :PERIOD_NAME)'
      '')
    Left = 115
    Top = 749
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'COA_SALDO'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'COA_DEBET'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'COA_CREDIT'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'COA_ID'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'PERIOD_NAME'
        Value = nil
      end>
  end
  object TIPE: TwwMyQuery
    Connection = db1
    SQL.Strings = (
      'SELECT '
      '  b.DISPLAY_NAME,'
      '  b.ERP_LOOKUP_VALUE_ID,'
      '  b.PRIMARY_FLAG'
      'FROM'
      '  erp_lookup_set a'
      
        '  INNER JOIN erp_lookup_value b ON (a.ERP_LOOKUP_SET_ID = b.ERP_' +
        'LOOKUP_SET_ID)'
      
        '  LEFT JOIN Erp_group_divisi g on b.ERP_LOOKUP_VALUE_ID = g.DIVI' +
        'SI_ID and g.erp_user_id = :USER   '
      'WHERE'
      '  PROGRAM_CODE = '#39'TIPE'#39
      'AND ACTIVE_FLAG  ='#39'Y'#39
      '&KONDISI'
      '&kondisiUser'
      'GROUP BY b.ERP_LOOKUP_VALUE_ID'
      ''
      'ORDER BY b.Primary_Flag DESC,b.DISPLAY_NAME')
    Left = 298
    Top = 661
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'USER'
        Value = nil
      end>
    MacroData = <
      item
        Name = 'kondisi'
      end
      item
        Name = 'kondisiUser'
      end>
  end
  object qQry: TwwMyQuery
    Connection = db1
    Left = 298
    Top = 576
  end
  object IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL
    Destination = 'abdi-svr01.ddns.net:21'
    Host = 'abdi-svr01.ddns.net'
    MaxLineAction = maException
    Port = 21
    DefaultPort = 0
    ReadTimeout = 60000
    SSLOptions.Method = sslvSSLv2
    SSLOptions.SSLVersions = [sslvSSLv2]
    SSLOptions.Mode = sslmClient
    SSLOptions.VerifyMode = []
    SSLOptions.VerifyDepth = 0
    Left = 376
    Top = 667
  end
  object IdFTP: TIdFTP
    IOHandler = IdSSLIOHandlerSocket
    Host = 'abdi-svr01.ddns.net'
    ConnectTimeout = 0
    Password = '93vc487c34jfjnfc45n'
    Username = 'seventhsoft'
    NATKeepAlive.UseKeepAlive = False
    NATKeepAlive.IdleTimeMS = 0
    NATKeepAlive.IntervalMS = 0
    ProxySettings.ProxyType = fpcmNone
    ProxySettings.Port = 0
    Left = 373
    Top = 578
  end
  object RESTClient1: TRESTClient
    Accept = 'application/json, text/plain; q=0.9, text/html;q=0.8,'
    AcceptCharset = 'utf-8, *;q=0.8'
    BaseURL = 'https://api.seventhsoft.id/api/send-notification'
    Params = <>
    SynchronizedEvents = False
    Left = 624
    Top = 576
  end
  object RESTRequest1: TRESTRequest
    AssignedValues = [rvConnectTimeout, rvReadTimeout]
    Client = RESTClient1
    Params = <
      item
        Name = 'API_REPORT_ID'
        Options = [poDoNotEncode]
        Value = '10200'
      end>
    Response = RESTResponse1
    SynchronizedEvents = False
    Left = 864
    Top = 576
  end
  object RESTResponse1: TRESTResponse
    ContentType = 'application/json'
    Left = 739
    Top = 576
  end
end
