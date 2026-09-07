unit uSvReport;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs, Data.DB, MemDS, DBAccess, MyAccess,
  wwMyQuery, DAScript, MyScript, vcl.wwdatsrc, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdFTP, IdBaseComponent, IdComponent,
  IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  Winapi.WinSvc, system.JSON, System.Win.Registry, EncdDecd, System.StrUtils,
  FlexCel.Report, System.Variants, Winapi.ShellAPI, REST.Types, REST.Client,
  Data.Bind.Components, Data.Bind.ObjectScope, FlexCel.Core, FlexCel.XlsAdapter,
  FlexCel.Render, vcl.FlexCel.Core, FMX.FlexCel.Core, System.IOUtils;

const
  V_jum_delimiter_awal: byte = 4;

type
  TArrString = array of string;

type
  TsvReport = class(TService)
    con1: TMyConnection;
    qReportWeb: TwwMyQuery;
    qErpDetail: TwwMyQuery;
    qReportWebAPI_REPORT_ID: TIntegerField;
    qReportWebJSON: TMemoField;
    qReportWebSERVER_ID: TIntegerField;
    qReportWebERP_RPT_ID: TIntegerField;
    qReportWebGENERATE_FLAG: TIntegerField;
    qReportWebPDF_FLAG: TIntegerField;
    qReportWebREAD_FLAG: TBooleanField;
    qReportWebNOTIF_FLAG: TBooleanField;
    qReportWebERROR: TMemoField;
    qReportWebERROR_USER: TMemoField;
    qReportWebPATH_FILE: TStringField;
    qReportWebOUTPUT_FILE: TStringField;
    qReportWebCREATED_DATE: TDateTimeField;
    qReportWebLAST_UPDATED_DATE: TDateTimeField;
    qReportWebERP_USER_ID: TIntegerField;
    qReportWebHOSTNAME: TStringField;
    qReportWebPORT: TStringField;
    qReportWebDATABASE_NAME: TStringField;
    qReportWebTEMPLATE_FILE: TStringField;
    qErpDetailERP_DETAIL_ID: TStringField;
    qErpDetailWHERE_COLUMN: TStringField;
    qErpDetailWHERE_VALUE: TStringField;
    Query2: TwwMyQuery;
    qDB: TwwMyQuery;
    qReportWebDel: TwwMyQuery;
    qDBschema_name: TStringField;
    qReportWebDelAPI_REPORT_ID: TIntegerField;
    qReportWebDelJSON: TMemoField;
    qReportWebDelSERVER_ID: TIntegerField;
    qReportWebDelERP_RPT_ID: TIntegerField;
    qReportWebDelGENERATE_FLAG: TIntegerField;
    qReportWebDelPDF_FLAG: TIntegerField;
    qReportWebDelREAD_FLAG: TBooleanField;
    qReportWebDelNOTIF_FLAG: TBooleanField;
    qReportWebDelERROR: TMemoField;
    qReportWebDelERROR_USER: TMemoField;
    qReportWebDelPATH_FILE: TStringField;
    qReportWebDelOUTPUT_FILE: TStringField;
    qReportWebDelCREATED_DATE: TDateTimeField;
    qReportWebDelLAST_UPDATED_DATE: TDateTimeField;
    qReportWebDelERP_USER_ID: TIntegerField;
    qReportWebDelHOSTNAME: TStringField;
    qReportWebDelPORT: TStringField;
    qReportWebDelDATABASE_NAME: TStringField;
    qReportWebDelTEMPLATE_FILE: TStringField;
    db1: TMyConnection;
    Query1: TwwMyQuery;
    sql1: TMyScript;
    wwMyQuery1: TwwMyQuery;
    qClient: TwwMyQuery;
    qNo_Doc: TwwMyQuery;
    qUser: TwwMyQuery;
    qERP_SQL: TwwMyQuery;
    q1: TwwMyQuery;
    ds1: TwwDataSource;
    q2: TwwMyQuery;
    ds2: TwwDataSource;
    q3: TwwMyQuery;
    ds3: TwwDataSource;
    q4: TwwMyQuery;
    ds4: TwwDataSource;
    q5: TwwMyQuery;
    ds5: TwwDataSource;
    q6: TwwMyQuery;
    ds6: TwwDataSource;
    q7: TwwMyQuery;
    ds7: TwwDataSource;
    q8: TwwMyQuery;
    ds8: TwwDataSource;
    q9: TwwMyQuery;
    ds9: TwwDataSource;
    q10: TwwMyQuery;
    ds10: TwwDataSource;
    q11: TwwMyQuery;
    ds11: TwwDataSource;
    q12: TwwMyQuery;
    ds12: TwwDataSource;
    q13: TwwMyQuery;
    ds13: TwwDataSource;
    q14: TwwMyQuery;
    ds14: TwwDataSource;
    q15: TwwMyQuery;
    ds15: TwwDataSource;
    q16: TwwMyQuery;
    ds16: TwwDataSource;
    q17: TwwMyQuery;
    ds17: TwwDataSource;
    q18: TwwMyQuery;
    ds18: TwwDataSource;
    q19: TwwMyQuery;
    ds19: TwwDataSource;
    q20: TwwMyQuery;
    ds20: TwwDataSource;
    qHeader: TwwMyQuery;
    dsHeader: TwwDataSource;
    qDetail: TwwMyQuery;
    dsDetail: TwwDataSource;
    qDetail1: TwwMyQuery;
    dsDetail1: TwwDataSource;
    qDetail2: TwwMyQuery;
    dsDetail2: TwwDataSource;
    qDetail3: TwwMyQuery;
    dsDetail3: TwwDataSource;
    qDetail4: TwwMyQuery;
    dsDetail4: TwwDataSource;
    qNo_Form: TwwMyQuery;
    dsNo_Form: TwwDataSource;
    qSql: TwwMyQuery;
    dsSql: TwwDataSource;
    qGroup: TwwMyQuery;
    dsGroup: TwwDataSource;
    qErp_rpt_sql_name: TwwMyQuery;
    qSetup: TwwMyQuery;
    QListEditor: TwwMyQuery;
    qPeriod: TwwMyQuery;
    CBank: TwwMyQuery;
    CGudang: TwwMyQuery;
    cSales: TwwMyQuery;
    MEREK: TwwMyQuery;
    qUpdate_Coa_Balance: TwwMyQuery;
    TIPE: TwwMyQuery;
    qQry: TwwMyQuery;
    IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
    IdFTP: TIdFTP;
    qPeriodPERIOD_NAME: TIntegerField;
    qPeriodPERIOD_DATE: TDateTimeField;
    qPeriodSTART_DATE: TDateTimeField;
    qPeriodPERIOD_STR: TStringField;
    qPeriodTGL_BARU: TDateTimeField;
    qPeriodYEARMONTH: TIntegerField;
    qPeriodYEAR: TSmallintField;
    qPeriodMONTH: TByteField;
    qPeriodOPEN_FLAG: TStringField;
    qPeriodPERIOD_AWAL: TStringField;
    RESTClient1: TRESTClient;
    RESTRequest1: TRESTRequest;
    RESTResponse1: TRESTResponse;
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceAfterUninstall(Sender: TService);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceExecute(Sender: TService);
    procedure db1AfterConnect(Sender: TObject);
    procedure qErpDetailAfterScroll(DataSet: TDataSet);
    procedure sql1Error(Sender: TObject; E: Exception; SQL: string; var Action: TErrorAction);
  private
    { Private declarations }
    sheet_name, Extensi, VGroup, Kondisi_Group, tempall: string;
    all_jgn_tampil, dari_active, reproses_acc_flag: boolean;
    db: string;
    masuk_lagi: boolean;
    urut_tgl, urut_trans, posisi: byte;
    oldNilai_TGL1, oldNilai_TGL2: tdatetime;
    patch, vErr: string;
    dbase: array[0..4] of string;
    vOutPutFile, vFtpFlag: string;   // <-- baru
    vFtpBol: Boolean;                 // <-- baru
    N, panjang_posisi: Integer;

    NOT_REPORT_all_flag, NOT_SALES_all_flag, NOT_GUDANG_all_flag, NOT_BANK_all_flag, NOT_DIVISI_all_flag, NOT_MEREK_all_flag: Boolean;

    FArray: Variant;
    branch_id, table_transaksi, nama_template, defa_template: string;
    langsung, satu_sql, Kondisi_and: Boolean;
    memo1: TStringList;
    Filename, file_excel, extensi1: string;
    vFileFtp, vFilePdf: string;
    vDbFolder, vReportFolder: string;

    function LoginUlang(): Boolean;
    function ImagePatch(NamaService: string): string;
    function TryClaimReport(vApiReportId: string): Boolean;
    function LoginManage(vDB, vHostName, vPort: string): Boolean;
    function Run_query(vstr: string): Boolean;
    function keluar(): Boolean;
    function pertamakali(): Boolean;
    function createExcel(vReportId: string): Boolean;     //isi_sql
    function Define_Sql: Boolean;
    function rubahsqlsebelumexecute(no_urut: integer; awal: string; var ada_sql1: boolean): string;
    function Run_sql(vstr: string): Boolean;
    function Proses_data1: boolean;
    function str_filter(Harus_cek_all_flag: boolean): string;
    function filter_Period_awal(PERIOD: string): string;
    function filter_Period_Lalu(PERIOD: string): string;
    function filter_Period(date1: TDateTime): string;
    function filter_YEAR(date1: TDateTime): string;
    function filter_YEARMONTH(date1: TDateTime): string;
    function filter_YEAR_lalu(date1: TDateTime): string;
    function filter_YEAR2(date1: TDateTime): string;
    function filter_YEARMONTH2(date1: Tdate): string;
    function filter_Period2(date1: Tdate): string;
    function bulan(vbln: byte): string;
    function BLN_SINGKAT(vbln: byte): string;
    function Reproses_accounting(vperiod: string): Boolean;
    function Proses_Coa_Balance(VPeriod: string; VProgress: byte): Boolean;
    function Update_coa(VPeriod, vcoa_id: string; vsaldo, vdebet, vkredit: real): Boolean;
    function gantisql(nourut: integer; sumber: string): string;
    function hitungsql(NOURUT: integer): integer;
    function IsDirAccessible(const APath: string): Boolean;
    function SetInitAwalModule: Boolean;
    function DeleteFolderRecursive(const APath: string): Boolean;

    procedure CopyMyFile(SourceFile, DestFile: string);
    procedure mulai;
    procedure logFileWeb(vString: string);
    procedure logFile(vString: string);
    procedure inisialisasi_v; //untuk menginisialisasi setiap variable yang akan diparsingkan
    procedure ProcessOneQueueItem;
    procedure SendNotification(param1: string);
    procedure cleanReport;
    procedure SetAwal;
    procedure ProcessExcelToPdf(const InFile, OutFile: string);
    procedure CleanupOldLogs(vHariSimpan: Integer);
    procedure PARENT_SQL_HEADER(NOURUT: INTEGER);
    procedure PARENT_SQL_DETAIL(NOURUT: INTEGER);
    procedure PARENT_SQL_DETAIL1(NOURUT: INTEGER);
    procedure PARENT_SQL_DETAIL2(NOURUT: INTEGER);
    procedure PARENT_SQL_DETAIL3(NOURUT: INTEGER);
    procedure PARENT_SQL_DETAIL4(NOURUT: INTEGER);
    procedure PARENT_SQL_NO_FORM(NOURUT: INTEGER);
    procedure PARENT_SQL_1(NOURUT: INTEGER);
    procedure PARENT_SQL_2(NOURUT: INTEGER);
    procedure PARENT_SQL_3(NOURUT: INTEGER);
    procedure PARENT_SQL_4(NOURUT: INTEGER);
    procedure PARENT_SQL_5(NOURUT: INTEGER);
    procedure PARENT_SQL_6(NOURUT: INTEGER);
    procedure PARENT_SQL_7(NOURUT: INTEGER);
    procedure PARENT_SQL_8(NOURUT: INTEGER);
    procedure PARENT_SQL_9(NOURUT: INTEGER);
    procedure PARENT_SQL_10(NOURUT: INTEGER);
    procedure PARENT_SQL_11(NOURUT: INTEGER);
    procedure PARENT_SQL_12(NOURUT: INTEGER);
    procedure PARENT_SQL_14(NOURUT: INTEGER);
    procedure PARENT_SQL_15(NOURUT: INTEGER);
    procedure PARENT_SQL_16(NOURUT: INTEGER);
    procedure PARENT_SQL_17(NOURUT: INTEGER);
    procedure PARENT_SQL_18(NOURUT: INTEGER);
    procedure PARENT_SQL_19(NOURUT: INTEGER);
    procedure PARENT_SQL_20(NOURUT: INTEGER);
  public
    vFileNotif, Group_id, FInstanceName, vHostNames, vError, vFTP: string;
    function GetServiceController: TServiceController; override;
    procedure SetInstanceName(const Value: string);
    procedure ChangeServiceConfiguration;   //nassrul 14/02/18
    property InstanceName: string read FInstanceName write SetInstanceName;
    { Public declarations }
  end;

const
  maxvararray = 30;
  maxcombo = 40;
  maxquery = 26;
  Tgroup_id1 = '&Group_id1';
  Tgroup_id = '&Group_id';
  AliasGroup = 'Xgroup';
  AliasGroup1 = 'Xgroup1';
  TDepartment_id = '&Department_id';
  AngkaNonAktif = 77;
  maxarrayparam = 100;

var
  anu, anu2, anu3: string;
  svReport: TsvReport;
  Ar1: array[0..maxvararray] of string = ('Asia Elektrindo', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', 'Robbin', '21', '22', '23', '24', '25', 'ADMIN', 'PASSWORD', '28', '29', '30');
  v_a_param_yg_dipasing: array[1..2, 0..maxcombo] of string;

  //NAMA VARIABEL YG MENUNJANG/MENDUKUNG TABEL ERP_RPT_PARAM
  yearmonth1, yearmonth2, YEAR1, YEAR2, year_lalu: string;
  PERIOD_LALU, PERIOD1, PERIOD2, PERIOD_AWAL: string;
  cekbox_id,                                       // id pada checkbox
isi_nama_id,                                         // isi id pada combobox pertama
isi_nama_id2, isi_id, isi_id2, isi_kata, isi_kata2                                        // isi id pada combobox KEDUA
: array[1..maxcombo] of string;              //
  Str, Str_Alias: array[1..maxcombo] of string; //untuk menyimpan nama alias dari tiap-tiap parameter
  COMBO_LOOKUP_FIELD: array[1..maxcombo] of integer;
  DISPLAY_VALUE_FLAG: array[1..maxcombo] of BOOLEAN;

  //NAMA VARIABEL YG DIPAKAI UNTUK MENDUKUNG QUERY
  PARAM_YG_DIAMBIL: array[0..maxquery] of TARRSTRING;
  JUM_PARAM_YG_DIAMBIL: array[0..maxquery] of INTEGER;

  //NAMA VARIABEL UNTUK UTILITY DARI TABEL ERP_RPT_PARAM
  jumparam: array[1..maxcombo] of integer;
  isinamaparam1: array[1..maxcombo] of TarrString;

  //NAMA VARIABLE YANG DIGUNAKAN UNTUK MENGISI KONDISI
  nama_Label_id, nama_Label_id2, nama_id,                                         // id pada combobox pertama
nama_id2,                                         // id pada combobox KEDUA
LABEL_MK_1_1, LABEL_MK_1_3, LABEL_MK_1_5, label_lookup: array[1..maxcombo] of string;              // id pada combobox kedua

       //NAMA VARIABLE YANG DIGUNAKAN UNTUK PARAMETER DIBAWAH
  PARAM_TYPE, LOOK_PARENT: array[1..MAXCOMBO] of INTEGER;
  COMBO_ALIAS, CHILD_FIELD, CHILD_VALUE: array[1..maxcombo] of string;              // id pada combobox kedua

  //NAMA VARIABEL YANG DI ERP_RPT_SQL
  LABEL_source: array[0..maxQUERY] of string;
  master_source: array[0..maxQUERY] of integer;
  Kondisiyangdiambil: array[0..maxquery] of tarrstring;
  paramkondisiyangdiambil: array[0..maxquery] of tarrstring;
  JUM_KONDISI_YG_DIAMBIL: array[0..maxquery] of integer;
  PARENT_SQL: array[0..MAXQUERY] of INTEGER;
  jumlah_isi_query: array[0..maxquery] of integer;

  //VARIABEL SESUAI DENGAN JUMLAH MAX PARAMETER
  punyaanak: array[1..maxarrayparam] of integer; // untuk mengecek apa punya anak
  punyabapak: array[1..maxarrayparam] of integer; // untuk mengecek apa punya BAPAK
  Posisianak: array[1..maxarrayparam] of integer;
  Posisibapak: array[1..maxarrayparam] of integer;

  //NAMA VARIABEL YG DIPAKAI UNTUK MENDEFINISIKAN NAMA YANG TABEL ERP_RPT_PARAM
  kbarang: array[1..maxcombo] of string;
  BANYAKBARIS: array[0..maxquery] of integer;
  OPERATOR_CHAR: array[1..maxcombo] of string;
  passing_display_combo: array[1..maxcombo] of string;
  oldNILAI_nama_field: array[1..30] of string;
  oldNILAI_satu: array[1..30] of string;
  oldNILAI_satu_id: array[1..30] of string;
  oldNILAI_dua: array[1..30] of string;
  oldNILAI_dua_id: array[1..30] of string;
  oldNILAI_ALL: array[1..30] of Boolean;

    //CHOICE_VAR : ARRAY [1..MAXCOMBO] OF integer;

  //NAMA VARIABEL YG DIGUNAKAN UNTUK BERHUBUNGAN DENGAN DATABASE
  myNewQuery, myNewQuery2: array[1..maxcombo] of TMyQuery;
  Kodisitiapparam, skondisiall: array[0..maxcombo] of string;  //untuk menyimpan isi makro tiap-tiap query

  //NAMA2 VARIABEL GLOBAL
  error_kesalahan: string;
  pesan_awal: string;
  no_error: integer;
  jarakbutton: integer = 22;
  ada_sql1, ada_group_id: boolean;
  nama_sql: string;

  //variable web
//  js, Items, Item: TlkJSONBase;

implementation

const
  vTemplateFolder = 'TEMPLATE';

{$R *.dfm}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  svReport.Controller(CtrlCode);
end;

function TsvReport.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

function pisahdelimeterkhususkondisi(sumber, delimeter: string; var param, isiparam: Tarrstring; var jumlah: integer): boolean;
var
  s: string;
  sukses: boolean;
begin
  s := sumber;
  param[jumlah] := copy(sumber, 1, V_jum_delimiter_awal - 1);
  isiparam[jumlah] := copy(sumber, V_jum_delimiter_awal + 1, length(sumber) - V_jum_delimiter_awal + 1);
  inc(jumlah);
  sukses := true;
 {
  awal := 1;
  sukses := false;
  i := 0;
  if sumber <>'' then
  repeat
  //for I := 0 to length(sumber) do begin
    s := sumber[i];
      if (s=delimeter) or (i>length(sumber)) then
      begin
        param[jumlah] := copy (sumber, awal, i-awal);
        isiparam[jumlah] := copy (sumber,i+1, length(sumber)-i+1);
        sumber := '';
        inc(jumlah);
        i:= -1;
        sukses := true;
      end;
    inc(i);
  //end;
  until (length(sumber)=0);}
  result := sukses;
  //result := pisahdelimeterkhususkondisi;
end;

procedure TsvReport.ServiceAfterInstall(Sender: TService);
var
  log: TextFile;
  Reg: TRegistry;
  Key: string;
begin
  if (FInstanceName <> '') then
  begin
    ChangeServiceConfiguration;
  end;
  Key := '\SYSTEM\CurrentControlSet\Services\Eventlog\Application\' + Self.Name;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey(Key, True) then
    begin
      Reg.WriteString('EventMessageFile', ParamStr(0));
      Reg.WriteInteger('TypesSupported', 7);
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TsvReport.ServiceAfterUninstall(Sender: TService);
var
  Reg: TRegistry;
  Key: string;
begin
  // Delete registry entries for event viewer.
  Key := '\SYSTEM\CurrentControlSet\Services\Eventlog\Application\' + Self.Name;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.KeyExists(Key) then
      Reg.DeleteKey(Key);
  finally
    Reg.Free;
  end;
end;

function TsvReport.LoginUlang(): Boolean;
begin
  Result := false;
  try
    logFileWeb('database Connecting ' + dbase[1] + ' on ' + dbase[0] + ':' + dbase[4]);
//    logFile('database Connecting ' + dbase[1] + ' on ' + dbase[0] + ':' + dbase[4]);
    con1.Disconnect;
    con1.Port := StrToInt(dbase[4]);
    con1.Server := dbase[0];
    con1.Database := dbase[1];
    con1.Username := 'Fatra';
    con1.Password := '73fangfang';
//    logFile('user_password ' + con1.Username + '_' + con1.Password);
    con1.Connect;
    logFileWeb('database Connect ' + con1.Database + ' on ' + con1.Server + ':' + IntToStr(con1.Port));
    Result := True;
  except
    on E: exception do
    begin
      logFile('LoginUlang connect database  ' + con1.Database + ' failed ' + E.Message + '!');
    end;
  end;
  Result := con1.Connected;
end;

procedure TsvReport.ServiceExecute(Sender: TService);
begin
  db := '''';
  panjang_posisi := 1;
  while not Terminated do
  begin
    if qReportWeb.Active then
      qReportWeb.Close;
    if (Time >= EncodeTime(0, 0, 0, 0)) and (Time < EncodeTime(4, 0, 0, 0)) then
    begin
      DeleteFolderRecursive(patch + 'Report\');
      CleanupOldLogs(10); // hapus log yang usianya lebih dari 10 hari
    end;

    if LoginUlang then
    begin
      try
        qReportWeb.Close;
        qReportWeb.MacroByName('KONDISI').Value := ' AND a.HOSTNAME = ' + QuotedStr(vHostNames);
        qReportWeb.Open;
        logFile('Jumlah data antrian ' + IntToStr(qReportWeb.RecordCount));

        while not qReportWeb.Eof do
        begin
          vFileFtp := '';
          cleanReport;
          ProcessOneQueueItem;
          qReportWeb.Next;
        end;
      except
        on E: exception do
        begin
          logFile(' report error ' + E.Message);
        end;
      end;
    end;
    Sleep(1000);
  end;
end;

procedure TsvReport.ServiceStart(Sender: TService; var Started: Boolean);
begin
  SetAwal;
end;

procedure TsvReport.SetInstanceName(const Value: string);          //nassrul 14/02/18
begin
  if (FInstanceName <> Value) then
  begin
    FInstanceName := Value;
    if (FInstanceName <> '') then
    begin
      Self.Name := Name + FInstanceName;
      Self.DisplayName := Format(DisplayName + ' (%s)', [FInstanceName]);
    end;
  end;
end;

function TsvReport.ImagePatch(NamaService: string): string;
var
  Registry: TRegistry;
  str, right, left: string;
  pos, temp: integer;
begin
  Registry := TRegistry.Create;
  try
    Registry.Lazywrite := false;
    Registry.RootKey := HKEY_LOCAL_MACHINE;
    Registry.OpenKey('\SYSTEM\CurrentControlSet\Services\' + NamaService + '\', true);
    str := Registry.ReadString('ImagePath');

    pos := length(str);
    right := str;
    temp := 0;
    while pos > 0 do
    begin
      pos := AnsiPos('\', right);
      temp := temp + pos;
      right := RightStr(right, length(right) - pos);
    end;
    left := LeftStr(str, temp);

    Result := left;
    Registry.CloseKey;
  finally
    Registry.Free;
  end; // try..finally
end;

procedure TsvReport.ChangeServiceConfiguration;   //nassrul 14/02/18
var
  mngr: Cardinal;
  svc: Cardinal;
  newpath: string;
begin
  // Open the service manager
  mngr := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  if (mngr = 0) then
    RaiseLastOSError;
  try
    // Open the service
    svc := OpenService(mngr, PChar(Self.Name), SERVICE_CHANGE_CONFIG);
    if (svc = 0) then
      RaiseLastOSError;
    try
      // Change the service params
      newpath := ParamStr(0) + ' ' + Format('-instance="%s"', [FInstanceName]); // + any other cmd line params you fancy
      ChangeServiceConfig(svc, SERVICE_NO_CHANGE, //  dwServiceType
        SERVICE_NO_CHANGE, //  dwStartType
        SERVICE_NO_CHANGE, //  dwErrorControl
        PChar(newpath),    //  <-- The only one we need to set/change
        nil,               //  lpLoadOrderGroup
        nil,               //  lpdwTagId
        nil,               //  lpDependencies
        nil,               //  lpServiceStartName
        nil,               //  lpPassword
        nil);              //  lpDisplayName
    finally
      CloseServiceHandle(svc);
    end;
  finally
    CloseServiceHandle(mngr);
  end;
end;

procedure TsvReport.logFileWeb(vString: string);
var
  slSave: TStringlist;
  vServerLog, vDatabaseLog, vLogFileName, vLogFolder: string;
begin
  vServerLog := db1.Server;
  vDatabaseLog := db1.Database;
  if qReportWeb.Active then
  begin
    vServerLog := qReportWebHOSTNAME.AsString;
    vDatabaseLog := qReportWebDATABASE_NAME.AsString;
  end;

  vLogFolder := IncludeTrailingPathDelimiter(patch + 'Log');
  if not DirectoryExists(vLogFolder) then
  begin
    try
      ForceDirectories(vLogFolder);
    except
      vLogFolder := patch;
    end;
  end;

  vLogFileName := vLogFolder + 'track' + FormatDateTime('yymmddhh', Now) + '.txt';

  slSave := TStringlist.Create;
  try
    try
      if FileExists(vLogFileName) then
      begin
        slSave.LoadFromFile(vLogFileName);
        slSave.Insert(0, FormatDateTime('dd mmmm yyyy', Now) + ' jam ' + FormatDateTime('hh:nn:ss', now) + ' server ' + vServerLog + ' db ' + vDatabaseLog + ' ' + vString);
      end
      else
        slSave.Text := FormatDateTime('dd mmmm yyyy', Now) + ' jam ' + FormatDateTime('hh:nn:ss', now) + ' server ' + vServerLog + ' db ' + vDatabaseLog + ' ' + vString;

      slSave.SaveToFile(vLogFileName);
    except
      // sengaja diredam
    end;
  finally
    slSave.Free;
  end;
end;

procedure TsvReport.logFile(vString: string);
var
  slSave: TStringlist;
  vServerLog, vDatabaseLog, vLogFileName, vLogFolder: string;
begin
  vServerLog := db1.Server;
  vDatabaseLog := db1.Database;
  if qReportWeb.Active then
  begin
    vServerLog := qReportWebHOSTNAME.AsString;
    vDatabaseLog := qReportWebDATABASE_NAME.AsString;
  end;

  // semua file log ditaruh di subfolder "Log\" di bawah folder exe,
  // dibuat otomatis kalau belum ada
  vLogFolder := IncludeTrailingPathDelimiter(patch + 'Log');
  if not DirectoryExists(vLogFolder) then
  begin
    try
      ForceDirectories(vLogFolder);
    except
      // kalau gagal buat folder (mis. masalah permission), fallback ke folder exe
      vLogFolder := patch;
    end;
  end;

  // nama file log per-instance, supaya beberapa service tidak menulis
  // ke file yang sama secara bersamaan
  vLogFileName := vLogFolder + 'log' + FormatDateTime('yymmddhh', Now) + '.txt';

  slSave := TStringlist.Create;
  try
    try
      if FileExists(vLogFileName) then
      begin
        slSave.LoadFromFile(vLogFileName);
        slSave.Insert(0, FormatDateTime('dd mmmm yyyy', Now) + ' jam ' + FormatDateTime('hh:nn:ss', now) + ' server ' + vServerLog + ' db ' + vDatabaseLog + ' ' + vString);
      end
      else
        slSave.Text := FormatDateTime('dd mmmm yyyy', Now) + ' jam ' + FormatDateTime('hh:nn:ss', now) + ' server ' + vServerLog + ' db ' + vDatabaseLog + ' ' + vString;

      slSave.SaveToFile(vLogFileName);
    except
      // sengaja diredam: kegagalan tulis log tidak boleh menghentikan proses utama
    end;
  finally
    slSave.Free;
  end;
end;

function TsvReport.IsDirAccessible(const APath: string): Boolean;
var
  vTestFile: string;
  vFile: TextFile;
begin
  Result := False;
  if Trim(APath) = '' then
    Exit;
  try
    if not DirectoryExists(APath) then
    begin
      if not ForceDirectories(APath) then
        Exit;
    end;

    vTestFile := IncludeTrailingPathDelimiter(APath) + '_test_access.tmp';
    try
      AssignFile(vFile, vTestFile);
      Rewrite(vFile);
      CloseFile(vFile);
      DeleteFile(vTestFile);
      Result := True;
    except
      Result := False;
    end;
  except
    Result := False;
  end;
end;

procedure TsvReport.SetAwal;
var
  i, vTitik: integer;
  jam_awal, vServerPort, Txt, aFTP: string;
  soft: TextFile;
  vSoftOpened: Boolean;   // <-- baru, penanda file benar-benar ter-Reset

  function Base64Decode(const S: string): string;
  var
    Input: TStringStream;
    Output: TStringStream;
  begin
    Input := TStringStream.Create(S);
    Output := TStringStream.Create('');
    try
      DecodeStream(Input, Output);
      Result := Output.DataString;
    finally
      Input.Free;
      Output.Free;
    end;
  end;

  function Base64DecodeCustom(const AEncoded, ASecret: string): string;
  var
    S: string;
    PosRead: Integer;
    SecretLen: Integer;
  begin
    S := '';
    SecretLen := Length(ASecret);
    PosRead := 1;
    while PosRead <= Length(AEncoded) do
    begin
      // Ambil 4 karakter Base64 asli
      S := S + Copy(AEncoded, PosRead, 4);
      Inc(PosRead, 4);

      // Jika setelahnya adalah secret, SKIP secret tersebut
      if (PosRead <= Length(AEncoded)) and (Copy(AEncoded, PosRead, SecretLen) = ASecret) then
        Inc(PosRead, SecretLen);
    end;
    Result := Base64Decode(S);
  end;

begin
  patch := ImagePatch(Self.Name);
  memo1 := TStringList.Create;
  {
  baris 1 ambil hostname dan port
  baris 2 ambil nama database
  baris 3 proses convert pdf pada server apa (hostname)
  baris 4 alamat ftp yang di password
  baris 5 output PDF
  baris 6 flag pakai FTP (Y/N)
  }
  logFile('======Setting awal======');
  vFileNotif := patch + 'appApiNotifReport.exe';
  i := 1;
  vSoftOpened := False;

  if FileExists(patch + '7soft.txt') then
  begin
    try
      assignfile(soft, patch + '7soft.txt');
      Reset(soft);
      vSoftOpened := True;

      while not eof(soft) do
      begin
        Readln(soft, Txt);
        if i = 1 then
          vServerPort := Txt
        else if i = 2 then
          dbase[1] := Txt
        else if i = 3 then
          vHostNames := Txt
        else if i = 4 then
          aFTP := Txt
        else if i = 5 then
          vOutPutFile := Txt
        else if i = 6 then
        begin
          vFtpFlag := Txt;
          Break;
        end;
        i := i + 1;
      end;
    except
      on E: exception do
        logFile('Error baca 7soft.txt: ' + E.Message);
    end;

    if vSoftOpened then
    begin
      try
        CloseFile(soft);
      except
        // abaikan jika sudah tertutup / gagal ditutup
      end;
    end;
  end
  else
  begin
    logFile('file 7soft.txt tidak ada di ' + patch);
  end;

  vFtpBol := vFtpFlag = 'Y';

  if aFTP <> '' then
    vFTP := Base64Decode(aFTP);
  if vFTP = '' then
    vFTP := 'https://seventhsoft.net';

  if aFTP <> '' then
    vFTP := Base64DecodeCustom(aFTP, 'Q'); // hapus karakter Q setiap 4 karakter
  if vFTP = '' then
    vFTP := 'abdi-svr01.ddns.net';

  dbase[2] := 'Fatra';
  dbase[3] := '73fangfang';
  vTitik := pos(':', vServerPort);
  if vTitik > 0 then
  begin
    dbase[0] := copy(vServerPort, 0, vTitik - 1);
    dbase[4] := copy(vServerPort, vTitik + 1, length(vServerPort));
  end
  else
  begin
    dbase[0] := vServerPort;
    dbase[4] := '3306';
  end;

  if IsDirAccessible(vOutPutFile) then
  begin
    logFile('Output PDF menggunakan vOutPutFile: ' + vOutPutFile);
  end
  else
  begin
    vOutPutFile := patch + 'REPORT\PDF\';
    logFile('vOutPutFile (' + vOutPutFile + ') tidak bisa diakses, pakai folder default: ' + vOutPutFile);
  end;
  logFile('connecting server:' + dbase[0] + ' -- port:' + dbase[4]);
  logFile('Baris 1 untuk nama server');
  logFile('Baris 2 untuk nama database manage');
  logFile('Baris 3 untuk proses convert pdf pada server apa');
  logFile('Baris 4 alamat ftp yang di password');
  logFile('Baris 5 output PDF');
  logFile('Baris 6 flag FTP (Y/N)');
  logFile('======Setting awal selesai======');
end;

procedure TsvReport.cleanReport;
var
  vQuery: TwwMyQuery;
  i: Integer;
begin
  for i := 1 to 20 do
  begin
    vQuery := TwwMyQuery(FindComponent('q' + IntToStr(i)));
    if vQuery <> nil then
    begin
      vQuery.SQL.Clear;
      vQuery.Close;
    end;
  end;
  qHeader.SQL.Clear;
  qHeader.Close;
  qDetail.SQL.Clear;
  qDetail.Close;
  for i := 1 to 4 do
  begin
    vQuery := TwwMyQuery(FindComponent('qDetail' + IntToStr(i)));
    if vQuery <> nil then
    begin
      vQuery.SQL.Clear;
      vQuery.Close;
    end;
  end;

end;

function TsvReport.TryClaimReport(vApiReportId: string): Boolean;
var
  qClaim: TMyQuery;
  vSql: string;
begin
  Result := False;
  qClaim := TMyQuery.Create(nil);
  try
    qClaim.Connection := con1;
    vSql := 'UPDATE API_REPORT ';
    vSql := vSql + 'SET GENERATE_FLAG = 2 , WORKER = ' + QuotedStr(FInstanceName);
    vSql := vSql + ' WHERE API_REPORT_ID = ' + vApiReportId;
    vSql := vSql + '  AND (GENERATE_FLAG IS NULL OR GENERATE_FLAG <> 2)';
    qClaim.SQL.Text := vSql;
    try
      qClaim.ExecSQL;
      Result := qClaim.RowsAffected > 0;
      if Result then
        logFile('Berhasil klaim report id ' + vApiReportId)
      else
        logFile('Report id ' + vApiReportId + ' sudah diklaim service lain, skip');
    except
      on E: exception do
      begin
        logFile('TryClaimReport gagal untuk id ' + vApiReportId + ' error ' + E.Message);
        Result := False;
      end;
    end;
  finally
    qClaim.Free;
  end;
end;

function TsvReport.LoginManage(vDB, vHostName, vPort: string): Boolean;
begin
  Result := false;
  try
    logFile('database Connecting ' + vDB + ' on ' + vHostName + ':' + vPort);
    logFileWeb('database Connecting ' + vDB + ' on ' + vHostName + ':' + vPort);
    db1.Disconnect;
    db1.Port := StrToInt(vPort);
    db1.Server := vHostName;
    db1.Database := vDB;
    db1.Username := 'Fatra';
    db1.Password := '73fangfang';
    db1.Connect;
    logFile('database Connect ' + db1.Database + ' on ' + db1.Server + ':' + IntToStr(db1.Port));
    logFileWeb('database Connect ' + db1.Database + ' on ' + db1.Server + ':' + IntToStr(db1.Port));
    Result := true;
  except
    on E: exception do
    begin
      logFile('LoginManage connect database  ' + db1.Database + ' failed ' + E.Message + '!');
    end;
  end;
end;

function TsvReport.Run_query(vstr: string): Boolean;
begin
  Result := False;
  Query1.CLose;
  Query1.SQL.Clear;
  Query1.SQL.text := vstr;
//  logFile(vstr);
  try
    Query1.Open;
    Result := True;
//    logFile('Query1 sukses');
  except
    logFile('Query1 ' + vstr + ' Bermasalah!');
  end;
end;

function TsvReport.keluar(): Boolean;
var
  Allitem, tipeparam, urut: integer;
begin
  logFile('Bersihkan query parameter');
  for urut := 0 to maxcombo do
  begin
    try
      if myNewQuery[urut] <> nil then
        FreeAndNil(myNewQuery[urut]);
    except
    end;
    try
      if myNewQuery2[urut] <> nil then
      begin
        FreeAndNil(myNewQuery2[urut]);
      end;
    except
    end;
  end;

end;

function pisahdelimeter(sumber, delimeter: string; var jumlah: integer): Tarrstring;
var
  s: string;
  temp: tarrstring;
  i, awal: integer;
begin
  s := sumber;
  jumlah := 0;
  awal := 1;
  setlength(temp, 20);
  i := 0;
  if sumber <> '' then
    repeat
  //for I := 0 to length(sumber) do begin
      s := sumber[i];
      if (s = delimeter) or (i > length(sumber)) then
      begin
        temp[jumlah] := copy(sumber, awal, i - awal);
        sumber := copy(sumber, i + 1, length(sumber) - i + 1);
        inc(jumlah);
        i := -1;
      end;
      inc(i);
  //end;
    until (length(sumber) = 0);

  result := temp;
end;

function TsvReport.pertamakali(): Boolean;
var
  i, jumlahparameter, tipeparam, allitem, MAXwidth, jum_param1, jum_SQL, j, k, R: integer;
  report_id_terbanyak_dalam_modul, s, namaparameter: string;
  vQuery, vQuery2, vlimit_global, vFieldName: string;
  panjangnamaparam1, namaparam1, sqlparam: Tarrstring;
  URUT: INTEGER;
  tgl_pertama, vOpenPertama, vOpenKedua: boolean;
begin
  vOpenPertama := False;
  vOpenKedua := False;
  Result := False;
//  logFile('Proses pertamakali');
  vlimit_global := '25';
  try
    QListEditor.Close;
    QListEditor.ParamByName('ERP_RPT_ID').AsString := qReportWebERP_RPT_ID.AsString;
    QListEditor.Open;
    logFile('QListEditor Open');
  except
    on E: exception do
    begin
      logFile('QListEditor failed open ' + E.Message + '!');
      vError := 'QListEditor failed open ' + E.Message + '!';
      Exit;
    end;
  end;

  qlistEditor.First;
  i := 1;
  MAXwidth := 0;
  logFile('deklarasi array punyaanak, punyabapak, posisianak, posisibapak');
  fillchar(punyaanak, sizeof(punyaanak), 0);
  fillchar(punyabapak, sizeof(punyabapak), 0);
  fillchar(posisianak, sizeof(posisianak), 0);
  fillchar(posisibapak, sizeof(posisibapak), 0);
  logFile('selesai deklarasi');

  tgl_pertama := false;
  logFile('jumlah data qlistEditor ' + IntToStr(QListEditor.RecordCount));
  qErpDetail.Filtered := False;
  logFile('jumlah data qErpDetail ' + IntToStr(qErpDetail.RecordCount));
  while not qlistEditor.Eof do
  begin
    qErpDetail.Filtered := False;
    qErpDetail.Filter := 'ERP_DETAIL_ID = ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString;
    qErpDetail.Filtered := True;
    logFile(' filter erp_detail_id ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString);
    URUT := qlistEditor.fieldbyname('urut').AsInteger;
    if qErpDetail.RecordCount > 0 then
    begin

      isi_nama_id[URUT] := '';
      isi_nama_id2[URUT] := '';
      logFile('Proses 1');

      nama_id[URUT] := qlistEditor.fieldbyname('MK_2_1_SAME_VALUE').ASstring; //'TABLE_NAME
      nama_id2[URUT] := qlistEditor.fieldbyname('MK_2_1_DIFF_VALUE').ASstring; //'TABLE_NAME
      OPERATOR_CHAR[URUT] := qlistEditor.fieldbyname('MK_OPERATOR').Asstring;

      posisianak[qlistEditor.fieldbyname('erp_RPT_PARAM_id').AsInteger] := qlistEditor.fieldbyname('urut').AsInteger;

        //variable untuk label
      nama_Label_id[URUT] := qlistEditor.FIELDBYNAME('MK_1_2_SAME_FIELD').ASstring;
      nama_Label_id2[URUT] := qlistEditor.FIELDBYNAME('MK_1_2_DIFF_FIELD').ASstring;
      COMBO_LOOKUP_FIELD[URUT] := qlistEditor.FIELDBYNAME('COMBO_LOOKUP_FIELD').ASinteger;
      DISPLAY_VALUE_FLAG[URUT] := qlistEditor.FIELDBYNAME('DISPLAY_VALUE_FLAG').ASSTRING = '0';
      LABEL_MK_1_1[URUT] := qlistEditor.FIELDBYNAME('MK_1_1').ASstring;
      LABEL_MK_1_3[URUT] := qlistEditor.FIELDBYNAME('MK_1_3').ASstring;
      LABEL_MK_1_5[URUT] := qlistEditor.FIELDBYNAME('MK_1_5').ASstring;
      label_lookup[URUT] := qlistEditor.FIELDBYNAME('MK_1_4_LOOKING_ID').ASstring;
      CHILD_FIELD[URUT] := qlistEditor.FIELDBYNAME('CHILD_FIELD').ASstring;
      COMBO_ALIAS[URUT] := qlistEditor.FIELDBYNAME('COMBO_ALIAS').ASstring;
      CHILD_VALUE[URUT] := qlistEditor.FIELDBYNAME('CHILD_VALUE').ASstring;
      LOOK_PARENT[URUT] := qlistEditor.FIELDBYNAME('PARENT_LOOK_FLAG').ASINTEGER;
      PARAM_TYPE[URUT] := qlistEditor.FIELDBYNAME('PARAM_TYPE').ASinteger;

      allitem := qlistEditor.FIELDBYNAME('DISPLAY_ALL_FLAG').ASinteger;
      logFile('Proses 2');

      tipeparam := qlistEditor.FIELDBYNAME('PARAM_TYPE').ASinteger;
        //sqlparam := pisahdelimeter(qlistEditor.FIELDBYNAME('SQL_USES').ASSTRING,'|',jum_sql);
      punyaanak[URUT] := qlistEditor.FIELDBYNAME('PARENT_FROM_ID').ASinteger;

      if punyaanak[URUT] <> 0 then
        punyabapak[punyaanak[URUT]] := i
      else
        punyabapak[punyaanak[URUT]] := 0;

        { fatra feb 2021
         sengaja di buat 2 param semua dulu trus di hide
         krn lap hpp = 2 param di periode tapi keluar cuma 1 periode akhir nya error
         baris = 2827
         ELSE if (((jumlahparameter=2) and (dblookuph[urut].value=dblookupd[urut].value))or ((jumlahparameter=1)and (punyaanak[urut]=0))) then

         saat di klik xls ...krn combo tdk pernah terbentuk

        jumparam[URUT] := qlistEditor.FIELDBYNAME('DISPLAY_DOUBLE_FLAG').ASinteger;
        jumlahparameter := qlistEditor.FIELDBYNAME('DISPLAY_DOUBLE_FLAG').ASinteger;
         }
      jumparam[URUT] := 2;
      jumlahparameter := 2;
      if qlistEditor.FIELDBYNAME('DISPLAY_DOUBLE_FLAG').ASinteger = 1 then
        jumlahparameter := qlistEditor.FIELDBYNAME('DISPLAY_DOUBLE_FLAG').ASinteger;

      logFile('Proses 3');
      if (tipeparam = 2) then //text biasa
      begin
      end
      else if (tipeparam = 1) or (tipeparam = 4) or (tipeparam = 3) then //combo box biasa
      begin

        logFile('Proses 4');
        ANU := INTTOSTR(i);

        myNewQuery[URUT] := TMyQuery.Create(Self);
        myNewQuery[URUT].Name := 'qPrm' + QListEditor.fieldbyname('ERP_RPT_PARAM_ID').AsString + '1';

        myNewQuery[URUT].Connection := db1;
        myNewQuery[URUT].Tag := URUT;
          //myNewQuery[urut].CachedUpdates := true;
        s := qlistEditor.FIELDBYNAME('COMBO_SQL').ASSTRING;
        ada_Group_id := AnsiContainsStr(s, Tgroup_id);
        if ada_Group_id then
        begin
          s := AnsiReplaceStr(s, Tgroup_id, ' AND erp_group_id = ' + db + Group_id + db);
        end;

        myNewQuery[URUT].SQL.Add(s);
          //If punyabapak[URUT] <> 0 THEN;
          //myNewQuery[URUT].MasterSource := MYNEW

          //if qlistEditor.FIELDBYNAME('TABLE_NAME').ASSTRING = 'NO_FAKTUR' then

          //tipeparam =1 ini dipakai jika pada query tidak terdapat nama tabel,
          //sehingga nama tabel diambil dari param_table1
        logFile('Proses 5');
        if (myNewQuery[URUT].FindMacro('limit') <> nil) then
        begin
          myNewQuery[URUT].MacroByName('limit').Value := ' limit 1';
          try
            myNewQuery[URUT].Open;
          except
            vOpenPertama := True;
          end;
        end
        else
        begin
          myNewQuery[URUT].SQL.Add(' LIMIT 1');
          try
            myNewQuery[URUT].Open;
          except
            vOpenPertama := True;
          end;
        end;
        logFile('isikan query ' + myNewQuery[URUT].Name + ' : ' + myNewQuery[URUT].SQL.Text);
        logFile('Proses 6');
        logFile('Proses tipe param ' + IntToStr(tipeparam));

        if tipeparam = 1 then
        begin
          logFile('Proses 6.1.1');
          if qlistEditor.fieldbyname('PARAM_TABLE1').asstring <> '' then
            Table_transaksi := qlistEditor.fieldbyname('PARAM_TABLE1').asstring;
          logFile('Proses 6.1.2 ' + Table_transaksi + ' jumlah macros ' + IntToStr(myNewQuery[URUT].MacroCount));
          if myNewQuery[URUT].MacroCount > 0 then
            myNewQuery[URUT].Macros[0].Value := ' ' + Table_transaksi + ' a ';
          logFile('Proses 6.1.3');
          try
            myNewQuery[URUT].Open;
          except
            vOpenPertama := True;
          end;
          urut_trans := URUT;
        end;

        logFile('Proses 6.1');
        myNewQuery[URUT].MacroByName('KONDISI2').Value := '';
        if myNewQuery[URUT].Findmacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            myNewQuery[URUT].MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            myNewQuery[URUT].MacroByName('kondisiUser').Value := ' ';

        end;
        logFile('Proses 6.2');
        if myNewQuery[URUT].findparam('USER') <> nil then
          myNewQuery[URUT].paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        logFile('Proses 6.3');
        if myNewQuery[URUT].Findmacro('KondisiGudang') <> nil then
        begin
          if NOT_GUDANG_all_flag then
            myNewQuery[URUT].MacroByName('KondisiGudang').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            myNewQuery[URUT].MacroByName('KondisiGudang').Value := ' ';
        end;
        logFile('Proses 6.4');
        if myNewQuery[URUT].FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            myNewQuery[URUT].MacroByName('kondisiDivisi').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            myNewQuery[URUT].MacroByName('kondisiDivisi').Value := ' ';
        end;
        logFile('Proses 6.5');
        if myNewQuery[URUT].FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            myNewQuery[URUT].MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            myNewQuery[URUT].MacroByName('kondisiMerek').Value := ' ';
        end;
        logFile('Proses 6.6');
        if myNewQuery[URUT].FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            myNewQuery[URUT].MacroByName('kondisiCabang').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            myNewQuery[URUT].MacroByName('kondisiCabang').Value := ' ';
        end;
        logFile('Proses 7');

        namaparam1 := pisahdelimeter(qlistEditor.FIELDBYNAME('COMBO_TITLES').ASSTRING, '|', jum_param1);
        isinamaparam1[URUT] := pisahdelimeter(qlistEditor.FIELDBYNAME('COMBO_FIELDS').ASSTRING, '|', jum_param1);


         {if (tipeparam <> 1) AND (
           ( UPPERCASE(isinamaparam1[urut][0])  = 'PERIODE') OR
           ( UPPERCASE(isinamaparam1[urut][0])  = 'WAREHOUSE_NAME')  OR
           ( UPPERCASE(isinamaparam1[urut][0])  = 'FIRST_NAME')  OR
           ( UPPERCASE(isinamaparam1[urut][0])  = 'ERP_USER_NAME')  OR
           ( UPPERCASE(isinamaparam1[urut][0])  = 'ACCOUNT_NAME')  OR
           ( UPPERCASE(isinamaparam1[urut][0])  = 'DISPLAY_NAME')

           )

            then
          begin
           try

            myNewQuery[urut].Active := true;
          except
            messagedlg('maaf, untuk parameter PERTAMA DI '+labelh[urut].caption+ ' maka TABEL TRANSAKSI harus diisi !',mtInformation, [mbok],0);
          end;
        end
        ELSE
          dblookuph[urut].OnEnter  := ComboNo_dokEnter ;
          }

        if myNewQuery[URUT].findparam('ERP_GROUP_ID') <> nil then
        begin
          myNewQuery[URUT].parambyname('ERP_GROUP_ID').asstring := Group_id;
        end;

       // maret 2021 ini sy remark agar no nota jg bs langsung saja  di aktifkan
       // jika tdk kita tdk tau
        if (tipeparam <> 1) then
        begin
          logFile('Proses 8');
          qErpDetail.First;
          logFile('WHERE_VALUE ' + qErpDetailWHERE_VALUE.AsString + ' WHERE_COLUMN ' + qErpDetailWHERE_COLUMN.AsString);
          if (qErpDetailWHERE_VALUE.AsString <> '') and (qErpDetailWHERE_COLUMN.AsString <> '') then
          begin
            vFieldName := qErpDetailWHERE_COLUMN.AsString;
            if (vFieldName <> '') and not vOpenPertama then
            begin
              if myNewQuery[URUT].FindField(vFieldName) <> nil then
              begin
                if myNewQuery[URUT].FieldByName(vFieldName).Origin <> '' then
                  vFieldName := myNewQuery[URUT].FieldByName(vFieldName).Origin
              end
              else
              begin
                logFile('maaf kolom ' + vFieldName + ' tidak ada di erp_rpt_detail_id=' + qErpDetailERP_DETAIL_ID.AsString);
                logFile(myNewQuery[URUT].FinalSQL);
                vError := 'maaf kolom ' + vFieldName + ' tidak ada di erp_rpt_detail_id=' + qErpDetailERP_DETAIL_ID.AsString;
                Exit;
              end;
            end;
            if vFieldName <> '' then
            begin
              if myNewQuery[URUT].FindMacro('KONDISI') <> nil then
              begin
                myNewQuery[URUT].MacroByName('KONDISI').Value := ' AND ' + vFieldName + ' = ' + QuotedStr(qErpDetailWHERE_VALUE.AsString);
                try
                  myNewQuery[URUT].Active := true;
                except
                  logFile('maaf, untuk parameter PERTAMA DI ' + qlistEditor.fieldbyname('DISPLAY_FORM').AsString + ' maka TABEL TRANSAKSI harus diisi !');
                  vError := 'maaf, untuk parameter PERTAMA DI ' + qlistEditor.fieldbyname('DISPLAY_FORM').AsString + ' maka TABEL TRANSAKSI harus diisi !';
                  Exit;
                end;
              end
              else
              begin
                logFile('maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' macro kondisi tidak ada!');
                vError := 'maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' macro kondisi tidak ada!';
                Exit;
              end;
            end;
          end;
        end;

        logFile('Proses 9');
        namaparam1 := pisahdelimeter(qlistEditor.fieldbyname('COMBO_TITLES').ASSTRING, '|', jum_param1);
        isinamaparam1[URUT] := pisahdelimeter(qlistEditor.fieldbyname('COMBO_FIELDS').ASSTRING, '|', jum_param1);
        panjangnamaparam1 := pisahdelimeter(qlistEditor.fieldbyname('COMBO_WIDTHS').ASSTRING, '|', jum_param1);

        if jumlahparameter = 2 then
        begin
          logFile('Proses 10');
          myNewQuery2[URUT] := TMyQuery.Create(Self);
          myNewQuery2[URUT].Connection := DB1;
          myNewQuery2[URUT].Tag := URUT;
          myNewQuery2[URUT].Name := 'qPrm' + QListEditor.fieldbyname('ERP_RPT_PARAM_ID').AsString + '2';
          //myNewQuery2[urut].CachedUpdates := true;

          myNewQuery2[URUT].SQL.Add(s);
          if (myNewQuery2[URUT].FindMacro('limit') <> nil) then
          begin
            myNewQuery2[URUT].MacroByName('limit').Value := ' limit 1';
            try
              myNewQuery2[URUT].Open;
            except
              vOpenKedua := True;
            end;
          end
          else
          begin
            myNewQuery2[URUT].SQL.Add(' LIMIT 1');
            try
              myNewQuery2[URUT].Open;
            except
              vOpenKedua := True;
            end;
          end;

          logFile('Proses 11');
          if tipeparam = 1 then
          begin
            myNewQuery2[URUT].Macros[0].Value := ' ' + Table_transaksi + ' a ';
            try
              myNewQuery2[URUT].Open;
            except
              vOpenPertama := True;
            end;
          end;
          myNewQuery2[URUT].MacroByName('KONDISI2').Value := '';

          if myNewQuery2[URUT].findparam('ERP_GROUP_ID') <> nil then
            myNewQuery2[URUT].parambyname('ERP_GROUP_ID').asstring := Group_id;

          if myNewQuery2[URUT].findparam('USER') <> nil then
            myNewQuery2[URUT].paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
          if myNewQuery2[URUT].Findmacro('KondisiGudang') <> nil then
          begin
            if NOT_GUDANG_all_flag then
              myNewQuery2[URUT].MacroByName('KondisiGudang').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
            else
              myNewQuery2[URUT].MacroByName('KondisiGudang').Value := ' ';
          end;
          if myNewQuery2[URUT].FindMacro('kondisiDivisi') <> nil then
          begin
            if NOT_DIVISI_all_flag then
              myNewQuery2[URUT].MacroByName('kondisiDivisi').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
            else
              myNewQuery2[URUT].MacroByName('kondisiDivisi').Value := ' ';
          end;
          if myNewQuery2[URUT].FindMacro('kondisiMerek') <> nil then
          begin
            if NOT_MEREK_all_flag then
              myNewQuery2[URUT].MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
            else
              myNewQuery2[URUT].MacroByName('kondisiMerek').Value := ' ';
          end;
          if myNewQuery2[URUT].FindMacro('kondisiCabang') <> nil then
          begin
            if branch_id <> '1001' then
              myNewQuery2[URUT].MacroByName('kondisiCabang').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
            else
              myNewQuery2[URUT].MacroByName('kondisiCabang').Value := ' ';
          end;
          logFile('Proses 12');

          if (tipeparam <> 1) then
          begin
            if myNewQuery[URUT].FindMacro('KONDISI') <> nil then
            begin
              if qErpDetail.RecordCount = 2 then
              begin
                qErpDetail.Last;
                if (qErpDetailWHERE_COLUMN.AsString <> '') and (qErpDetailWHERE_VALUE.AsString <> '') then
                begin
                  vFieldName := qErpDetailWHERE_COLUMN.AsString;
                  if (vFieldName <> '') and not vOpenKedua then
                  begin
                    if myNewQuery2[URUT].FindField(vFieldName) <> nil then
                    begin
                      if myNewQuery2[URUT].FieldByName(vFieldName).Origin <> '' then
                        vFieldName := myNewQuery2[URUT].FieldByName(vFieldName).Origin
                    end
                    else
                    begin
                      logFile('maaf kolom ' + vFieldName + ' tidak ada di erp_rpt_detail_id=' + qErpDetailERP_DETAIL_ID.AsString);
                      vError := 'maaf kolom ' + vFieldName + ' tidak ada di erp_rpt_detail_id=' + qErpDetailERP_DETAIL_ID.AsString;
                      Exit;
                    end;
                  end;
                  if vFieldName <> '' then
                  begin
                    myNewQuery2[URUT].MacroByName('KONDISI').Value := ' AND ' + vFieldName + ' = ' + QuotedStr(qErpDetailWHERE_VALUE.AsString);
                    try
                      myNewQuery2[URUT].Active := true;
                    except
                      logFile('maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' maka TABEL TRANSAKSI harus diisi !');
                      vError := 'maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' maka TABEL TRANSAKSI harus diisi !';
                      Exit;
                    end;
                  end;
                end;
              end
              else
              begin
                logFile('maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' json tidak ada!');
                vError := 'maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' json tidak ada!';
                Exit;
              end;
            end
            else
            begin
              logFile('maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' macro kondisi tidak ada!');
              vError := 'maaf, untuk parameter KEDUA DI ' + qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING + ' macro kondisi tidak ada!';
              Exit;
            end;
          end;
          logFile('Proses 13');

          panjangnamaparam1 := pisahdelimeter(qlistEditor.FIELDBYNAME('COMBO_WIDTHS').ASSTRING, '|', jum_param1);
        end;

      end
      else if tipeparam = 5 then //untuk tanggal
      begin
        logFile('Proses 14');
        if not tgl_pertama then
        begin
          urut_tgl := URUT;
          tgl_pertama := true;
        end;

        if jumlahparameter = 2 then
        begin            //if (ALLITEM = 1) AND (DISPLAY_VALUE_FLAG[URUT]) then dblookuptanggal2[urut].Enabled := False;
        end;
      end;

              // fatra nov 2020
      logFile('Proses 15');
      R := 1;
      while true do
      begin
        logFile('Proses 16');
        if oldNILAI_nama_field[R] = qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING then
        begin
          if ((tipeparam = 1) or (tipeparam = 4) or (tipeparam = 3)) and (oldNILAI_satu[R] <> '') then //combo box biasa
          begin
            if (qlistEditor.FIELDBYNAME('DISPLAY_FORM').ASSTRING <> 'Periode') then
            begin
              if (nama_id[URUT] <> 'DISPLAY_NAME') and (myNewQuery[URUT].findfield(nama_id[URUT]) <> nil) and (nama_id[URUT] <> 'STATUS_ID') then
              begin
                logFile('Proses 17');

                try
                  myNewQuery[URUT].close;
                          // 01 OKT 2021 INI RAWAN BUG ...krn bisa jd ada field yg ambiguous
                  if (UPPERCASE(nama_id[URUT]) = 'WAREHOUSE_ID') or (UPPERCASE(nama_id[URUT]) = 'COA_ID') then
                    s := ' and a.' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_satu_id[R])
                  else if (UPPERCASE(nama_id[URUT]) = 'ITEM_ID') then
                    s := ' and i.' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_satu_id[R])
                  else
                    s := ' and ' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_satu_id[R]);

                  myNewQuery[URUT].macrobyname('kondisi2').Value := s;
                  myNewQuery[URUT].open;
                except
                  on E: exception do
                  begin
                    logFile('open ' + myNewQuery[URUT].Name + ' Bermasalah error ' + E.Message + '!');
                    vError := 'open ' + myNewQuery[URUT].Name + ' Bermasalah error ' + E.Message + '!';
                    Exit;
                  end;
                end;
              end
              else
              begin
              end;
            end; // END PERIODE
            if (oldNILAI_dua_id[R] <> '') then
            begin
              logFile('Proses 18');
              if myNewQuery2[URUT] = nil then
              begin
                logFile('myNewQuery2[' + IntToStr(URUT) + '] adalah nil, skip');
              end
              else
              begin
                try
                  myNewQuery2[URUT].close;
                          // 01 OKT 2021 INI RAWAN BUG ...krn bisa jd ada field yg ambiguous
                  if (UPPERCASE(nama_id[URUT]) = 'WAREHOUSE_ID') or (UPPERCASE(nama_id[URUT]) = 'COA_ID') then
                    s := ' and a.' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_dua_id[R])
                  else if (UPPERCASE(nama_id[URUT]) = 'ITEM_ID') then
                    s := ' and i.' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_dua_id[R])
                  else
                    s := ' and ' + nama_id[URUT] + ' = ' + quotedstr(oldNILAI_dua_id[R]);

                  myNewQuery2[URUT].macrobyname('kondisi2').Value := s;
                  myNewQuery2[URUT].open;
                except
                  on E: exception do
                  begin
                    logFile('open ' + myNewQuery2[URUT].Name + ' Bermasalah error ' + E.Message + '!');
                    vError := 'open ' + myNewQuery2[URUT].Name + ' Bermasalah error ' + E.Message + '!';
                    Exit;
                  end;
                end;

              end;
            end;

          end;
          BREAK;
                     //dblookuph[urut].
                     //dblookupd[urut].name.value := oldNILAI_dua[i];
        end
        else
        begin
          R := R + 1;
          if R > panjang_posisi then
            BREAK;
        end;

      end;

      inc(i);
    end;
    logFile(' selesai filter erp_detail_id ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString);
    qlistEditor.next;
  end;

 //Bevel1.height := (I-1)*jarakbutton+TABTOP;
 //label3.caption := inttostr(maxwidth+23);
 //Bevel1.Width := 760;
 //Bevel1.Width := maxwidth+23;
 //exit;
  if i > 1 then
  begin
   //PanelControl.Top := (I)*jarakbutton+TABTOP;
   //Richedit1.Top := (I)*jarakbutton+TABTOP;
    INC(i);
   {BDesign.top := (I-1)*jarakbutton+TABTOP;
   BFastReport.top := (I-1)*jarakbutton+TABTOP;
   BTemplate.top := (I-1)*jarakbutton+TABTOP;
   Preview.top := (I-1)*jarakbutton+TABTOP;
   BitBtn1.top := (I-1)*jarakbutton+TABTOP;
   Button1.top := (I-1)*jarakbutton+TABTOP;}

    inc(i, 2);
   //XREPORT.width := 822;
   //XReport.width := Bevel1.Width+ 62;
   //XREPORT.Height := (I+1)*jarakbutton+TABTOP + 37;
//   XReport.Position := poDesktopCenter;
  // XREPORT.visible := true;
  end
  else
  begin
  end;
  Result := True;
end;

function TsvReport.rubahsqlsebelumexecute(no_urut: integer; awal: string; var ada_sql1: boolean): string;
var
  S: string;
  ada_Group_id, ada_Group_id1, Ada_Department_id: boolean;
  dbl: string;
begin
  dbl := '''';
  S := awal;
  ada_Group_id1 := AnsiContainsStr(S, Tgroup_id1);
  if ada_Group_id1 then
    S := AnsiReplaceStr(S, Tgroup_id1, ' AND ' + AliasGroup1 + '.erp_group_id = ' + dbl + Group_id + dbl);

  ada_Group_id := AnsiContainsStr(S, Tgroup_id);
  if ada_Group_id then
    S := AnsiReplaceStr(S, Tgroup_id, ' AND ' + AliasGroup + '.erp_group_id = ' + dbl + Group_id + dbl);

  ada_sql1 := AnsiContainsStr(S, LABEL_SOURCE[no_urut]);
  if ada_sql1 then
  begin
    QERP_SQL.Active := FALSE;
    QERP_SQL.ParamByName('NO_URUT').AsString := INTTOSTR(MASTER_SOURCE[no_urut]);
    QERP_SQL.Active := TRUE;
    S := AnsiReplaceStr(S, LABEL_SOURCE[no_urut], QERP_SQL.fieldbyname('str_sql').asstring);
    QERP_SQL.Active := FALSE;
  end;

  result := S;
end;

//DIJALANKAN PADA WAKTU AWAL PEMBUATAN CREATE
function TsvReport.Define_Sql: Boolean;
var
  i, j: byte;
  NO_URUT: BYTE;
  ket, s: string;
  list: Integer;
      //jumlah : integer;
  masuk: boolean;
begin
  try
    Result := True;
    qSql.Execute;
    //qErp_stok.execute;
    FILLCHAR(PARAM_YG_DIAMBIL, sizeof(PARAM_YG_DIAMBIL), 0);
    FILLCHAR(JUM_PARAM_YG_DIAMBIL, sizeof(JUM_PARAM_YG_DIAMBIL), 0);
    FILLCHAR(LABEL_SOURCE, sizeof(LABEL_SOURCE), 0);
    if not qSql.eof then
    begin
      satu_sql := qsql.RecordCount = 1;
      kondisi_and := true;
      for i := 1 to qsql.RecordCount do
      begin
        NO_URUT := qsql.fieldbyname('NO_URUT').asinteger;

        MASTER_SOURCE[NO_URUT] := qsql.FIELDBYNAME('MASTER_SOURCE').asINTEGER;
        LABEL_SOURCE[NO_URUT] := qsql.FIELDBYNAME('LABEL_SOURCE').asstring;
        try
          BANYAKBARIS[NO_URUT] := qsql.fieldbyname('COUNT_COLUMN').asinteger;
        except
          BANYAKBARIS[NO_URUT] := 1;
        end;
        PARAM_YG_DIAMBIL[NO_URUT] := pisahdelimeter(qsql.FIELDBYNAME('PARAM_YG_DIAMBIL').ASSTRING, '|', JUM_PARAM_YG_DIAMBIL[NO_URUT]);
        for list := 0 to Length(PARAM_YG_DIAMBIL[NO_URUT]) - 1 do
        begin
          logFile('list parameter urut ' + IntToStr(NO_URUT) + ' id ' + qSql.fieldbyname('ERP_RPT_SQL_ID').AsString + ' : ' + PARAM_YG_DIAMBIL[NO_URUT][list]);
        end;

        PARENT_SQL[NO_URUT] := qsql.FIELDBYNAME('PARENT_SQL').asINTEGER;
          //Kondisiyangdiambil[NO_URUT] := qsql.FIELDBYNAME('KONDISI').asstring;

        memo1.Clear;
        memo1.Text := qsql.FIELDBYNAME('KONDISI').AsString;

        fillchar(Kondisiyangdiambil[NO_URUT], sizeof(Kondisiyangdiambil[NO_URUT]), 0);
        fillchar(paramkondisiyangdiambil[NO_URUT], sizeof(paramkondisiyangdiambil[NO_URUT]), 0);
        setlength(Kondisiyangdiambil[NO_URUT], maxcombo);
        setlength(paramkondisiyangdiambil[NO_URUT], maxcombo);
        j := 0;
        masuk := true;
        jum_kondisi_yg_diambil[NO_URUT] := 0;
        try
          while (j < memo1.Count) and (masuk) do
          begin
            ket := memo1.Strings[j];
//            logFile(memo1.Strings[j]);
            masuk := pisahdelimeterkhususkondisi(ket, '=', paramKondisiyangdiambil[NO_URUT], Kondisiyangdiambil[NO_URUT], jum_kondisi_yg_diambil[NO_URUT]);
            inc(j);
          end;
        except
          on E: exception do
          begin
            logFile(E.Message);
            Result := False;
            Exit;
          end;
        end;

        case NO_URUT of
          0: // isi header
            begin
              Langsung := True;
              qheader.SQL.Clear;
              qheader.SQL.add(qsql.fieldbyname('str_sql').AsString);
              //DSheader.DataSource := dsHeader;

              s := qHeader.SQL.Text;

              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);

              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Kata-kata ''' + LABEL_SOURCE[NO_URUT] + ''' tidak ditemukan pada isi query QHeader';
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] = '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) query QHeader tidak mengganti apa-apa padahal master_source berisi ' + inttostr(master_source[NO_URUT]);
                inc(no_error);
              end;
              if (parent_sql[NO_URUT] = NO_URUT) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Isi Parent_Sql pada qHeader(' + inttostr(NO_URUT) + ') tidak boleh diisi dengan ' + inttostr(NO_URUT);
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] = 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (HINT) Isi Label_Source di query QHeader yaitu ''' + LABEL_SOURCE[NO_URUT] + ''' sebaiknya dihilangkan';
                inc(no_error);
              end;

              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qHeader.sql.clear;
              qHeader.sql.add(s);
              //RichEdit1.Lines.Add(s);
//              logFile('proses isi sql qheader');
            end;
          1: // isi q1
            begin
              q1.SQL.Clear;
              q1.SQL.add(qsql.fieldbyname('str_sql').AsString);
              //DS1.DataSource := ds1;

              s := q1.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q1.sql.clear;
              q1.sql.add(s);
//              logFile('proses isi sql q1');
            end;
          2: // isi q2
            begin
              q2.SQL.Clear;
              q2.SQL.add(qsql.fieldbyname('str_sql').AsString);
              //DS2.DataSource := ds2;

              s := q2.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q2.sql.clear;
              q2.sql.add(s);
//              logFile('proses isi sql q2');
            end;
          3: // isi qdetail
            begin
              qdetail.SQL.Clear;
              qdetail.SQL.add(qsql.fieldbyname('str_sql').AsString);
              //DSdetail.DataSource := ddetail;

              s := qDetail.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. Pada isi query QDetail tidak terdapat kata-kata ' + LABEL_SOURCE[NO_URUT];
                inc(no_error);
              end;
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Kata-kata ''' + LABEL_SOURCE[NO_URUT] + ''' tidak ditemukan pada isi query qDetail';
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] = '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) query qDetail tidak mengganti apa-apa padahal master_source berisi ' + inttostr(master_source[NO_URUT]);
                inc(no_error);
              end;
              if (parent_sql[NO_URUT] = NO_URUT) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Isi Parent_Sql pada qDetail(' + inttostr(NO_URUT) + ') tidak boleh diisi dengan ' + inttostr(NO_URUT);
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] = 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (HINT) Isi Label_Source di query qDetail yaitu ''' + LABEL_SOURCE[NO_URUT] + ''' sebaiknya dihilangkan';
                inc(no_error);
              end;
              qDetail.sql.clear;
              qDetail.sql.add(s);
            end;
          4: // isi no urut
            begin
              qNo_form.SQL.Clear;
              qNo_form.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSNo_form.DataSource := dsNo_Form;

              s := qNo_Form.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qNo_Form.sql.clear;
              qNo_Form.sql.add(s);
            end;
          5: // isi q5
            begin
              q5.SQL.Clear;
              q5.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DS5.DataSource := ds5;

              s := q5.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q5.sql.clear;
              q5.sql.add(s);
            end;
          6: // isi q6
            begin
              q6.SQL.Clear;
              q6.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DS6.DataSource := ds6;

              s := q6.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q6.sql.clear;
              q6.sql.add(s);
            end;
          7: // isi q7
            begin
              q7.SQL.Clear;
              q7.SQL.add(qsql.fieldbyname('str_sql').AsString);
              //DS7.DataSource := ds7;

              s := q7.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q7.sql.clear;
              q7.sql.add(s);
            end;
          8: // isi q8
            begin
              q8.SQL.Clear;
              q8.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DS8.DataSource := ds8;

              s := q8.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q8.sql.clear;
              q8.sql.add(s);
            end;
          9: // isi Group
            begin
              qGroup.SQL.Clear;
              qGroup.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DSGroup.DataSource := dsGroup;

              s := qGroup.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Kata-kata ''' + LABEL_SOURCE[NO_URUT] + ''' tidak ditemukan pada isi query qGroup';
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] = '') and (master_source[NO_URUT] <> 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) query qGroup tidak mengganti apa-apa padahal master_source berisi ' + inttostr(master_source[NO_URUT]);
                inc(no_error);
              end;
              if (parent_sql[NO_URUT] = NO_URUT) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (WARNING) Isi Parent_Sql pada qGroup(' + inttostr(NO_URUT) + ') tidak boleh diisi dengan ' + inttostr(NO_URUT);
                inc(no_error);
              end;
              if (not ada_sql1) and (LABEL_SOURCE[NO_URUT] <> '') and (master_source[NO_URUT] = 0) then
              begin
                error_kesalahan := error_kesalahan + #13 + #10 + inttostr(no_error) + '. (HINT) Isi Label_Source di query qGroup yaitu ''' + LABEL_SOURCE[NO_URUT] + ''' sebaiknya dihilangkan';
                inc(no_error);
              end;
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qGroup.sql.clear;
              qGroup.sql.add(s);
            end;
          10: // isi q3
            begin
              q3.SQL.Clear;
              q3.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DS3.DataSource := ds3;

              s := q3.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q3.sql.clear;
              q3.sql.add(s);
            end;
          11: // isi q4
            begin
              q4.SQL.Clear;
              q4.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DS4.DataSource := ds4;

              s := q4.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q4.sql.clear;
              q4.sql.add(s);
            end;
          12: // isi q9
            begin
              q9.SQL.Clear;
              q9.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DS9.DataSource := ds9;

              s := q9.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q9.sql.clear;
              q9.sql.add(s);
            end;
          13: // isi qDetail1
            begin
              qDetail1.SQL.Clear;
              qDetail1.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DSDetail1.DataSource := dsDetail1;

              s := qDetail1.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qDetail1.sql.clear;
              qDetail1.sql.add(s);
            end;
          14: // isi qDetail2
            begin
              qDetail2.SQL.Clear;
              qDetail2.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DSDetail2.DataSource := dsDetail2;

              s := qDetail2.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qDetail2.sql.clear;
              qDetail2.sql.add(s);
            end;
          15: // isi qDetail3
            begin
              qDetail3.SQL.Clear;
              qDetail3.SQL.add(qsql.fieldbyname('str_sql').AsString);
             // DSDetail3.DataSource := dsDetail3;

              s := qDetail3.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qDetail3.sql.clear;
              qDetail3.sql.add(s);
            end;
          16: // isi qDetail4
            begin
              qDetail4.SQL.Clear;
              qDetail4.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := qDetail4.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              qDetail4.sql.clear;
              qDetail4.sql.add(s);
            end;
          17: // isi q10
            begin
              q10.SQL.Clear;
              q10.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q10.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q10.sql.clear;
              q10.sql.add(s);
            end;
          18: // isi q11
            begin
              q11.SQL.Clear;
              q11.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q11.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q11.sql.clear;
              q11.sql.add(s);
            end;
          19: // isi q12
            begin
              q12.SQL.Clear;
              q12.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q12.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q12.sql.clear;
              q12.sql.add(s);
            end;
          20: // isi q14
            begin
              q14.SQL.Clear;
              q14.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q14.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q14.sql.clear;
              q14.sql.add(s);
            end;
          21: // isi q15
            begin
              q15.SQL.Clear;
              q15.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q15.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q15.sql.clear;
              q15.sql.add(s);
            end;
          22: // isi q16
            begin
              q16.SQL.Clear;
              q16.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q16.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q16.sql.clear;
              q16.sql.add(s);
            end;
          23: // isi q17
            begin
              q17.SQL.Clear;
              q17.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q17.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q17.sql.clear;
              q17.sql.add(s);
            end;
          24: // isi q18
            begin
              q18.SQL.Clear;
              q18.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DSDetail4.DataSource := dsDetail4;

              s := q18.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q18.sql.clear;
              q18.sql.add(s);
            end;
          25: // isi qq19
            begin
              q19.SQL.Clear;
              q19.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DS19.DataSource := ds19;

              s := q19.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q19.sql.clear;
              q19.sql.add(s);
            end;
          26: // isi q20
            begin
              q20.SQL.Clear;
              q20.SQL.add(qsql.fieldbyname('str_sql').AsString);
            //  DS20.DataSource := ds20;

              s := q20.SQL.Text;
              s := rubahsqlsebelumexecute(NO_URUT, s, ada_sql1);
              j := 0;
              while j < jum_kondisi_yg_diambil[NO_URUT] do
              begin
                s := AnsiReplaceStr(s, paramKondisiyangdiambil[NO_URUT][j], Kondisiyangdiambil[NO_URUT][j]);
                inc(j);
              end;
              q20.sql.clear;
              q20.sql.add(s);
            end;
        end;

        qSql.next;
      end;
    end;
    logFile(error_kesalahan);
  except
    on E: exception do
    begin
      logFile('qSql open Error ' + E.Message);
      Result := False;
      Exit;
    end;
  end;
end;

function TsvReport.Run_sql(vstr: string): Boolean;
begin
  Result := False;
  sql1.SQL.Clear;
  sql1.SQL.Add(vstr);
  try
    sql1.Execute;
    Result := True;
  except
    vErr := '';
    logFile('proses update ' + vstr + ' Bermasalah error ' + vErr + '!');
  end;
end;

procedure TsvReport.inisialisasi_v; //untuk menginisialisasi setiap variable yang akan diparsingkan
begin
  fillchar(v_a_param_yg_dipasing, sizeof(v_a_param_yg_dipasing), 0);
  PERIOD_LALU := '';
  PERIOD_AWAL := '';

  PERIOD1 := '';
  PERIOD2 := '';
  yearmonth1 := '';
  yearmonth2 := '';
  YEAR1 := '';
  YEAR2 := '';
  year_lalu := '';
end;

function TsvReport.filter_Period_awal(PERIOD: string): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := ' ORDER BY PERIOD_NAME ';
  qPeriod.active := true;
  qPeriod.Locate('PERIOD_NAME', PERIOD, [loCaseInsensitive]);

  if not qPeriod.eof then
    result := qPeriodPERIOD_AWAL.AsString
  else
    result := '';
end;

function TsvReport.filter_Period_Lalu(PERIOD: string): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := ' ORDER BY PERIOD_NAME ';
  qPeriod.active := true;
  qPeriod.Locate('PERIOD_NAME', PERIOD, [loCaseInsensitive]);

  qPeriod.prior;
  if not qPeriod.eof and (PERIOD <> qPeriodPERIOD_NAME.AsString) then
    result := qPeriodPERIOD_NAME.AsString
  else
    result := '';
end;

function TsvReport.filter_Period(date1: TDateTime): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriodPERIOD_NAME.AsString
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriodPERIOD_NAME.AsString;
  end;
  {qPeriod.active := false;
  qPeriod.active := true;
  if qPeriod.Locate('yearmonth',formatdatetime('yyyymm',date1),[loCaseInsensitive]) then
    result := qPeriodPERIOD_NAME.AsString
  else
  begin
    result := formatdatetime('mmm-yy',date1);
  end;}
end;

function TsvReport.filter_YEAR(date1: TDateTime): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriod.fieldbyname('YEAR').asstring
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriod.fieldbyname('YEAR').asstring;
  end;

end;

function TsvReport.filter_YEARMONTH(date1: TDateTime): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriod.fieldbyname('YEARMONTH').asstring
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriod.fieldbyname('YEARMONTH').asstring;
  end;
end;

function TsvReport.filter_YEAR_lalu(date1: TDateTime): string;
var
  start_date1: integer;
  start_date2: tdate;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
  begin
    result := qPeriod.fieldbyname('YEAR').asstring;
    start_date1 := qPeriod.fieldbyname('YEAR').asINTEGER;
    start_date2 := qPeriodSTART_DATE.AsDateTime;
  end
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    start_date1 := qPeriod.fieldbyname('YEAR').asinteger;
    start_date2 := qPeriodSTART_DATE.AsDateTime;
  end;
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + inttostr(start_date1 - 1) + formatdatetime('mm-dd', start_date2 - 1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriod.fieldbyname('YEAR').asstring
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriod.fieldbyname('YEAR').asstring;
  end;
end;

function TsvReport.filter_YEAR2(date1: TDateTime): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriod.fieldbyname('YEAR').asstring
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME DESC';
    qPeriod.execute;
    result := qPeriod.fieldbyname('YEAR').asstring;
  end;
end;

function TsvReport.filter_YEARMONTH2(date1: Tdate): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriod.fieldbyname('YEARMONTH').asstring
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriod.fieldbyname('YEARMONTH').asstring;
  end;
end;

function TsvReport.filter_Period2(date1: Tdate): string;
begin
  qPeriod.active := false;
  qPeriod.macros[0].Value := 'where start_date <= ' + db + formatdatetime('yyyy-mm-dd', date1) + db + ' and end_date >= ' + db + formatdatetime('yyyy-mm-dd', date1) + db;
  qPeriod.active := true;
  if not qPeriod.eof then
    result := qPeriodPERIOD_NAME.AsString
  else
  begin
    qPeriod.macros[0].Value := ' GROUP BY PERIOD_NAME ASC';
    qPeriod.execute;
    result := qPeriodPERIOD_NAME.AsString;
  end;
end;

function TsvReport.bulan(vbln: byte): string;
var
  nilai: string;
begin
  case vbln of
    1:
      nilai := 'JANUARI ';
    2:
      nilai := 'FEBRUARI ';
    3:
      nilai := 'MARET ';
    4:
      nilai := 'APRIL ';
    5:
      nilai := 'MEI ';
    6:
      nilai := 'JUNI ';
    7:
      nilai := 'JULI ';
    8:
      nilai := 'AGUSTUS ';
    9:
      nilai := 'SEPTEMBER ';
    10:
      nilai := 'OKTOBER ';
    11:
      nilai := 'NOVEMBER ';
    12:
      nilai := 'DESEMBER ';
  end;

  result := nilai;
end;

function TsvReport.BLN_SINGKAT(vbln: byte): string;
var
  nilai: string;
begin
  case vbln of
    1:
      nilai := 'JAN  ';
    2:
      nilai := 'FEB ';
    3:
      nilai := 'MAR ';
    4:
      nilai := 'APR ';
    5:
      nilai := 'MEI ';
    6:
      nilai := 'JUN ';
    7:
      nilai := 'JUL ';
    8:
      nilai := 'AGT ';
    9:
      nilai := 'SEP ';
    10:
      nilai := 'OKT ';
    11:
      nilai := 'NOV ';
    12:
      nilai := 'DES ';
  end;

  result := nilai;
end;

procedure TsvReport.sql1Error(Sender: TObject; E: Exception; SQL: string; var Action: TErrorAction);
begin
  vErr := E.Message;
  Action := eaFail;
end;

function TsvReport.str_filter(Harus_cek_all_flag: boolean): string;
var
  VStr, str_tambah, str_alias_baru, vmonth, param, vPnlCaption: string;
  sql_param: Tarrstring;
  i, VAR_ARRAY: Byte;
  perlu_reproses_flag, all_item_fatra, kondisi_perlu: boolean;
  pilih_allitem, allitem, tipeparam, jumlahparameter, jum_sql, j, urut: integer;
  vUrut: Integer;
  dtTgl1, dtTgl2: TDateTime;
begin
  qErpDetail.Filtered := False;
//  logFile('Proses jumlah filter ' + IntToStr(qErpDetail.RecordCount));
  for i := 1 to 12 do
    str[i] := '';
  for j := 0 to 20 do
    skondisiall[j] := '';

  try
    QListEditor.Close;
    QListEditor.ParamByName('ERP_RPT_ID').AsString := qReportWebERP_RPT_ID.AsString;
    QListEditor.Open;
  except
    on E: exception do
    begin
      logFile('QListEditor failed open ' + E.Message + '!');
      Exit;
    end;
  end;
  qlistEditor.First;
  i := 1;

//  qErpDetail.First;
//  while not qErpDetail.Eof do
//  begin
//    logFile(' ada erp_detail_id ' + qErpDetailERP_DETAIL_ID.AsString + ' column ' + qErpDetailWHERE_COLUMN.AsString + ' value ' + qErpDetailWHERE_VALUE.AsString);
//    qErpDetail.Next;
//  end;

  //mengambil semua data sesuai urutan pada tabel erp_rpt_detail
  while not qlistEditor.Eof do
  begin
    qErpDetail.Filtered := False;
    qErpDetail.Filter := 'ERP_DETAIL_ID = ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString;
    qErpDetail.Filtered := True;
//    logFile('proses query parameter filter erp_rpt_detail ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString);
    urut := qlistEditor.FieldByName('URUT').ASinteger;
    if qErpDetail.RecordCount > 0 then
    begin
//      logFile('terdeksi filter prameter erp_rpt_detail ' + QListEditor.fieldbyname('ERP_RPT_DETAIL_ID').AsString);
      param := copy(qlistEditor.FieldByName('DISPLAY_FORM').asstring, 1, 7);
      Run_query('SELECT URUT, DISPLAY_DOUBLE_FLAG, display_form, DISPLAY_ALL_FLAG,  SQL_USES, NAMA_ALIAS from  ERP_RPT_DETAIL a ' + 'INNER JOIN ERP_RPT_PARAM p ON (a.ERP_RPT_PARAM_ID = p.ERP_RPT_PARAM_ID) ' + ' WHERE   a.ERP_RPT_ID =  ' + qReportWebERP_RPT_ID.AsString + ' and  a.ERP_RPT_PARAM_ID =  ' + qlistEditor.FieldByName('ERP_RPT_PARAM_ID').asstring);

      vUrut := Query1.FieldByName('URUT').AsInteger;
      begin
        jumlahparameter := qlistEditor.FieldByName('DISPLAY_DOUBLE_FLAG').AsInteger;

        logFile('proses urut ' + IntToStr(vUrut));
        if not Query1.Eof then
        begin
          jumlahparameter := Query1.FieldByName('DISPLAY_DOUBLE_FLAG').AsInteger;
          str_alias_baru := Query1.FieldByName('NAMA_ALIAS').AsString;
          allitem := Query1.FieldByName('DISPLAY_ALL_FLAG').AsInteger;
          sql_param := pisahdelimeter(Query1.FieldByName('SQL_USES').AsString, '|', jum_sql);
          kondisi_perlu := true;
        end
        else
        begin
          allitem := 1;
          kondisi_perlu := False;
        end;

        tipeparam := QListEditor.FieldByName('PARAM_TYPE').AsInteger;
        VAR_ARRAY := QListEditor.FieldByName('VAR_ARRAY').AsInteger;

        pilih_allitem := 0;
        str[urut] := '';
        if pilih_allitem = 0 then
        begin
          str_tambah := '';
          if tipeparam = 2 then
          begin
          end
          else
            posisi := 1;
          while true do
          begin
            if oldNILAI_nama_field[posisi] = Query1.fieldbyname('display_form').AsString then
              break
            else
            begin
              if oldNILAI_nama_field[posisi] = '' then
              begin
                oldNILAI_nama_field[posisi] := Query1.FieldByName('display_form').AsString;
                panjang_posisi := panjang_posisi + 1;
              end
              else
              begin
                posisi := posisi + 1;

              end;
            end;
          end;

          if (tipeparam = 4) or (tipeparam = 1) or (tipeparam = 3) then
          begin
            if pilih_allitem = 0 then
            begin
              if (1 = 1) then //otomatis[urut]) then
              begin

              //ini untuk menunjukkan apakah nama_id pada tabel sama dengan nama_id 'CHILD_VALUE' pada paramlookup1 di erp_RPT_PARAM
              //gunanya jika sama maka mengisi 'isi_nama_id' pada field yang dibawah group item
              //jika tidak sama maka mengisi 'isi_nama_id' dengan combo biasanya
                isi_nama_id[urut] := '';
                logFile('proses field ' + nama_id[urut]);
                if myNewQuery[urut].Active then
                  logFile(myNewQuery[urut].Name + ' active')
                else
                  logFile(myNewQuery[urut].Name + ' false');
                if myNewQuery[urut].findfield(nama_id[urut]) <> nil then
                begin
                  logFile('0.1.1');
             // 12 agustus 2021
                  if nama_id[urut] <> '' then
                    if (isi_nama_id[urut] = '') and (myNewQuery[urut].FIELDBYNAME(nama_id[urut]).Asstring <> '') then
                      isi_nama_id[urut] := myNewQuery[urut].FIELDBYNAME(nama_id[urut]).Asstring;
                  logFile('0.1.2');
               //else
                 //isi_nama_id[urut] := isi_id[urut];

              // akhir 12 agustus 2021

                  v_a_param_yg_dipasing[1, vUrut] := isi_nama_id[urut];
                  logFile('0.1.3');

                  oldNILAI_satu_id[posisi] := isi_nama_id[urut];
                  logFile('0.1.4');
                end;
                if isi_kata[urut] <> '' then
                  oldNILAI_satu[posisi] := isi_kata[urut]
                else if nama_id[urut] <> '' then
                  oldNILAI_satu[posisi] := myNewQuery[urut].FIELDBYNAME(nama_id[urut]).AsString;

                logFile('0.1');
                if UPPERCASE(nama_id[urut]) = 'PERIOD_NAME' then
                begin
                  PERIOD1 := isi_nama_id[urut];
                  if Period1 <> '' then
                  begin

                    try
                      STRTOINT(PERIOD1);
                      year1 := copy(Period1, 1, 4);
                      yearmonth1 := PERIOD1;
                    //copy(v_a_param_yg_dipasing[1,1],1,4)+copy(v_a_param_yg_dipasing[1,1],5,2);

                      if year1 <> '' then
                      begin
                        try
                          year_LALU := inttostr(strtoint(year1) - 1);
                        except
                        end;
                      end;
                      PERIOD_LALU := filter_Period_Lalu(Period1);
                      PERIOD_AWAL := filter_Period_awal(Period1);
                    except
                    end;
                  end;

                end; //period
                logFile('0.2');

                logFile('1');
                if not (VAR_ARRAY = AngkaNonAktif) then
                begin
                  if nama_id[urut] <> '' then
                    FArray[VAR_ARRAY] := myNewQuery[urut].FIELDBYNAME(nama_id[urut]).AsString;
                  ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
                end;

              //anu  := dblookuph[urut].value;
                all_item_fatra := false;

                if qErpDetailWHERE_VALUE.AsString = '' then
                begin
                  all_item_fatra := True;
                  logFile('2')
                end
                else

                  logFile('2.1.1');
                logFile(nama_id2[urut]);
                if (not all_item_fatra) then
                //ini untuk mengisi field dibawah group item yang tidak terdapat 'Kolom ALL'
                  if (jumlahparameter = 1) and (punyaanak[urut] <> 0) then
                  begin
                    logFile('2.1');
                    //isi_nama_id[urut] := myNewQuery[urut].FieldByName(nama_id[urut]).asstring;
                    isi_nama_id[urut] := isi_id[urut];
                    //arti dari str_alias
                    //alias dari childrennya i diposisi keberapa pada variable 'posisianak'
                    if str_alias_baru <> '' then
                      str_tambah := str_alias_baru + '.';

                    //if TAMPILKAN_KONDISI[urut] = 1 then
                    begin
                      if LOOK_PARENT[urut] = 0 then
                      begin
                        if LABEL_LOOKUP[urut] <> '' then
                        begin
                        // 12 agustus 2021
                          logFile('2.2');
                          if (isi_nama_id2[urut] = '') then
                            isi_nama_id2[urut] := myNewQuery[urut].FieldByName(label_lookup[urut]).asstring;
                          logFile('2.3');
                        end
                        else
                        begin
                          isi_nama_id2[urut] := '';
                        end;

                        str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + nama_label_id[urut] + LABEL_MK_1_3[urut] + isi_nama_id2[urut] + LABEL_MK_1_5[urut] + ' ' + operator_char[urut] + ' ' + db + isi_nama_id[urut] + db + ')'
                        //str[urut] := ' and '+LABEL_MK_1_1[URUT] +str_tambah+ nama_label_id[URUT]+LABEL_MK_1_3[URUT] + operator_char[urut]+LABEL_MK_1_5[URUT] +isi_nama_id[urut]+')'
                      end
                      else
                      begin
                        anu3 := myNewQuery[urut].FieldByName(nama_Label_id[urut]).asstring;
                        str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + anu3 + '_ID ' + LABEL_MK_1_3[urut] + LABEL_MK_1_5[urut] + operator_char[urut] + ' ' + isi_nama_id[urut] + ')';
                      end;
                    end;
                    //v_a_param_yg_dipasing[1,urut] :=isi_nama_id[urut];
                  end
                  else if (((jumlahparameter = 2) and (myNewQuery2[urut].FieldByName(nama_id[urut]).asstring = myNewQuery[urut].FieldByName(nama_id[urut]).asstring)) or ((jumlahparameter = 1) and (punyaanak[urut] = 0))) then
                  begin

                    logFile('2.4');
                    if UPPERCASE(nama_id[urut]) = 'PERIOD_NAME' then
                    begin
                      if isi_nama_id2[urut] <> '' then
                        PERIOD2 := isi_nama_id2[urut]
                      else if myNewQuery2[urut] <> nil then
                        PERIOD2 := myNewQuery2[urut].FieldByName(nama_id2[urut]).asstring;
                    end;
                    if str_alias_baru <> '' then
                      str_tambah := str_alias_baru + '.';

                    logFile('2.5');
                  //if TAMPILKAN_KONDISI[urut] = 1 then
                    begin
                      if LABEL_LOOKUP[urut] = '' then
                      begin
                        str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + nama_Label_id[urut] + LABEL_MK_1_3[urut] + LABEL_MK_1_5[urut] + operator_char[urut] + ' ' + db + isi_nama_id[urut] + db + ') ';
                      end
                      else
                      begin
                        str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + nama_Label_id[urut] + LABEL_MK_1_3[urut] + myNewQuery[urut].FieldByName(Label_LOOKUP[urut]).asstring + LABEL_MK_1_5[urut] + operator_char[urut] + ' ' + db + isi_nama_id[urut] + db + ') ';
                      end;
                    end;
                    v_a_param_yg_dipasing[2, vUrut] := isi_nama_id[urut];

                  end
                  else if (jumlahparameter = 2) and (myNewQuery[urut].FieldByName(nama_id2[urut]).AsString <> myNewQuery2[urut].FieldByName(nama_id2[urut]).AsString) then
                  begin
                    logFile('2.5.1');
                    if str_alias_baru <> '' then
                      str_tambah := str_alias_baru + '.';

                    logFile('2.6');
                    if isi_id2[urut] = '' then
                    begin
                      if not myNewQuery[urut].Active then
                        myNewQuery[urut].Open;
                      v_a_param_yg_dipasing[2, vUrut] := myNewQuery[urut].FIELDBYNAME(nama_id2[urut]).Asstring
                    end
                    else
                      v_a_param_yg_dipasing[2, vUrut] := isi_id2[urut];

                    logFile('2.7');
                   // if  (isi_nama_id[urut] = '') then
                    begin
                    //  if (myNewQuery[urut].FieldByName(nama_id2[urut]).asstring = '')   then
                     // 12 agustus 2021
                      if isi_nama_id[urut] <> '' then
                      begin
                        isi_nama_id[urut] := myNewQuery[urut].FieldByName(nama_id2[urut]).asstring;
                      end
                      else
                      begin
                        logFile('2.8');
                        isi_nama_id[urut] := myNewQuery2[urut].FieldByName(nama_id2[urut]).asstring;
                      end;
                    end;
                   // mei 2021 fatra krn jika di isi akan jadi kata2 bug buku besar smp dengan isi nya bukan id
                   //if  (isi_nama_id2[urut] = '') then

                    begin
                       // if (myNewQuery2[urut].FieldByName(nama_id2[urut]).asstring = '') then

                       // 12 agustus 2021
                      if isi_nama_id2[urut] <> '' then
                      begin
                        isi_nama_id2[urut] := myNewQuery[urut].FieldByName(nama_id2[urut]).asstring;
                      end
                      else
                      begin
                        logFile('2.9');
                        if not myNewQuery2[urut].Active then
                          myNewQuery2[urut].Open;
                        logFile('2.10');
                        isi_nama_id2[urut] := myNewQuery2[urut].FieldByName(nama_id2[urut]).asstring;
                      end;
                    end;

                    oldNILAI_dua[posisi] := isi_nama_id2[urut];
                   //dblookupd[urut].value;
                    if isi_id2[urut] <> '' then
                      oldNILAI_dua_id[posisi] := isi_id2[urut]
                    else
                      oldNILAI_dua_id[posisi] := myNewQuery2[urut].FieldByName(nama_id[urut]).asstring;

                    if VAR_ARRAY < 31 then
                    begin
                      FArray[VAR_ARRAY] := isi_nama_id[urut] + ' - ' + isi_nama_id2[urut];
                      ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
                    end;
                   //if TAMPILKAN_KONDISI[urut] = 1 then
                    begin
                      if LABEL_LOOKUP[urut] = '' then
                      begin
                      //if CHOICE_VAR[urut] = 2 then
                        begin
                          str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + nama_LABEL_ID2[urut] + ' between ' + db + isi_nama_id[urut] + db + ' and ' + db + isi_nama_id2[urut] + db + ') ';
                       //Str[urut]   := ' and '+LABEL_MK_1_1[URUT] +str_tambah+ nama_LABEL_ID2[URUT] + ' between ' +
                       //dbl + dblookuph[urut].Value + dbl+' and '+ dbl+ dblookupd[urut].Value +  dbl +') ';
                        end
                      {else
                      begin
                       Str[urut]   := ' and '+LABEL_MK_1_1[URUT] +str_tambah+ nama_ID[urut]+ ' between ' +
                        dbl + isi_nama_id[URUT] + dbl+' and '+ dbl+ isi_nama_id2[URUT] +  dbl +') ';
                      end};
                      end
                      else
                      begin
                        str[urut] := ' and ' + LABEL_MK_1_1[urut] + str_tambah + nama_Label_id2[urut] + LABEL_MK_1_3[urut] + myNewQuery[urut].FieldByName(Label_LOOKUP[urut]).asstring + LABEL_MK_1_5[urut] + ' BETWEEN ' + db + isi_nama_id[urut] + db + ' and ' + db + isi_nama_id2[urut] + db + ') ';
                      end;
                    end;
                    logFile('2.11');
                  // mei 2021 diganti di atas krn jd nya kata2 jk dgn script ini
                  //v_a_param_yg_dipasing[2,urut] :=isi_nama_id2[urut];
                    if UPPERCASE(nama_id[urut]) = 'PERIOD_NAME' then
                    begin
                      PERIOD2 := isi_nama_id2[urut];

                    end;
                  end;
                logFile('3');
              end
            end
            else
            begin
              if not (VAR_ARRAY = AngkaNonAktif) then
              begin
                FArray[VAR_ARRAY] := 'SEMUA';
                ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
              end;
             //semua item dipilih
            end;
            logFile('4');
           {else
           begin //otomatis
            if UPPERCASE(nama_id[urut]) = UPPERCASE(CHILD_VALUE[urut]) then
              isi_nama_id[urut] := myNewQuery[urut].FIELDBYNAME(CHILD_VALUE[urut]).Asstring
            else  isi_nama_id[urut] := dblookuph[urut].Value;

            v_a_param_yg_dipasing[1,i] := isi_nama_id[urut];//dblookuph[urut].Value;

            if (jumlahparameter=2) then
            begin
                 isi_nama_id2[urut] := dblookupd[urut].Value;
                 v_a_param_yg_dipasing[2,i] :=isi_nama_id2[urut];
              end;
           end;}
          end
          else if tipeparam = 5 then
          begin
            qErpDetail.First;
            qQry.SQL.Clear;
            qQry.SQL.Text := 'SELECT CAST(' + QuotedStr(qErpDetailWHERE_VALUE.AsString) + ' AS DATETIME) AS TGL';
            qQry.Open;
            dtTgl1 := qQry.Fields[0].AsDateTime;
            logFile('isilan parameter tgl 1 ' + qQry.SQL.Text);
            if qErpDetail.RecordCount = 2 then
            begin
              qErpDetail.Last;
              qQry.SQL.Clear;
              qQry.SQL.Text := 'SELECT CAST(' + QuotedStr(qErpDetailWHERE_VALUE.AsString) + ' AS DATETIME) AS TGL';
              qQry.Open;
              dtTgl2 := qQry.Fields[0].AsDateTime;
              logFile('isilan parameter tgl 2 ' + qQry.SQL.Text);
            end;
         // jumlahparameter := qlistEditor.FIELDBYNAME('DISPLAY_DOUBLE_FLAG').ASinteger;
            begin
              if str_alias_baru <> '' then
                str_tambah := str_alias_baru + '.';


            //IF (otomatis[urut]) then
              begin
                if (((jumlahparameter = 2) and (dtTgl1 = dtTgl2)) or (jumlahparameter = 1)) then
                begin
                //if TAMPILKAN_KONDISI[urut] = 1 then
                  str[urut] := ' and (DATE_FORMAT(' + str_tambah + nama_Label_id[urut] + ',' + db + '%Y-%m-%d' + db + ' ) ' + operator_char[urut] + ' ' + db +
                          //formatdatetime('dd/mm/yyyy',dblookuptanggal[urut].DATEtime)+dbl +')';
                    formatdatetime('yyyy-mm-dd', dtTgl1) + db + ')';
//                  logFile('Proses filter ' + str[urut]);
//                FArray[5] := 'TANGGAL : '+ datetostr(dblookuptanggal[urut].DATE);
                  FArray[VAR_ARRAY] := formatdatetime('dd-MMM-yyyy', dtTgl1);
                  oldNilai_TGL1 := dtTgl1;
                  Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
                end;
                if ((jumlahparameter = 2) and (dtTgl1 <> dtTgl2)) then
                begin
                //if TAMPILKAN_KONDISI[urut] = 1 then
                  oldNilai_TGL1 := dtTgl1;
                  oldNilai_TGL2 := dtTgl2;
                  str[urut] := ' and (' + str_tambah + nama_Label_id2[urut] + ' between ' + db + formatdatetime('yyyy-mm-dd hh:nn:ss', dtTgl1) + db + ' and ' + db + formatdatetime('yyyy-mm-dd hh:nn:ss', dtTgl2) + db + ') ';
                  FArray[VAR_ARRAY] := formatdatetime('dd-MMM-yyyy', dtTgl1) + ' - ' + formatdatetime('dd-MMM-yyyy', dtTgl2);
                  Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
                end;
              end;
              begin
                PERIOD1 := filter_period(dtTgl1);
                year1 := filter_YEAR(dtTgl1);
                yearmonth1 := filter_YEARMONTH(dtTgl1);
                year_LALU := filter_Year_Lalu(dtTgl1);
                PERIOD_AWAL := filter_Period_awal(Period1);
                PERIOD_LALU := filter_Period_Lalu(Period1);
                if (jumlahparameter = 2) then
                begin
                  year2 := filter_YEAR2(dtTgl2);
                  yearmonth2 := filter_YEARMONTH2(dtTgl2);
                  PERIOD2 := filter_period2(dtTgl2);
                end;

                v_a_param_yg_dipasing[1, vUrut] := formatdatetime('yyyy-mm-dd hh:nn:ss', dtTgl1); //datetostr(dblookuptanggal[urut].date);
                if copy(v_a_param_yg_dipasing[1, i], 12, 5) = '00:00' then
                  v_a_param_yg_dipasing[1, vUrut] := formatdatetime('yyyy-mm-dd', dtTgl1); //datetostr(dblookuptanggal[urut].date);


                if (jumlahparameter = 2) then
                  v_a_param_yg_dipasing[2, vUrut] := formatdatetime('yyyy-mm-dd hh:mm:ss', dtTgl2);

              end
            end;
          end;
        end
        else if VAR_ARRAY < 50 then
        begin

          FArray[VAR_ARRAY] := 'SEMUA';
          Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
        end;

        if kondisi_perlu then
        begin
          for j := 0 to jum_sql - 1 do
            skondisiall[strtoint(sql_param[j])] := skondisiall[strtoint(sql_param[j])] + str[urut];
        end;

      end; // end visible label


      logFile('5');
      vmonth := COPY(PERIOD1, 5, 2);
      if vmonth <> '' then
      begin
        VAR_ARRAY := 15;
        try
          STRTOINT(vmonth);
          FArray[VAR_ARRAY] := BULAN(STRTOINT(vmonth));
          Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];

          VAR_ARRAY := 16;
          FArray[VAR_ARRAY] := BLN_SINGKAT(STRTOINT(vmonth));
          Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];

        except
        end;
      end;

      vmonth := COPY(PERIOD_LALU, 5, 2);
      if vmonth <> '' then
      begin
        try
          STRTOINT(vmonth);

          VAR_ARRAY := 17;
          FArray[VAR_ARRAY] := BULAN(STRTOINT(vmonth));
          Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];

          VAR_ARRAY := 18;
          FArray[VAR_ARRAY] := BLN_SINGKAT(STRTOINT(vmonth));
          Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
        except
        end;
      end;
      logFile('6');

      if PERIOD1 <> '' then
      begin
        VAR_ARRAY := 11;
        FArray[VAR_ARRAY] := trim(FArray[15]) + '-' + COPY(PERIOD1, 1, 4);

        if PERIOD2 <> PERIOD1 then
        begin
          vmonth := COPY(PERIOD2, 5, 2);
          if vmonth <> '' then
          begin
            try
              STRTOINT(vmonth);
              FArray[VAR_ARRAY] := FArray[VAR_ARRAY] + ' - ' + trim(BULAN(STRTOINT(vmonth))) + '-' + COPY(PERIOD2, 1, 4);
            except
            end;
          end;
        end;
        Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];

      end;

      logFile('7');
      VAR_ARRAY := 21;
      FArray[VAR_ARRAY] := PERIOD1;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 22;
      FArray[VAR_ARRAY] := PERIOD2;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 23;
      FArray[VAR_ARRAY] := YEAR1;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 24;
      FArray[VAR_ARRAY] := YEAR2;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 25;
      FArray[VAR_ARRAY] := PERIOD_LALU;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 26;
      FArray[VAR_ARRAY] := '0';
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
      VAR_ARRAY := 30;
      FArray[VAR_ARRAY] := YEAR_LALU;
      Ar1[VAR_ARRAY] := FArray[VAR_ARRAY];
    end;
    logFile('kondisi urutan ' + IntToStr(urut) + ' ' + str[urut]);
    inc(i);
    qlistEditor.next;
  end;
  Str_Filter := VStr;
end;

function TsvReport.Update_coa(VPeriod, vcoa_id: string; vsaldo, vdebet, vkredit: real): Boolean;
begin
  Result := False;
  try
    qUpdate_Coa_Balance.Parambyname('COA_SALDO').asfloat := vsaldo;
    qUpdate_Coa_Balance.Parambyname('COA_DEBET').asfloat := vdebet;
    qUpdate_Coa_Balance.Parambyname('COA_CREDIT').asfloat := vkredit;
    qUpdate_Coa_Balance.Parambyname('COA_ID').AsString := vcoa_id;
    qUpdate_Coa_Balance.Parambyname('PERIOD_NAME').AsString := VPeriod;
    qUpdate_Coa_Balance.Execute;
  except
    on E: exception do
    begin
      logFile('Error qUpdate_Coa_Balance  ' + E.Message);
      Exit;
    end;
  end;
  Result := True;
end;

function TsvReport.Proses_Coa_Balance(VPeriod: string; VProgress: byte): Boolean;
var
  jml_rec: integer;
begin
  Result := False;
    // isi dari all transaksi
  if not run_sql(' UPDATE COA_BALANCE SET  COA_DEBET =  0,  COA_CREDIT = 0 WHERE PERIOD_NAME = ' + VPeriod) then
    Exit;
  if not Run_query('SELECT * FROM ERP_SQL WHERE TRANS_NO = 77') then
    Exit;
  qQry.SQL.Clear;
  qQry.SQL.add('select coa_id, DOCUMENT_NO,  DOCUMENT_CLASS_CODE,sum(debet * ifnull(kurs,1)) as debet, sum(kredit*ifnull(kurs,1)) as kredit from (' + Query1.fieldbyname('str_sql').AsString + ' ) as xx ' + ' group by coa_id ');
  qQry.MacroByName('x1').Value := ' and False';
  qQry.MacroByName('P1').Value := ' and (a.Period_name = ' + VPeriod + ')';
  try
    qQry.Active := True;
  except
    on E: exception do
    begin
      logFile('Error SQL di F4.77 saat isi debet dan kredit ke saldo pada bulan itu  ' + E.Message);
      Exit;
    end;
  end;

  while not qQry.eof do
  begin
         // if  q1.fieldbyname('COA_ID').asstring = '251'   then
    if qQry.fieldbyname('COA_ID').asstring = '' then
    begin
      logFile('Ada yg Salah dengan SQL.77 COA_ID nya BLANK, Transaksi :' + qQry.fieldbyname('DOCUMENT_CLASS_CODE').asstring + ' No : ' + qQry.fieldbyname('DOCUMENT_NO').asstring);
      Exit;
    end
    else
    begin
      if not Update_coa(VPeriod, qQry.fieldbyname('COA_ID').asstring, 0, qQry.fieldbyname('debet').asfloat, qQry.fieldbyname('kredit').asfloat) then
        Exit;
    end;
    qQry.Next;
  end;
  Result := True;

 //  Progressbar1.StepBy(VProgress - Progressbar1.Position);
end;

function TsvReport.Reproses_accounting(vperiod: string): Boolean;
begin
  if not run_sql('SET @REPROSES_HPP = 1') then
    Exit;
  if not run_sql('SET @SINKRON = 1') then
    Exit;
  if not run_sql('SET @CLOSING = 1') then
    Exit;
  if not run_sql('SET @TRIGGER_DISABLED = 1') then
    Exit;
  if not run_sql('SET @TEMP_SINKRON = 1') then
    Exit;
  Proses_Coa_Balance(vperiod, 77);
  if not run_sql(' update PERIOD set last_update_date = sysdate(), LAST_UPDATE_BY = 777 where period_name = ' + vperiod) then
    Exit;
  if not run_sql('SET @REPROSES_HPP = 0') then
    Exit;
  if not run_sql('SET @SINKRON = 0') then
    Exit;
  if not run_sql('SET @CLOSING = 0') then
    Exit;
  if not run_sql('SET @TEMP_SINKRON = 0') then
    Exit;
  if not run_sql('SET @TRIGGER_DISABLED = null ') then
    Exit;

end;

function TsvReport.gantisql(nourut: integer; sumber: string): string;
var
  akhir, tempparam: string;
  i: integer;
begin
  akhir := sumber;
  i := 0;
  while i < JUM_PARAM_YG_DIAMBIL[nourut] do
  begin
    tempparam := midstr(PARAM_YG_DIAMBIL[nourut][i], 2, length(PARAM_YG_DIAMBIL[nourut][i]) - 1);

    try
      if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
      begin
        if RIGHTSTR(tempparam, 1) = '1' then
        begin
          akhir := AnsiReplaceStr(akhir, ':YEARMONTH1', '''' + yearmonth1 + '''');
        end
        else
        begin
          akhir := AnsiReplaceStr(akhir, ':YEARMONTH2', '''' + yearmonth2 + '''');
        end;
      end
      else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
      begin
        akhir := AnsiReplaceStr(akhir, ':YEAR_LALU', '''' + year_LALU + '''');
      end
      else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
      begin
        if RIGHTSTR(tempparam, 1) = '1' then
          akhir := AnsiReplaceStr(akhir, ':PERIOD_NAME1', '''' + period1 + '''')
        else
          akhir := AnsiReplaceStr(akhir, ':PERIOD_NAME2', '''' + period2 + '''');
      end
      else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
      begin
        akhir := AnsiReplaceStr(akhir, ':PERIOD_AWAL', '''' + PERIOD_AWAL + '''')
      end
      else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
      begin
        akhir := AnsiReplaceStr(akhir, ':PERIOD_LALU', '''' + PERIOD_LALU + '''')
      end
      else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
      begin
        if RIGHTSTR(tempparam, 1) = '1' then
        begin
          akhir := AnsiReplaceStr(akhir, ':YEAR1', '''' + year1 + '''');
        end
        else
        begin
          akhir := AnsiReplaceStr(akhir, ':YEAR2', '''' + year2 + '''');
        end;
      end
      else
      begin
        if PARAM_TYPE[strtoint(leftstr(tempparam, 1))] = 5 then
        begin
          anu := quotedstr(v_a_param_yg_dipasing[strtoint(leftstr(tempparam, 1)), strtoint(rightstr(tempparam, length(tempparam) - 1))]);
        end
        else
          anu := quotedstr(v_a_param_yg_dipasing[strtoint(leftstr(tempparam, 1)), strtoint(rightstr(tempparam, length(tempparam) - 1))]);

        anu2 := ':' + tempparam;
        akhir := AnsiReplaceStr(akhir, anu2, anu);
      end;
    except
    end;
    inc(i);
  end;

  gantisql := akhir;
end;

function TsvReport.hitungsql(NOURUT: integer): integer;
var
  nilai: integer;
begin
  case NOURUT of
    0:
      nilai := qHeader.RecordCount;
    1:
      nilai := q1.RecordCount;
    2:
      nilai := q2.RecordCount;
    3:
      nilai := qDetail.RecordCount;
    4:
      nilai := qNo_Form.RecordCount;
    5:
      nilai := q5.RecordCount;
    6:
      nilai := q6.RecordCount;
    7:
      nilai := q7.RecordCount;
    8:
      nilai := q8.RecordCount;
    9:
      nilai := qGroup.RecordCount;
    10:
      nilai := q3.RecordCount;
    11:
      nilai := q4.RecordCount;
    12:
      nilai := q9.RecordCount;
    13:
      nilai := qDetail1.RecordCount;
    14:
      nilai := qDetail2.RecordCount;
    15:
      nilai := qDetail3.RecordCount;
    16:
      nilai := qDetail4.RecordCount;
    17:
      nilai := q10.RecordCount;
    18:
      nilai := q11.RecordCount;
    19:
      nilai := q12.RecordCount;
    20:
      nilai := q14.RecordCount;
    21:
      nilai := q15.RecordCount;
    22:
      nilai := q16.RecordCount;
    23:
      nilai := q17.RecordCount;
    24:
      nilai := q18.RecordCount;
    25:
      nilai := q19.RecordCount;
    26:
      nilai := q20.RecordCount;

  end;

  result := nilai;
end;

procedure TsvReport.PARENT_SQL_HEADER(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qHeader.MasterSource := dsHeader;
    1:
      qHeader.MasterSource := ds1;
    2:
      qHeader.MasterSource := ds2;
    3:
      qHeader.MasterSource := dsDetail;
    4:
      qHeader.MasterSource := dsNo_Form;
    5:
      qHeader.MasterSource := ds5;
    6:
      qHeader.MasterSource := ds6;
    7:
      qHeader.MasterSource := ds7;
    8:
      qHeader.MasterSource := ds8;
    9:
      qHeader.MasterSource := dsGroup;
    10:
      qHeader.MasterSource := ds3;
    11:
      qHeader.MasterSource := ds4;
    12:
      qHeader.MasterSource := ds9;
    13:
      qHeader.MasterSource := dsDetail1;
    14:
      qHeader.MasterSource := dsDetail2;
    15:
      qHeader.MasterSource := dsDetail3;
    16:
      qHeader.MasterSource := dsDetail4;
    17:
      qHeader.MasterSource := ds10;
    18:
      qHeader.MasterSource := ds11;
    19:
      qHeader.MasterSource := ds12;
    20:
      qHeader.MasterSource := ds14;
    21:
      qHeader.MasterSource := ds15;
    22:
      qHeader.MasterSource := ds16;
    23:
      qHeader.MasterSource := ds17;
    24:
      qHeader.MasterSource := ds18;
    25:
      qHeader.MasterSource := ds19;
    26:
      qHeader.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_DETAIL(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qDetail.MasterSource := dsHeader;
    1:
      qDetail.MasterSource := ds1;
    2:
      qDetail.MasterSource := ds2;
    3:
      qDetail.MasterSource := dsDetail;
    4:
      qDetail.MasterSource := dsNo_Form;
    5:
      qDetail.MasterSource := ds5;
    6:
      qDetail.MasterSource := ds6;
    7:
      qDetail.MasterSource := ds7;
    8:
      qDetail.MasterSource := ds8;
    9:
      qDetail.MasterSource := dsGroup;
    10:
      qDetail.MasterSource := ds3;
    11:
      qDetail.MasterSource := ds4;
    12:
      qDetail.MasterSource := ds9;
    13:
      qDetail.MasterSource := dsDetail1;
    14:
      qDetail.MasterSource := dsDetail2;
    15:
      qDetail.MasterSource := dsDetail3;
    16:
      qDetail.MasterSource := dsDetail4;
    17:
      qDetail.MasterSource := ds10;
    18:
      qDetail.MasterSource := ds11;
    19:
      qDetail.MasterSource := ds12;
    20:
      qDetail.MasterSource := ds14;
    21:
      qDetail.MasterSource := ds15;
    22:
      qDetail.MasterSource := ds16;
    23:
      qDetail.MasterSource := ds17;
    24:
      qDetail.MasterSource := ds18;
    25:
      qDetail.MasterSource := ds19;
    26:
      qDetail.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_DETAIL1(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qDetail1.MasterSource := dsHeader;
    1:
      qDetail1.MasterSource := ds1;
    2:
      qDetail1.MasterSource := ds2;
    3:
      qDetail1.MasterSource := dsDetail;
    4:
      qDetail1.MasterSource := dsNo_Form;
    5:
      qDetail1.MasterSource := ds5;
    6:
      qDetail1.MasterSource := ds6;
    7:
      qDetail1.MasterSource := ds7;
    8:
      qDetail1.MasterSource := ds8;
    9:
      qDetail1.MasterSource := dsGroup;
    10:
      qDetail1.MasterSource := ds3;
    11:
      qDetail1.MasterSource := ds4;
    12:
      qDetail1.MasterSource := ds9;
    13:
      qDetail1.MasterSource := dsDetail1;
    14:
      qDetail1.MasterSource := dsDetail2;
    15:
      qDetail1.MasterSource := dsDetail3;
    16:
      qDetail1.MasterSource := dsDetail4;
    17:
      qDetail1.MasterSource := ds10;
    18:
      qDetail1.MasterSource := ds11;
    19:
      qDetail1.MasterSource := ds12;
    20:
      qDetail1.MasterSource := ds14;
    21:
      qDetail1.MasterSource := ds15;
    22:
      qDetail1.MasterSource := ds16;
    23:
      qDetail1.MasterSource := ds17;
    24:
      qDetail1.MasterSource := ds18;
    25:
      qDetail1.MasterSource := ds19;
    26:
      qDetail1.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_DETAIL2(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qDetail2.MasterSource := dsHeader;
    1:
      qDetail2.MasterSource := ds1;
    2:
      qDetail2.MasterSource := ds2;
    3:
      qDetail2.MasterSource := dsDetail;
    4:
      qDetail2.MasterSource := dsNo_Form;
    5:
      qDetail2.MasterSource := ds5;
    6:
      qDetail2.MasterSource := ds6;
    7:
      qDetail2.MasterSource := ds7;
    8:
      qDetail2.MasterSource := ds8;
    9:
      qDetail2.MasterSource := dsGroup;
    10:
      qDetail2.MasterSource := ds3;
    11:
      qDetail2.MasterSource := ds4;
    12:
      qDetail2.MasterSource := ds9;
    13:
      qDetail2.MasterSource := dsDetail1;
    14:
      qDetail2.MasterSource := dsDetail2;
    15:
      qDetail2.MasterSource := dsDetail3;
    16:
      qDetail2.MasterSource := dsDetail4;
    17:
      qDetail2.MasterSource := ds10;
    18:
      qDetail2.MasterSource := ds11;
    19:
      qDetail2.MasterSource := ds12;
    20:
      qDetail2.MasterSource := ds14;
    21:
      qDetail2.MasterSource := ds15;
    22:
      qDetail2.MasterSource := ds16;
    23:
      qDetail2.MasterSource := ds17;
    24:
      qDetail2.MasterSource := ds18;
    25:
      qDetail2.MasterSource := ds19;
    26:
      qDetail2.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_DETAIL3(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qDetail3.MasterSource := dsHeader;
    1:
      qDetail3.MasterSource := ds1;
    2:
      qDetail3.MasterSource := ds2;
    3:
      qDetail3.MasterSource := dsDetail;
    4:
      qDetail3.MasterSource := dsNo_Form;
    5:
      qDetail3.MasterSource := ds5;
    6:
      qDetail3.MasterSource := ds6;
    7:
      qDetail3.MasterSource := ds7;
    8:
      qDetail3.MasterSource := ds8;
    9:
      qDetail3.MasterSource := dsGroup;
    10:
      qDetail3.MasterSource := ds3;
    11:
      qDetail3.MasterSource := ds4;
    12:
      qDetail3.MasterSource := ds9;
    13:
      qDetail3.MasterSource := dsDetail1;
    14:
      qDetail3.MasterSource := dsDetail2;
    15:
      qDetail3.MasterSource := dsDetail3;
    16:
      qDetail3.MasterSource := dsDetail4;
    17:
      qDetail3.MasterSource := ds10;
    18:
      qDetail3.MasterSource := ds11;
    19:
      qDetail3.MasterSource := ds12;
    20:
      qDetail3.MasterSource := ds14;
    21:
      qDetail3.MasterSource := ds15;
    22:
      qDetail3.MasterSource := ds16;
    23:
      qDetail3.MasterSource := ds17;
    24:
      qDetail3.MasterSource := ds18;
    25:
      qDetail3.MasterSource := ds19;
    26:
      qDetail3.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_DETAIL4(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qDetail4.MasterSource := dsHeader;
    1:
      qDetail4.MasterSource := ds1;
    2:
      qDetail4.MasterSource := ds2;
    3:
      qDetail4.MasterSource := dsDetail;
    4:
      qDetail4.MasterSource := dsNo_Form;
    5:
      qDetail4.MasterSource := ds5;
    6:
      qDetail4.MasterSource := ds6;
    7:
      qDetail4.MasterSource := ds7;
    8:
      qDetail4.MasterSource := ds8;
    9:
      qDetail4.MasterSource := dsGroup;
    10:
      qDetail4.MasterSource := ds3;
    11:
      qDetail4.MasterSource := ds4;
    12:
      qDetail4.MasterSource := ds9;
    13:
      qDetail4.MasterSource := dsDetail1;
    14:
      qDetail4.MasterSource := dsDetail2;
    15:
      qDetail4.MasterSource := dsDetail3;
    16:
      qDetail4.MasterSource := dsDetail4;
    17:
      qDetail4.MasterSource := ds10;
    18:
      qDetail4.MasterSource := ds11;
    19:
      qDetail4.MasterSource := ds12;
    20:
      qDetail4.MasterSource := ds14;
    21:
      qDetail4.MasterSource := ds15;
    22:
      qDetail4.MasterSource := ds16;
    23:
      qDetail4.MasterSource := ds17;
    24:
      qDetail4.MasterSource := ds18;
    25:
      qDetail4.MasterSource := ds19;
    26:
      qDetail4.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_No_Form(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      qNo_Form.MasterSource := dsHeader;
    1:
      qNo_Form.MasterSource := ds1;
    2:
      qNo_Form.MasterSource := ds2;
    3:
      qNo_Form.MasterSource := dsDetail;
    4:
      qNo_Form.MasterSource := dsNo_Form;
    5:
      qNo_Form.MasterSource := ds5;
    6:
      qNo_Form.MasterSource := ds6;
    7:
      qNo_Form.MasterSource := ds7;
    8:
      qNo_Form.MasterSource := ds8;
    9:
      qNo_Form.MasterSource := dsGroup;
    10:
      qNo_Form.MasterSource := ds3;
    11:
      qNo_Form.MasterSource := ds4;
    12:
      qNo_Form.MasterSource := ds9;
    13:
      qNo_Form.MasterSource := dsDetail1;
    14:
      qNo_Form.MasterSource := dsDetail2;
    15:
      qNo_Form.MasterSource := dsDetail3;
    16:
      qNo_Form.MasterSource := dsDetail4;
    17:
      qNo_Form.MasterSource := ds10;
    18:
      qNo_Form.MasterSource := ds11;
    19:
      qNo_Form.MasterSource := ds12;
    20:
      qNo_Form.MasterSource := ds14;
    21:
      qNo_Form.MasterSource := ds15;
    22:
      qNo_Form.MasterSource := ds16;
    23:
      qNo_Form.MasterSource := ds17;
    24:
      qNo_Form.MasterSource := ds18;
    25:
      qNo_Form.MasterSource := ds19;
    26:
      qNo_Form.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_1(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      q1.MasterSource := dsHeader;
    1:
      q1.MasterSource := ds1;
    2:
      q1.MasterSource := ds2;
    3:
      q1.MasterSource := dsDetail;
    4:
      q1.MasterSource := dsNo_Form;
    5:
      q1.MasterSource := ds5;
    6:
      q1.MasterSource := ds6;
    7:
      q1.MasterSource := ds7;
    8:
      q1.MasterSource := ds8;
    9:
      q1.MasterSource := dsGroup;
    10:
      q1.MasterSource := ds3;
    11:
      q1.MasterSource := ds4;
    12:
      q1.MasterSource := ds9;
    13:
      q1.MasterSource := dsDetail1;
    14:
      q1.MasterSource := dsDetail2;
    15:
      q1.MasterSource := dsDetail3;
    16:
      q1.MasterSource := dsDetail4;
    17:
      q1.MasterSource := ds10;
    18:
      q1.MasterSource := ds11;
    19:
      q1.MasterSource := ds12;
    20:
      q1.MasterSource := ds14;
    21:
      q1.MasterSource := ds15;
    22:
      q1.MasterSource := ds16;
    23:
      q1.MasterSource := ds17;
    24:
      q1.MasterSource := ds18;
    25:
      q1.MasterSource := ds19;
    26:
      q1.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_2(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      q2.MasterSource := dsHeader;
    1:
      q2.MasterSource := ds1;
    2:
      q2.MasterSource := ds2;
    3:
      q2.MasterSource := dsDetail;
    4:
      q2.MasterSource := dsNo_Form;
    5:
      q2.MasterSource := ds5;
    6:
      q2.MasterSource := ds6;
    7:
      q2.MasterSource := ds7;
    8:
      q2.MasterSource := ds8;
    9:
      q2.MasterSource := dsGroup;
    10:
      q2.MasterSource := ds3;
    11:
      q2.MasterSource := ds4;
    12:
      q2.MasterSource := ds9;
    13:
      q2.MasterSource := dsDetail1;
    14:
      q2.MasterSource := dsDetail2;
    15:
      q2.MasterSource := dsDetail3;
    16:
      q2.MasterSource := dsDetail4;
    17:
      q2.MasterSource := ds10;
    18:
      q2.MasterSource := ds11;
    19:
      q2.MasterSource := ds12;
    20:
      q2.MasterSource := ds14;
    21:
      q2.MasterSource := ds15;
    22:
      q2.MasterSource := ds16;
    23:
      q2.MasterSource := ds17;
    24:
      q2.MasterSource := ds18;
    25:
      q2.MasterSource := ds19;
    26:
      q2.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_3(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q3.MasterSource := dsHeader;
    1:
      Q3.MasterSource := ds1;
    2:
      Q3.MasterSource := ds2;
    3:
      Q3.MasterSource := dsDetail;
    4:
      Q3.MasterSource := dsNo_Form;
    5:
      Q3.MasterSource := ds5;
    6:
      Q3.MasterSource := ds6;
    7:
      Q3.MasterSource := ds7;
    8:
      Q3.MasterSource := ds8;
    9:
      Q3.MasterSource := dsGroup;
    10:
      Q3.MasterSource := ds3;
    11:
      Q3.MasterSource := ds4;
    12:
      Q3.MasterSource := ds9;
    13:
      Q3.MasterSource := dsDetail1;
    14:
      Q3.MasterSource := dsDetail2;
    15:
      Q3.MasterSource := dsDetail3;
    16:
      Q3.MasterSource := dsDetail4;
    17:
      Q3.MasterSource := ds10;
    18:
      Q3.MasterSource := ds11;
    19:
      Q3.MasterSource := ds12;
    20:
      Q3.MasterSource := ds14;
    21:
      Q3.MasterSource := ds15;
    22:
      Q3.MasterSource := ds16;
    23:
      Q3.MasterSource := ds17;
    24:
      Q3.MasterSource := ds18;
    25:
      Q3.MasterSource := ds19;
    26:
      Q3.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_4(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q4.MasterSource := dsHeader;
    1:
      Q4.MasterSource := ds1;
    2:
      Q4.MasterSource := ds2;
    3:
      Q4.MasterSource := dsDetail;
    4:
      Q4.MasterSource := dsNo_Form;
    5:
      Q4.MasterSource := ds5;
    6:
      Q4.MasterSource := ds6;
    7:
      Q4.MasterSource := ds7;
    8:
      Q4.MasterSource := ds8;
    9:
      Q4.MasterSource := dsGroup;
    10:
      Q4.MasterSource := ds3;
    11:
      Q4.MasterSource := ds4;
    12:
      Q4.MasterSource := ds9;
    13:
      Q4.MasterSource := dsDetail1;
    14:
      Q4.MasterSource := dsDetail2;
    15:
      Q4.MasterSource := dsDetail3;
    16:
      Q4.MasterSource := dsDetail4;
    17:
      Q4.MasterSource := ds10;
    18:
      Q4.MasterSource := ds11;
    19:
      Q4.MasterSource := ds12;
    20:
      Q4.MasterSource := ds14;
    21:
      Q4.MasterSource := ds15;
    22:
      Q4.MasterSource := ds16;
    23:
      Q4.MasterSource := ds17;
    24:
      Q4.MasterSource := ds18;
    25:
      Q4.MasterSource := ds19;
    26:
      Q4.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_5(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q5.MasterSource := dsHeader;
    1:
      Q5.MasterSource := ds1;
    2:
      Q5.MasterSource := ds2;
    3:
      Q5.MasterSource := dsDetail;
    4:
      Q5.MasterSource := dsNo_Form;
    5:
      Q5.MasterSource := ds5;
    6:
      Q5.MasterSource := ds6;
    7:
      Q5.MasterSource := ds7;
    8:
      Q5.MasterSource := ds8;
    9:
      Q5.MasterSource := dsGroup;
    10:
      Q5.MasterSource := ds3;
    11:
      Q5.MasterSource := ds4;
    12:
      Q5.MasterSource := ds9;
    13:
      Q5.MasterSource := dsDetail1;
    14:
      Q5.MasterSource := dsDetail2;
    15:
      Q5.MasterSource := dsDetail3;
    16:
      Q5.MasterSource := dsDetail4;
    17:
      Q5.MasterSource := ds10;
    18:
      Q5.MasterSource := ds11;
    19:
      Q5.MasterSource := ds12;
    20:
      Q5.MasterSource := ds14;
    21:
      Q5.MasterSource := ds15;
    22:
      Q5.MasterSource := ds16;
    23:
      Q5.MasterSource := ds17;
    24:
      Q5.MasterSource := ds18;
    25:
      Q5.MasterSource := ds19;
    26:
      Q5.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_6(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q6.MasterSource := dsHeader;
    1:
      Q6.MasterSource := ds1;
    2:
      Q6.MasterSource := ds2;
    3:
      Q6.MasterSource := dsDetail;
    4:
      Q6.MasterSource := dsNo_Form;
    5:
      Q6.MasterSource := ds5;
    6:
      Q6.MasterSource := ds6;
    7:
      Q6.MasterSource := ds7;
    8:
      Q6.MasterSource := ds8;
    9:
      Q6.MasterSource := dsGroup;
    10:
      Q6.MasterSource := ds3;
    11:
      Q6.MasterSource := ds4;
    12:
      Q6.MasterSource := ds9;
    13:
      Q6.MasterSource := dsDetail1;
    14:
      Q6.MasterSource := dsDetail2;
    15:
      Q6.MasterSource := dsDetail3;
    16:
      Q6.MasterSource := dsDetail4;
    17:
      Q6.MasterSource := ds10;
    18:
      Q6.MasterSource := ds11;
    19:
      Q6.MasterSource := ds12;
    20:
      Q6.MasterSource := ds14;
    21:
      Q6.MasterSource := ds15;
    22:
      Q6.MasterSource := ds16;
    23:
      Q6.MasterSource := ds17;
    24:
      Q6.MasterSource := ds18;
    25:
      Q6.MasterSource := ds19;
    26:
      Q6.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_7(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q7.MasterSource := dsHeader;
    1:
      Q7.MasterSource := ds1;
    2:
      Q7.MasterSource := ds2;
    3:
      Q7.MasterSource := dsDetail;
    4:
      Q7.MasterSource := dsNo_Form;
    5:
      Q7.MasterSource := ds5;
    6:
      Q7.MasterSource := ds6;
    7:
      Q7.MasterSource := ds7;
    8:
      Q7.MasterSource := ds8;
    9:
      Q7.MasterSource := dsGroup;
    10:
      Q7.MasterSource := ds3;
    11:
      Q7.MasterSource := ds4;
    12:
      Q7.MasterSource := ds9;
    13:
      Q7.MasterSource := dsDetail1;
    14:
      Q7.MasterSource := dsDetail2;
    15:
      Q7.MasterSource := dsDetail3;
    16:
      Q7.MasterSource := dsDetail4;
    17:
      Q7.MasterSource := ds10;
    18:
      Q7.MasterSource := ds11;
    19:
      Q7.MasterSource := ds12;
    20:
      Q7.MasterSource := ds14;
    21:
      Q7.MasterSource := ds15;
    22:
      Q7.MasterSource := ds16;
    23:
      Q7.MasterSource := ds17;
    24:
      Q7.MasterSource := ds18;
    25:
      Q7.MasterSource := ds19;
    26:
      Q7.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_8(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q8.MasterSource := dsHeader;
    1:
      Q8.MasterSource := ds1;
    2:
      Q8.MasterSource := ds2;
    3:
      Q8.MasterSource := dsDetail;
    4:
      Q8.MasterSource := dsNo_Form;
    5:
      Q8.MasterSource := ds5;
    6:
      Q8.MasterSource := ds6;
    7:
      Q8.MasterSource := ds7;
    8:
      Q8.MasterSource := ds8;
    9:
      Q8.MasterSource := dsGroup;
    10:
      Q8.MasterSource := ds3;
    11:
      Q8.MasterSource := ds4;
    12:
      Q8.MasterSource := ds9;
    13:
      Q8.MasterSource := dsDetail1;
    14:
      Q8.MasterSource := dsDetail2;
    15:
      Q8.MasterSource := dsDetail3;
    16:
      Q8.MasterSource := dsDetail4;
    17:
      Q8.MasterSource := ds10;
    18:
      Q8.MasterSource := ds11;
    19:
      Q8.MasterSource := ds12;
    20:
      Q8.MasterSource := ds14;
    21:
      Q8.MasterSource := ds15;
    22:
      Q8.MasterSource := ds16;
    23:
      Q8.MasterSource := ds17;
    24:
      Q8.MasterSource := ds18;
    25:
      Q8.MasterSource := ds19;
    26:
      Q8.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_9(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q9.MasterSource := dsHeader;
    1:
      Q9.MasterSource := ds1;
    2:
      Q9.MasterSource := ds2;
    3:
      Q9.MasterSource := dsDetail;
    4:
      Q9.MasterSource := dsNo_Form;
    5:
      Q9.MasterSource := ds5;
    6:
      Q9.MasterSource := ds6;
    7:
      Q9.MasterSource := ds7;
    8:
      Q9.MasterSource := ds8;
    9:
      Q9.MasterSource := dsGroup;
    10:
      Q9.MasterSource := ds3;
    11:
      Q9.MasterSource := ds4;
    12:
      Q9.MasterSource := ds9;
    13:
      Q9.MasterSource := dsDetail1;
    14:
      Q9.MasterSource := dsDetail2;
    15:
      Q9.MasterSource := dsDetail3;
    16:
      Q9.MasterSource := dsDetail4;
    17:
      Q9.MasterSource := ds10;
    18:
      Q9.MasterSource := ds11;
    19:
      Q9.MasterSource := ds12;
    20:
      Q9.MasterSource := ds14;
    21:
      Q9.MasterSource := ds15;
    22:
      Q9.MasterSource := ds16;
    23:
      Q9.MasterSource := ds17;
    24:
      Q9.MasterSource := ds18;
    25:
      Q9.MasterSource := ds19;
    26:
      Q9.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_10(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q10.MasterSource := dsHeader;
    1:
      Q10.MasterSource := ds1;
    2:
      Q10.MasterSource := ds2;
    3:
      Q10.MasterSource := dsDetail;
    4:
      Q10.MasterSource := dsNo_Form;
    5:
      Q10.MasterSource := ds5;
    6:
      Q10.MasterSource := ds6;
    7:
      Q10.MasterSource := ds7;
    8:
      Q10.MasterSource := ds8;
    9:
      Q10.MasterSource := dsGroup;
    10:
      Q10.MasterSource := ds3;
    11:
      Q10.MasterSource := ds4;
    12:
      Q10.MasterSource := ds9;
    13:
      Q10.MasterSource := dsDetail1;
    14:
      Q10.MasterSource := dsDetail2;
    15:
      Q10.MasterSource := dsDetail3;
    16:
      Q10.MasterSource := dsDetail4;
    17:
      Q10.MasterSource := ds10;
    18:
      Q10.MasterSource := ds11;
    19:
      Q10.MasterSource := ds12;
    20:
      Q10.MasterSource := ds14;
    21:
      Q10.MasterSource := ds15;
    22:
      Q10.MasterSource := ds16;
    23:
      Q10.MasterSource := ds17;
    24:
      Q10.MasterSource := ds18;
    25:
      Q10.MasterSource := ds19;
    26:
      Q10.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_11(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q11.MasterSource := dsHeader;
    1:
      Q11.MasterSource := ds1;
    2:
      Q11.MasterSource := ds2;
    3:
      Q11.MasterSource := dsDetail;
    4:
      Q11.MasterSource := dsNo_Form;
    5:
      Q11.MasterSource := ds5;
    6:
      Q11.MasterSource := ds6;
    7:
      Q11.MasterSource := ds7;
    8:
      Q11.MasterSource := ds8;
    9:
      Q11.MasterSource := dsGroup;
    10:
      Q11.MasterSource := ds3;
    11:
      Q11.MasterSource := ds4;
    12:
      Q11.MasterSource := ds9;
    13:
      Q11.MasterSource := dsDetail1;
    14:
      Q11.MasterSource := dsDetail2;
    15:
      Q11.MasterSource := dsDetail3;
    16:
      Q11.MasterSource := dsDetail4;
    17:
      Q11.MasterSource := ds10;
    18:
      Q11.MasterSource := ds11;
    19:
      Q11.MasterSource := ds12;
    20:
      Q11.MasterSource := ds14;
    21:
      Q11.MasterSource := ds15;
    22:
      Q11.MasterSource := ds16;
    23:
      Q11.MasterSource := ds17;
    24:
      Q11.MasterSource := ds18;
    25:
      Q11.MasterSource := ds19;
    26:
      Q11.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_12(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q12.MasterSource := dsHeader;
    1:
      Q12.MasterSource := ds1;
    2:
      Q12.MasterSource := ds2;
    3:
      Q12.MasterSource := dsDetail;
    4:
      Q12.MasterSource := dsNo_Form;
    5:
      Q12.MasterSource := ds5;
    6:
      Q12.MasterSource := ds6;
    7:
      Q12.MasterSource := ds7;
    8:
      Q12.MasterSource := ds8;
    9:
      Q12.MasterSource := dsGroup;
    10:
      Q12.MasterSource := ds3;
    11:
      Q12.MasterSource := ds4;
    12:
      Q12.MasterSource := ds9;
    13:
      Q12.MasterSource := dsDetail1;
    14:
      Q12.MasterSource := dsDetail2;
    15:
      Q12.MasterSource := dsDetail3;
    16:
      Q12.MasterSource := dsDetail4;
    17:
      Q12.MasterSource := ds10;
    18:
      Q12.MasterSource := ds11;
    19:
      Q12.MasterSource := ds12;
    20:
      Q12.MasterSource := ds14;
    21:
      Q12.MasterSource := ds15;
    22:
      Q12.MasterSource := ds16;
    23:
      Q12.MasterSource := ds17;
    24:
      Q12.MasterSource := ds18;
    25:
      Q12.MasterSource := ds19;
    26:
      Q12.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_14(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q14.MasterSource := dsHeader;
    1:
      Q14.MasterSource := ds1;
    2:
      Q14.MasterSource := ds2;
    3:
      Q14.MasterSource := dsDetail;
    4:
      Q14.MasterSource := dsNo_Form;
    5:
      Q14.MasterSource := ds5;
    6:
      Q14.MasterSource := ds6;
    7:
      Q14.MasterSource := ds7;
    8:
      Q14.MasterSource := ds8;
    9:
      Q14.MasterSource := dsGroup;
    10:
      Q14.MasterSource := ds3;
    11:
      Q14.MasterSource := ds4;
    12:
      Q14.MasterSource := ds9;
    13:
      Q14.MasterSource := dsDetail1;
    14:
      Q14.MasterSource := dsDetail2;
    15:
      Q14.MasterSource := dsDetail3;
    16:
      Q14.MasterSource := dsDetail4;
    17:
      Q14.MasterSource := ds10;
    18:
      Q14.MasterSource := ds11;
    19:
      Q14.MasterSource := ds12;
    20:
      Q14.MasterSource := ds14;
    21:
      Q14.MasterSource := ds15;
    22:
      Q14.MasterSource := ds16;
    23:
      Q14.MasterSource := ds17;
    24:
      Q14.MasterSource := ds18;
    25:
      Q14.MasterSource := ds19;
    26:
      Q14.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_15(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q15.MasterSource := dsHeader;
    1:
      Q15.MasterSource := ds1;
    2:
      Q15.MasterSource := ds2;
    3:
      Q15.MasterSource := dsDetail;
    4:
      Q15.MasterSource := dsNo_Form;
    5:
      Q15.MasterSource := ds5;
    6:
      Q15.MasterSource := ds6;
    7:
      Q15.MasterSource := ds7;
    8:
      Q15.MasterSource := ds8;
    9:
      Q15.MasterSource := dsGroup;
    10:
      Q15.MasterSource := ds3;
    11:
      Q15.MasterSource := ds4;
    12:
      Q15.MasterSource := ds9;
    13:
      Q15.MasterSource := dsDetail1;
    14:
      Q15.MasterSource := dsDetail2;
    15:
      Q15.MasterSource := dsDetail3;
    16:
      Q15.MasterSource := dsDetail4;
    17:
      Q15.MasterSource := ds10;
    18:
      Q15.MasterSource := ds11;
    19:
      Q15.MasterSource := ds12;
    20:
      Q15.MasterSource := ds14;
    21:
      Q15.MasterSource := ds15;
    22:
      Q15.MasterSource := ds16;
    23:
      Q15.MasterSource := ds17;
    24:
      Q15.MasterSource := ds18;
    25:
      Q15.MasterSource := ds19;
    26:
      Q15.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_16(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q16.MasterSource := dsHeader;
    1:
      Q16.MasterSource := ds1;
    2:
      Q16.MasterSource := ds2;
    3:
      Q16.MasterSource := dsDetail;
    4:
      Q16.MasterSource := dsNo_Form;
    5:
      Q16.MasterSource := ds5;
    6:
      Q16.MasterSource := ds6;
    7:
      Q16.MasterSource := ds7;
    8:
      Q16.MasterSource := ds8;
    9:
      Q16.MasterSource := dsGroup;
    10:
      Q16.MasterSource := ds3;
    11:
      Q16.MasterSource := ds4;
    12:
      Q16.MasterSource := ds9;
    13:
      Q16.MasterSource := dsDetail1;
    14:
      Q16.MasterSource := dsDetail2;
    15:
      Q16.MasterSource := dsDetail3;
    16:
      Q16.MasterSource := dsDetail4;
    17:
      Q16.MasterSource := ds10;
    18:
      Q16.MasterSource := ds11;
    19:
      Q16.MasterSource := ds12;
    20:
      Q16.MasterSource := ds14;
    21:
      Q16.MasterSource := ds15;
    22:
      Q16.MasterSource := ds16;
    23:
      Q16.MasterSource := ds17;
    24:
      Q16.MasterSource := ds18;
    25:
      Q16.MasterSource := ds19;
    26:
      Q16.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_17(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q17.MasterSource := dsHeader;
    1:
      Q17.MasterSource := ds1;
    2:
      Q17.MasterSource := ds2;
    3:
      Q17.MasterSource := dsDetail;
    4:
      Q17.MasterSource := dsNo_Form;
    5:
      Q17.MasterSource := ds5;
    6:
      Q17.MasterSource := ds6;
    7:
      Q17.MasterSource := ds7;
    8:
      Q17.MasterSource := ds8;
    9:
      Q17.MasterSource := dsGroup;
    10:
      Q17.MasterSource := ds3;
    11:
      Q17.MasterSource := ds4;
    12:
      Q17.MasterSource := ds9;
    13:
      Q17.MasterSource := dsDetail1;
    14:
      Q17.MasterSource := dsDetail2;
    15:
      Q17.MasterSource := dsDetail3;
    16:
      Q17.MasterSource := dsDetail4;
    17:
      Q17.MasterSource := ds10;
    18:
      Q17.MasterSource := ds11;
    19:
      Q17.MasterSource := ds12;
    20:
      Q17.MasterSource := ds14;
    21:
      Q17.MasterSource := ds15;
    22:
      Q17.MasterSource := ds16;
    23:
      Q17.MasterSource := ds17;
    24:
      Q17.MasterSource := ds18;
    25:
      Q17.MasterSource := ds19;
    26:
      Q17.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_18(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q18.MasterSource := dsHeader;
    1:
      Q18.MasterSource := ds1;
    2:
      Q18.MasterSource := ds2;
    3:
      Q18.MasterSource := dsDetail;
    4:
      Q18.MasterSource := dsNo_Form;
    5:
      Q18.MasterSource := ds5;
    6:
      Q18.MasterSource := ds6;
    7:
      Q18.MasterSource := ds7;
    8:
      Q18.MasterSource := ds8;
    9:
      Q18.MasterSource := dsGroup;
    10:
      Q18.MasterSource := ds3;
    11:
      Q18.MasterSource := ds4;
    12:
      Q18.MasterSource := ds9;
    13:
      Q18.MasterSource := dsDetail1;
    14:
      Q18.MasterSource := dsDetail2;
    15:
      Q18.MasterSource := dsDetail3;
    16:
      Q18.MasterSource := dsDetail4;
    17:
      Q18.MasterSource := ds10;
    18:
      Q18.MasterSource := ds11;
    19:
      Q18.MasterSource := ds12;
    20:
      Q18.MasterSource := ds14;
    21:
      Q18.MasterSource := ds15;
    22:
      Q18.MasterSource := ds16;
    23:
      Q18.MasterSource := ds17;
    24:
      Q18.MasterSource := ds18;
    25:
      Q18.MasterSource := ds19;
    26:
      Q18.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_19(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q19.MasterSource := dsHeader;
    1:
      Q19.MasterSource := ds1;
    2:
      Q19.MasterSource := ds2;
    3:
      Q19.MasterSource := dsDetail;
    4:
      Q19.MasterSource := dsNo_Form;
    5:
      Q19.MasterSource := ds5;
    6:
      Q19.MasterSource := ds6;
    7:
      Q19.MasterSource := ds7;
    8:
      Q19.MasterSource := ds8;
    9:
      Q19.MasterSource := dsGroup;
    10:
      Q19.MasterSource := ds3;
    11:
      Q19.MasterSource := ds4;
    12:
      Q19.MasterSource := ds9;
    13:
      Q19.MasterSource := dsDetail1;
    14:
      Q19.MasterSource := dsDetail2;
    15:
      Q19.MasterSource := dsDetail3;
    16:
      Q19.MasterSource := dsDetail4;
    17:
      Q19.MasterSource := ds10;
    18:
      Q19.MasterSource := ds11;
    19:
      Q19.MasterSource := ds12;
    20:
      Q19.MasterSource := ds14;
    21:
      Q19.MasterSource := ds15;
    22:
      Q19.MasterSource := ds16;
    23:
      Q19.MasterSource := ds17;
    24:
      Q19.MasterSource := ds18;
    25:
      Q19.MasterSource := ds19;
    26:
      Q19.MasterSource := ds20;
  end;
end;

procedure TsvReport.PARENT_SQL_20(NOURUT: INTEGER);
begin
  case PARENT_SQL[NOURUT] of
    0:
      Q20.MasterSource := dsHeader;
    1:
      Q20.MasterSource := ds1;
    2:
      Q20.MasterSource := ds2;
    3:
      Q20.MasterSource := dsDetail;
    4:
      Q20.MasterSource := dsNo_Form;
    5:
      Q20.MasterSource := ds5;
    6:
      Q20.MasterSource := ds6;
    7:
      Q20.MasterSource := ds7;
    8:
      Q20.MasterSource := ds8;
    9:
      Q20.MasterSource := dsGroup;
    10:
      Q20.MasterSource := ds3;
    11:
      Q20.MasterSource := ds4;
    12:
      Q20.MasterSource := ds9;
    13:
      Q20.MasterSource := dsDetail1;
    14:
      Q20.MasterSource := dsDetail2;
    15:
      Q20.MasterSource := dsDetail3;
    16:
      Q20.MasterSource := dsDetail4;
    17:
      Q20.MasterSource := ds10;
    18:
      Q20.MasterSource := ds11;
    19:
      Q20.MasterSource := ds12;
    20:
      Q20.MasterSource := ds14;
    21:
      Q20.MasterSource := ds15;
    22:
      Q20.MasterSource := ds16;
    23:
      Q20.MasterSource := ds17;
    24:
      Q20.MasterSource := ds18;
    25:
      Q20.MasterSource := ds19;
    26:
      Q20.MasterSource := ds20;
  end;
end;

function TsvReport.Proses_data1: boolean;
var
  SKondisi: string;
  tempparam, tempparam_baru: string;
  NOURUT: INTEGER;
  i: integer;
  jumlahrecordcount: integer;
  q: TmyQuery;
  DATAADA, variable_periode_flag: boolean;
  Repstart: TFlexCelReport;

  procedure BuatFolderUnik(const BaseDir: string);
  var
    NewDir: string;
    i: Integer;
  begin
    NewDir := BaseDir;
    i := 1;
    while DirectoryExists(NewDir) do
    begin
      NewDir := BaseDir + '_' + IntToStr(i);
      Inc(i);
    end;
    if ForceDirectories(NewDir) then
      logFile('Folder dibuat: ' + NewDir)
    else
      logFile('Gagal membuat folder');
  end;

begin

  //function str_filter untuk mengisi ke macro di tiap-tiap query, berdasar pada tabel erp_rpt_detail
  //yang mana penyimpanan semua macro pada variable skondisiall dan berdasar pada tabel erp_rpt_detail, sql_uses
  //nama_alias pada tabel erp_rpt_detail juga dipakai berdasarkan pada isi 'sql_uses'
  SKondisi := Str_filter(true);
  jumlahrecordcount := 0;
  vErr := '';
  Result := True;
  tempall := tempall + 'Group: ' + #10;
  DATAADA := false;

  Repstart := TFlexCelReport.Create(true);
  try
    try
      if (qsql.fieldbyname('r_group').asstring = 'MANAGEMENT') then
      begin
        //perlu_reproses_flag := true;
        run_query('select open_flag, LAST_UPDATE_BY  from period where period_name = ' + Period1);
        {if (query1.fieldbyname('beda_tgl').asinteger = 0 ) then
           if (query1.fieldbyname('beda_jam').asfloat < 0.01) then
               perlu_reproses_flag := false;

         }
        if (query1.fieldbyname('OPEN_FLAG').asstring = 'Y') then
        begin

          reproses_acc_flag := true;
          Reproses_accounting(PERIOD1);
        end;
      end;
      variable_periode_flag := false;
      if pos('PERIOD', uppercase(qsql.FIELDBYNAME('KONDISI').AsString)) > 0 then
      begin
        variable_periode_flag := true;
      end;

      if qGroup.SQL.text <> '' then
      begin
        NOURUT := 9;
        qGroup.Active := false;
        if qGroup.findparam('USER') <> nil then
          qGroup.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qGroup.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qGroup.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qGroup.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qGroup.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qGroup.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qGroup.MacroByName('kondisiSales').Value := ' ';
        end;
        if qGroup.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qGroup.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qGroup.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qGroup.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qGroup.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qGroup.MacroByName('kondisiMerek').Value := ' ';
        end;
        qGroup.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
        {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          try
            if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
            begin
              if RIGHTSTR(tempparam, 1) = '1' then
                qGroup.ParamByName(tempparam).asstring := YEARMONTH1
              else
                qGroup.ParamByName(tempparam).asstring := YEARMONTH2;
            end
            else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
            begin
              if RIGHTSTR(tempparam, 1) = '1' then
                qGroup.ParamByName(tempparam).asstring := PERIOD1
              else
                qGroup.ParamByName(tempparam).asstring := PERIOD2;
            end
            else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
            begin
              qGroup.ParamByName(tempparam).asstring := PERIOD_AWAL;
            end
            else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
            begin
              qGroup.ParamByName(tempparam).asstring := YEAR_LALU;
            end
            else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
            begin
              QGroup.ParamByName(tempparam).asstring := PERIOD_LALU;
            end
            else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
            begin
              if RIGHTSTR(tempparam, 1) = '1' then
                qGroup.ParamByName(tempparam).asstring := YEAR1
              else
                qGroup.ParamByName(tempparam).asstring := YEAR2;
            end
            else
            begin
              // passing parameter mis :11 ke query
              qGroup.ParamByName(tempparam).asstring := v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))];
               //dbl+dbl+dbl+'2006-04'+dbl+dbl+dbl;
              //quotedstr(v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru,1)),strtoint(rightstr(tempparam_baru,length(tempparam_baru)-1 ))]);
            end;
          except
            logfile('Isi qGroup tidak ada parameter :' + tempparam);
          end;
          inc(i);
        end;
     //try
        if qGroup.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            qGroup.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            qGroup.MacroByName('kondisiUser').Value := ' ';
        end;

        tempall := tempall + gantisql(NOURUT, qGroup.FinalSQL) + #10 + #13;
        try
          qGroup.Active := true;
          Repstart.AddTable('qGroup', qGroup);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qGroup failed open ' + E.Message + '!');
            logFile('qGroup query : ' + qGroup.FinalSQL + '!');
            Exit;
          end;
        end;

        logFile('qGroup : ' + qGroup.FinalSQL);
        if not qgroup.eof then
          DATAADA := true;
      //messagedlg('nilai dari field count '+inttostr(qGroup.FieldList.Count),mtinformation,[mbOk],0);;
    //except
      //messagedlg('maaf terdapat error di qGroup '+#10+#13+qGroup.sql.Text+#10+#13+'yang diisi pada qGroup '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
    //end;
     //jumlah_isi_query[no_group]
     //logfile('jumlah isi Group '+inttostr(qGroup.RecordCount));
        jumlahrecordcount := qGroup.RecordCount * banyakbaris[NOURUT];
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qGroup.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;

      end;
      tempall := tempall + 'Header: ' + #10;

      if qheader.SQL.text <> '' then
      begin
        NOURUT := 0;
        qHeader.Active := false;
        if qHeader.findparam('USER') <> nil then
          qHeader.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qHeader.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qHeader.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qHeader.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qHeader.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qHeader.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qHeader.MacroByName('kondisiSales').Value := ' ';
        end;
        if qHeader.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qHeader.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qHeader.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qHeader.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qHeader.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qHeader.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_HEADER(NOURUT);
        qHeader.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

      //try
          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qHeader.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qHeader.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qHeader.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qHeader.ParamByName(tempparam).asstring := PERIOD1
            else
              qHeader.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qHeader.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QHeader.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qHeader.ParamByName(tempparam).asstring := YEAR1
            else
              qHeader.ParamByName(tempparam).asstring := YEAR2;
          end
          else
          //v_a_param_yg_dipasing harus memakai id
          //tgl 11 belum diambil
        //  qHeader.ParamByName(tempparam).asstring := '333';//v_a_param_yg_dipasing[1,16];

            qHeader.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          // agar keluar ''''' qHeader.ParamByName(tempparam_baru).asstring := quotedstr(v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru,1)),strtoint(rightstr(tempparam_baru,1))]);
          //
        //except
        //       logfile('Isi qHeader tidak ada parameter :'+TEMPPARAM);
        //end;
          inc(i);
        end;
    //try

        if qHeader.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            qHeader.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            qHeader.MacroByName('kondisiUser').Value := ' ';
        end;

        tempall := tempall + gantisql(NOURUT, qHeader.FinalSQL) + #10 + #13;
//      qHeader.Active := true;

    {  q :=  tmyquery(qheader.MasterSource);
      IF (PARENT_SQL[NOURUT]=AngkaNonAktif) or (not q.Eof) then  begin
     }

        try
          qHeader.Active := true;
          Repstart.AddTable('qHeader', qHeader);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qHeader failed open ' + E.Message + '!');
            logFile('qHeader query : ' + qHeader.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qHeader : ' + qHeader.FinalSQL);
        if not qHeader.eof then
          DATAADA := true;
     // end;
      //messagedlg('nilai dari field count '+inttostr(qGroup.FieldList.Count),mtinformation,[mbOk],0);;
    //except
      //messagedlg('maaf terdapat error di Qheader '+#10+#13+qHeader.sql.Text+#10+#13+'yang diisi pada Qheader '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);

    //end;
     //logfile(' jumlah isi header '+inttostr(qHeader.RecordCount));
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
     //logfile('jumlah record sebanyak '+IntToStr(jumlahrecordcount)+'.');
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qHeader.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'Detail: ' + #10;
      if qdetail.SQL.text <> '' then
      begin
        NOURUT := 3;
        qDetail.Active := false;
        if qDetail.findparam('USER') <> nil then
          qDetail.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qDetail.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qDetail.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qDetail.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qDetail.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail.MacroByName('kondisiSales').Value := ' ';
        end;
        if qDetail.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qDetail.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qDetail.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qDetail.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_DETAIL(NOURUT);
        qdetail.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

      //try
          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qdetail.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qdetail.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qDetail.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qDetail.ParamByName(tempparam).asstring := PERIOD1
            else
              qDetail.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qDetail.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QDetail.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qDetail.ParamByName(tempparam).asstring := YEAR1
            else
              qDetail.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qdetail.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
        //except
        //       logfile('Isi qDetail tidak ada parameter :'+TEMPPARAM);
        //end;
          inc(i);
        end;
    //try
        if qDetail.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            qDetail.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            qDetail.MacroByName('kondisiUser').Value := ' ';
        end;
        tempall := tempall + gantisql(NOURUT, qDetail.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(qdetail.MasterSource);
    //  IF (PARENT_SQL[NOURUT]=AngkaNonAktif) or (not q.Eof) then  begin
        try
          qDetail.Active := true;
          Repstart.AddTable('qDetail', qDetail);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qDetail failed open ' + E.Message + '!');
            logFile('qDetail query : ' + qDetail.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qDetail : ' + qDetail.FinalSQL);
        if not qdetail.eof then
          DATAADA := true;
    //  end;
    //except
    //  messagedlg('maaf terdapat error di QDetail '+#10+#13+qDetail.sql.Text+#10+#13+'yang diisi pada QDetail '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
    //end;
    //qdetail.
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
     //logfile('jumlah isi detail '+inttostr(qdetail.RecordCount));
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qDetail.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

 {if not langsung then
   Begin
    //if (Not CItem.Checked and (Item1.Value <> '')) OR (Group1.Value <> '') then
    qHeader.sql.clear;
    qHeader.Active := false;
       qHeader.sql.Add(q1.SQL.Text);
       qHeader.Macrobyname('Kondisi').value := skondisiall[1];
    qHeader.Active := true;
    End;

  {  else
    Begin
     qHeader.sql.Add(q2.SQL.Text);
    End;
  End;
  // }
      tempall := tempall + 'Detail1: ' + #10;
      if qdetail1.SQL.text <> '' then
      begin
        NOURUT := 13;
        qDetail1.Active := false;
        if qDetail1.findparam('USER') <> nil then
          qDetail1.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qDetail1.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qDetail1.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail1.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qDetail1.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qDetail1.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail1.MacroByName('kondisiSales').Value := ' ';
        end;
        if qDetail1.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qDetail1.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail1.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qDetail1.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qDetail1.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail1.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_DETAIL1(NOURUT);
        qdetail1.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qdetail1.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qdetail1.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qDetail1.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail1.ParamByName(tempparam).asstring := PERIOD1
            else
              QDetail1.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qDetail1.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QDetail1.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail1.ParamByName(tempparam).asstring := YEAR1
            else
              QDetail1.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qdetail1.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
        if qDetail1.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            qDetail1.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            qDetail1.MacroByName('kondisiUser').Value := ' ';
        end;
    //try
        tempall := tempall + gantisql(NOURUT, qDetail1.FinalSQL) + #10 + #13;
     // q :=  tmyquery(qdetail1.MasterSource);
      //IF (PARENT_SQL[NOURUT]=AngkaNonAktif) or (not q.Eof) then begin
        try
          qDetail1.Active := true;
          Repstart.AddTable('qDetail1', qDetail1);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qDetail1 failed open ' + E.Message + '!');
            logFile('qDetail1 query : ' + qDetail1.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qDetail1 : ' + qDetail1.FinalSQL);
        if not qDetail1.eof then
          DATAADA := true;
      //end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
    //except
    //  messagedlg('maaf terdapat error di qDetail1 '+#10+#13+qDetail1.sql.Text+#10+#13+'yang diisi pada qDetail1 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
    //end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qDetail1.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'Detail2: ' + #10;
      if qdetail2.SQL.text <> '' then
      begin
        NOURUT := 14;
        qDetail2.Active := false;
        if qDetail2.findparam('USER') <> nil then
          qDetail2.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qDetail2.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qDetail2.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail2.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qDetail2.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qDetail2.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail2.MacroByName('kondisiSales').Value := ' ';
        end;
        if qDetail2.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qDetail2.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail2.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qDetail2.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qDetail2.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail2.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_DETAIL2(NOURUT);
        qDetail2.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qDetail2.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qDetail2.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qDetail2.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail2.ParamByName(tempparam).asstring := PERIOD1
            else
              QDetail2.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qDetail2.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QDetail2.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail2.ParamByName(tempparam).asstring := YEAR1
            else
              QDetail2.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qDetail2.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
        if qDetail2.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            qDetail2.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            qDetail2.MacroByName('kondisiUser').Value := ' ';
        end;
    //try
        tempall := tempall + gantisql(NOURUT, qDetail2.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(qdetail2.MasterSource);
    //  IF (PARENT_SQL[NOURUT]=AngkaNonAktif) or (not q.Eof) then begin
        try
          qDetail2.Active := true;
          Repstart.AddTable('qDetail2', qDetail2);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qDetail2 failed open ' + E.Message + '!');
            logFile('qDetail2 query : ' + qDetail2.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qDetail2 : ' + qDetail2.FinalSQL);
        if not qDetail2.eof then
          DATAADA := true;
    //  end;
    //except
    //  messagedlg('maaf terdapat error di qDetail2 '+#10+#13+qDetail2.sql.Text+#10+#13+'yang diisi pada qDetail2 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
    //end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qDetail2.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'Detail3: ' + #10;
      if qdetail3.SQL.text <> '' then
      begin
        NOURUT := 15;
        qDetail3.Active := false;
        if qDetail3.findparam('USER') <> nil then
          qDetail3.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qDetail3.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qDetail3.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail3.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qDetail3.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qDetail3.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail3.MacroByName('kondisiSales').Value := ' ';
        end;
        if qDetail3.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qDetail3.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail3.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qDetail3.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qDetail3.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qDetail3.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_DETAIL3(NOURUT);
        qDetail3.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qDetail3.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qDetail3.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qDetail3.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail3.ParamByName(tempparam).asstring := PERIOD1
            else
              QDetail3.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qDetail3.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QDetail3.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail3.ParamByName(tempparam).asstring := YEAR1
            else
              QDetail3.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qDetail3.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, qDetail3.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(qdetail3.MasterSource);
   //   IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          qDetail3.Active := true;
          Repstart.AddTable('qDetail3', qDetail3);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qDetail3 failed open ' + E.Message + '!');
            logFile('qDetail3 query : ' + qDetail3.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qDetail3 : ' + qDetail3.FinalSQL);
        if not qDetail3.eof then
          DATAADA := true;
   //   end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di qDetail3 '+#10+#13+qDetail3.sql.Text+#10+#13+'yang diisi pada qDetail3 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qDetail3.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'Detail4: ' + #10;
      if qdetail4.SQL.text <> '' then
      begin
        NOURUT := 16;
        qDetail4.Active := false;
        if qdetail4.findparam('USER') <> nil then
          qdetail4.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qdetail4.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qdetail4.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qdetail4.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qdetail4.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qdetail4.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qdetail4.MacroByName('kondisiSales').Value := ' ';
        end;
        if qdetail4.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qdetail4.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qdetail4.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qdetail4.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qdetail4.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qdetail4.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_DETAIL4(NOURUT);
        qDetail4.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qDetail4.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qDetail4.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qDetail4.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail4.ParamByName(tempparam).asstring := PERIOD1
            else
              QDetail4.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qDetail4.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QDetail4.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QDetail4.ParamByName(tempparam).asstring := YEAR1
            else
              QDetail4.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qDetail4.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, qDetail4.FinalSQL) + #10 + #13;
      //q :=  tmyquery(qdetail4.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          qDetail4.Active := true;
          Repstart.AddTable('qDetail4', qDetail4);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qDetail4 failed open ' + E.Message + '!');
            logFile('qDetail4 query : ' + qDetail4.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qDetail4 : ' + qDetail4.FinalSQL);
        if not qDetail4.eof then
          DATAADA := true;
      //end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di qDetail4 '+#10+#13+qDetail4.sql.Text+#10+#13+'yang diisi pada qDetail4 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qDetail4.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'No_Form: ' + #10;
      if qNo_form.SQL.text <> '' then
      begin
        NOURUT := 4;
        qNo_Form.Active := false;
        if qNo_Form.findparam('USER') <> nil then
          qNo_Form.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if qNo_Form.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            qNo_Form.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qNo_Form.MacroByName('kondisiCabang').Value := ' ';
        end;
        if qNo_Form.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            qNo_Form.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qNo_Form.MacroByName('kondisiSales').Value := ' ';
        end;
        if qNo_Form.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            qNo_Form.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qNo_Form.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if qNo_Form.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            qNo_Form.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            qNo_Form.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_No_Form(NOURUT);
        qNo_Form.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              qNo_Form.ParamByName(tempparam).asstring := YEARMONTH1
            else
              qNo_Form.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            qNo_Form.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QNo_Form.ParamByName(tempparam).asstring := PERIOD1
            else
              QNo_Form.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            qNo_Form.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            QNo_Form.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              QNo_Form.ParamByName(tempparam).asstring := YEAR1
            else
              QNo_Form.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            qNo_Form.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, qNo_Form.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(qNo_Form.MasterSource);
    //  IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          qNo_Form.Active := true;
          Repstart.AddTable('qNo_Form', qNo_Form);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('qNo_Form failed open ' + E.Message + '!');
            logFile('qNo_Form query : ' + qNo_Form.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('qNo_Form : ' + qNo_Form.FinalSQL);
        if not qNo_Form.eof then
          DATAADA := true;
     // end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di qNo_Form '+#10+#13+qNo_Form.sql.Text+#10+#13+'yang diisi pada qNo_Form '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(qNo_form.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q1: ' + #10;
      if q1.SQL.text <> '' then
      begin
        NOURUT := 1;
        q1.Active := false;
        if q1.findparam('USER') <> nil then
          q1.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q1.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q1.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q1.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q1.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q1.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q1.MacroByName('kondisiSales').Value := ' ';
        end;
        if q1.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q1.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q1.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q1.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q1.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q1.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_1(NOURUT);
        q1.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);

      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q1.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q1.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q1.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q1.ParamByName(tempparam).asstring := PERIOD1
            else
              q1.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q1.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q1.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q1.ParamByName(tempparam).asstring := YEAR1
            else
              q1.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q1.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
        if q1.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            q1.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            q1.MacroByName('kondisiUser').Value := ' ';
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q1.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(q1.MasterSource);
    //  IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          q1.Active := true;
          Repstart.AddTable('q1', q1);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q1 failed open ' + E.Message + '!');
            logFile('q1 query : ' + q1.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q1 : ' + q1.FinalSQL);
        if not q1.eof then
          DATAADA := true;
     // end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q1 '+#10+#13+q1.sql.Text+#10+#13+'yang diisi pada q1 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q1.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;

      end;
      tempall := tempall + 'q2: ' + #10;
      if q2.SQL.text <> '' then
      begin
        NOURUT := 2;
        q2.Active := false;
        if q2.findparam('USER') <> nil then
          q2.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q2.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q2.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q2.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q2.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q2.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q2.MacroByName('kondisiSales').Value := ' ';
        end;
        if q2.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q2.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q2.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q2.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q2.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q2.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_2(NOURUT);
        q2.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q2.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q2.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q2.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q2.ParamByName(tempparam).asstring := PERIOD1
            else
              q2.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q2.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q2.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q2.ParamByName(tempparam).asstring := YEAR1
            else
              q2.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q2.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
        if q2.FindMacro('kondisiUser') <> nil then
        begin
          if NOT_BANK_all_flag then
            q2.MacroByName('kondisiUser').Value := 'and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' '
          else
            q2.MacroByName('kondisiUser').Value := ' ';
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q2.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q2.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          q2.Active := true;
          Repstart.AddTable('q2', q2);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q2 failed open ' + E.Message + '!');
            logFile('q2 query : ' + q2.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q2 : ' + q2.FinalSQL);
        if not q2.eof then
          DATAADA := true;
      //end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q2 '+#10+#13+q2.sql.Text+#10+#13+'yang diisi pada q2 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q2.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q3: ' + #10;
      if q3.SQL.text <> '' then
      begin
        NOURUT := 10;
        q3.Active := false;
        if q3.findparam('USER') <> nil then
          q3.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q3.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q3.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q3.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q3.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q3.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q3.MacroByName('kondisiSales').Value := ' ';
        end;
        if q3.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q3.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q3.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q3.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q3.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q3.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_3(NOURUT);
        q3.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q3.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q3.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q3.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q3.ParamByName(tempparam).asstring := PERIOD1
            else
              q3.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q3.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q3.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q3.ParamByName(tempparam).asstring := YEAR1
            else
              q3.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q3.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q3.FinalSQL) + #10 + #13;
    //  q :=  tmyquery(q3.MasterSource);
     // IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          q3.Active := true;
          Repstart.AddTable('q3', q3);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q3 failed open ' + E.Message + '!');
            logFile('q3 query : ' + q3.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q3 : ' + q3.FinalSQL);
        if not q3.eof then
          DATAADA := true;
     // end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q3 '+#10+#13+q3.sql.Text+#10+#13+'yang diisi pada q3 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q3.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;

      end;
      tempall := tempall + 'q4: ' + #10;
      if q4.SQL.text <> '' then
      begin
        NOURUT := 11;
        q4.Active := false;
        if q4.findparam('USER') <> nil then
          q4.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q4.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q4.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q4.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q4.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q4.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q4.MacroByName('kondisiSales').Value := ' ';
        end;
        if q4.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q4.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q4.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q4.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q4.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q4.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_4(NOURUT);
        q4.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q4.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q4.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q4.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q4.ParamByName(tempparam).asstring := PERIOD1
            else
              q4.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q4.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q4.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q4.ParamByName(tempparam).asstring := YEAR1
            else
              q4.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q4.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q4.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q4.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then begin
        try
          q4.Active := true;
          Repstart.AddTable('q4', q4);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q4 failed open ' + E.Message + '!');
            logFile('q4 query : ' + q4.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q4 : ' + q4.FinalSQL);
        if not q4.eof then
          DATAADA := true;
     // end;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q4 '+#10+#13+q4.sql.Text+#10+#13+'yang diisi pada q4 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q4.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q5: ' + #10;
      if q5.SQL.text <> '' then
      begin
        NOURUT := 5;
        q5.Active := false;
        if q5.findparam('USER') <> nil then
          q5.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q5.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q5.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q5.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q5.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q5.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q5.MacroByName('kondisiSales').Value := ' ';
        end;
        if q5.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q5.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q5.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q5.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q5.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q5.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_5(NOURUT);
        q5.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q5.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q5.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q5.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q5.ParamByName(tempparam).asstring := PERIOD1
            else
              q5.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q5.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q5.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q5.ParamByName(tempparam).asstring := YEAR1
            else
              q5.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q5.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q5.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q5.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q5.Active := true;
          Repstart.AddTable('q5', q5);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q5 failed open ' + E.Message + '!');
            logFile('q5 query : ' + q5.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q5 : ' + q5.FinalSQL);
        if not q5.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q5 '+#10+#13+q5.sql.Text+#10+#13+'yang diisi pada q5 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q5.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q6: ' + #10;
      if q6.SQL.text <> '' then
      begin
        NOURUT := 6;
        q6.Active := false;
        if q6.findparam('USER') <> nil then
          q6.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q6.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q6.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q6.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q6.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q6.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q6.MacroByName('kondisiSales').Value := ' ';
        end;
        if q6.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q6.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q6.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q6.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q6.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q6.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_6(NOURUT);
        q6.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q6.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q6.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q6.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q6.ParamByName(tempparam).asstring := PERIOD1
            else
              q6.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q6.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q6.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q6.ParamByName(tempparam).asstring := YEAR1
            else
              q6.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q6.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q6.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q6.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q6.Active := true;
          Repstart.AddTable('q6', q6);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q6 failed open ' + E.Message + '!');
            logFile('q6 query : ' + q6.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q6 : ' + q6.FinalSQL);
        if not q6.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q6 '+#10+#13+q6.sql.Text+#10+#13+'yang diisi pada q6 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q6.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q7: ' + #10;
      if q7.SQL.text <> '' then
      begin
        NOURUT := 7;
        q7.Active := false;
        if q7.findparam('USER') <> nil then
          q7.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q7.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q7.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q7.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q7.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q7.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q7.MacroByName('kondisiSales').Value := ' ';
        end;
        if q7.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q7.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q7.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q7.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q7.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q7.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_7(NOURUT);
        q7.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q7.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q7.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q7.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q7.ParamByName(tempparam).asstring := PERIOD1
            else
              q7.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q7.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q7.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q7.ParamByName(tempparam).asstring := YEAR1
            else
              q7.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q7.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q7.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q7.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q7.Active := true;
          Repstart.AddTable('q7', q7);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q7 failed open ' + E.Message + '!');
            logFile('q7 query : ' + q7.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q7 : ' + q7.FinalSQL);
        if not q7.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q7 '+#10+#13+q7.sql.Text+#10+#13+'yang diisi pada q7 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q7.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q8: ' + #10;
      if q8.SQL.text <> '' then
      begin
        NOURUT := 8;
        q8.Active := false;
        if q8.findparam('USER') <> nil then
          q8.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q8.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q8.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q8.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q8.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q8.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q8.MacroByName('kondisiSales').Value := ' ';
        end;
        if q8.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q8.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q8.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q8.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q8.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q8.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_8(NOURUT);
        q8.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q8.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q8.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q8.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q8.ParamByName(tempparam).asstring := PERIOD1
            else
              q8.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q8.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q8.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q8.ParamByName(tempparam).asstring := YEAR1
            else
              q8.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q8.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q8.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q8.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q8.Active := true;
          Repstart.AddTable('q8', q8);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q8 failed open ' + E.Message + '!');
            logFile('q8 query : ' + q8.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q8 : ' + q8.FinalSQL);
        if not q8.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q8 '+#10+#13+q8.sql.Text+#10+#13+'yang diisi pada q8 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q8.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;
      tempall := tempall + 'q9: ' + #10;
      if q9.SQL.text <> '' then
      begin
        NOURUT := 12;
        q9.Active := false;
        if q9.findparam('USER') <> nil then
          q9.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q9.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q9.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q9.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q9.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q9.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q9.MacroByName('kondisiSales').Value := ' ';
        end;
        if q9.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q9.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q9.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q9.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q9.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q9.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_9(NOURUT);
        q9.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q9.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q9.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q9.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q9.ParamByName(tempparam).asstring := PERIOD1
            else
              q9.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q9.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q9.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q9.ParamByName(tempparam).asstring := YEAR1
            else
              q9.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q9.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q9.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q9.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q9.Active := true;
          Repstart.AddTable('q9', q9);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q9 failed open ' + E.Message + '!');
            logFile('q9 query : ' + q9.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q9 : ' + q9.FinalSQL);
        if not q9.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q9 '+#10+#13+q9.sql.Text+#10+#13+'yang diisi pada q9 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q9.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q10: ' + #10;
      if q10.SQL.text <> '' then
      begin
        NOURUT := 17;
        q10.Active := false;
        if q10.findparam('USER') <> nil then
          q10.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q10.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q10.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q10.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q10.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q10.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q10.MacroByName('kondisiSales').Value := ' ';
        end;
        if q10.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q10.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q10.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q10.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q10.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q10.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_10(NOURUT);
        q10.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q10.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q10.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q10.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q10.ParamByName(tempparam).asstring := PERIOD1
            else
              q10.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q10.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q10.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q10.ParamByName(tempparam).asstring := YEAR1
            else
              q10.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q10.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q10.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q10.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q10.Active := true;
          Repstart.AddTable('q10', q10);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q10 failed open ' + E.Message + '!');
            logFile('q10 query : ' + q10.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q10 : ' + q10.FinalSQL);
        if not q10.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q10 '+#10+#13+q10.sql.Text+#10+#13+'yang diisi pada q10 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q10.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q11: ' + #10;
      if q11.SQL.text <> '' then
      begin
        NOURUT := 18;
        q11.Active := false;
        if q11.findparam('USER') <> nil then
          q11.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q11.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q11.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q11.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q11.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q11.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q11.MacroByName('kondisiSales').Value := ' ';
        end;
        if q11.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q11.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q11.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q11.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q11.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q11.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_11(NOURUT);
        q11.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q11.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q11.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q11.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q11.ParamByName(tempparam).asstring := PERIOD1
            else
              q11.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q11.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q11.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q11.ParamByName(tempparam).asstring := YEAR1
            else
              q11.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q11.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q11.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q11.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q11.Active := true;
          Repstart.AddTable('q11', q11);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q11 failed open ' + E.Message + '!');
            logFile('q11 query : ' + q11.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q11 : ' + q11.FinalSQL);
        if not q11.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q11 '+#10+#13+q11.sql.Text+#10+#13+'yang diisi pada q11 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q11.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q12: ' + #10;
      if q12.SQL.text <> '' then
      begin
        NOURUT := 19;
        q12.Active := false;
        if q12.findparam('USER') <> nil then
          q12.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q12.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q12.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q12.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q12.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q12.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q12.MacroByName('kondisiSales').Value := ' ';
        end;
        if q12.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q12.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q12.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q12.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q12.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q12.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_12(NOURUT);
        q12.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q12.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q12.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q12.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q12.ParamByName(tempparam).asstring := PERIOD1
            else
              q12.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q12.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q12.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q12.ParamByName(tempparam).asstring := YEAR1
            else
              q12.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q12.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q12.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q12.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q12.Active := true;
          Repstart.AddTable('q12', q12);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q12 failed open ' + E.Message + '!');
            logFile('q12 query : ' + q12.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q12 : ' + q12.FinalSQL);
        if not q12.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q12 '+#10+#13+q12.sql.Text+#10+#13+'yang diisi pada q12 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q12.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q14: ' + #10;
      if q14.SQL.text <> '' then
      begin
        NOURUT := 20;
        q14.Active := false;
        if q14.findparam('USER') <> nil then
          q14.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q14.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q14.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q14.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q14.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q14.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q14.MacroByName('kondisiSales').Value := ' ';
        end;
        if q14.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q14.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q14.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q14.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q14.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q14.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_14(NOURUT);
        q14.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q14.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q14.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q14.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q14.ParamByName(tempparam).asstring := PERIOD1
            else
              q14.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q14.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q14.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q14.ParamByName(tempparam).asstring := YEAR1
            else
              q14.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q14.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q14.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q14.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q14.Active := true;
          Repstart.AddTable('q14', q14);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q14 failed open ' + E.Message + '!');
            logFile('q14 query : ' + q14.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q14 : ' + q14.FinalSQL);
        if not q14.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q14 '+#10+#13+q14.sql.Text+#10+#13+'yang diisi pada q14 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q14.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q15: ' + #10;
      if q15.SQL.text <> '' then
      begin
        NOURUT := 21;
        q15.Active := false;
        if q15.findparam('USER') <> nil then
          q15.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q15.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q15.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q15.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q15.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q15.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q15.MacroByName('kondisiSales').Value := ' ';
        end;
        if q15.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q15.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q15.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q15.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q15.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q15.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_15(NOURUT);
        q15.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q15.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q15.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q15.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q15.ParamByName(tempparam).asstring := PERIOD1
            else
              q15.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q15.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q15.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q15.ParamByName(tempparam).asstring := YEAR1
            else
              q15.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q15.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q15.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q15.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q15.Active := true;
          Repstart.AddTable('q15', q15);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q15 failed open ' + E.Message + '!');
            logFile('q15 query : ' + q15.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q15 : ' + q15.FinalSQL);
        if not q15.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q15 '+#10+#13+q15.sql.Text+#10+#13+'yang diisi pada q15 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q15.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q16: ' + #10;
      if q16.SQL.text <> '' then
      begin
        NOURUT := 22;
        q16.Active := false;
        if q16.findparam('USER') <> nil then
          q16.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q16.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q16.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q16.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q16.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q16.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q16.MacroByName('kondisiSales').Value := ' ';
        end;
        if q16.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q16.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q16.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q16.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q16.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q16.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_16(NOURUT);
        q16.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q16.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q16.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q16.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q16.ParamByName(tempparam).asstring := PERIOD1
            else
              q16.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q16.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q16.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q16.ParamByName(tempparam).asstring := YEAR1
            else
              q16.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q16.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q16.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q16.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q16.Active := true;
          Repstart.AddTable('q16', q16);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q16 failed open ' + E.Message + '!');
            logFile('q16 query : ' + q16.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q16 : ' + q16.FinalSQL);
        if not q16.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q16 '+#10+#13+q16.sql.Text+#10+#13+'yang diisi pada q16 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q16.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q17: ' + #10;
      if q17.SQL.text <> '' then
      begin
        NOURUT := 23;
        q17.Active := false;
        if q17.findparam('USER') <> nil then
          q17.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q17.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q17.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q17.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q17.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q17.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q17.MacroByName('kondisiSales').Value := ' ';
        end;
        if q17.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q17.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q17.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q17.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q17.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q17.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_17(NOURUT);
        q17.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q17.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q17.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q17.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q17.ParamByName(tempparam).asstring := PERIOD1
            else
              q17.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q17.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q17.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q17.ParamByName(tempparam).asstring := YEAR1
            else
              q17.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q17.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q17.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q17.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q17.Active := true;
          Repstart.AddTable('q17', q17);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q17 failed open ' + E.Message + '!');
            logFile('q17 query : ' + q17.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q17 : ' + q17.FinalSQL);
        if not q17.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q17 '+#10+#13+q17.sql.Text+#10+#13+'yang diisi pada q17 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q17.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q18: ' + #10;
      if q18.SQL.text <> '' then
      begin
        NOURUT := 24;
        q18.Active := false;
        if q18.findparam('USER') <> nil then
          q18.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q18.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q18.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q18.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q18.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q18.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q18.MacroByName('kondisiSales').Value := ' ';
        end;
        if q18.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q18.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q18.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q18.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q18.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q18.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_18(NOURUT);
        q18.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q18.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q18.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q18.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q18.ParamByName(tempparam).asstring := PERIOD1
            else
              q18.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q18.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q18.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q18.ParamByName(tempparam).asstring := YEAR1
            else
              q18.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q18.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q18.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q18.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q18.Active := true;
          Repstart.AddTable('q18', q18);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q18 failed open ' + E.Message + '!');
            logFile('q18 query : ' + q18.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q18 : ' + q18.FinalSQL);
        if not q18.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q18 '+#10+#13+q18.sql.Text+#10+#13+'yang diisi pada q18 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q18.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q19: ' + #10;
      if q19.SQL.text <> '' then
      begin
        NOURUT := 25;
        q19.Active := false;
        if q19.findparam('USER') <> nil then
          q19.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q19.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q19.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q19.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q19.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q19.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q19.MacroByName('kondisiSales').Value := ' ';
        end;
        if q19.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q19.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q19.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q19.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q19.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q19.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_19(NOURUT);
        q19.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q19.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q19.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q19.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q19.ParamByName(tempparam).asstring := PERIOD1
            else
              q19.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q19.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q19.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q19.ParamByName(tempparam).asstring := YEAR1
            else
              q19.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q19.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q19.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q19.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q19.Active := true;
          Repstart.AddTable('q19', q19);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q19 failed open ' + E.Message + '!');
            logFile('q19 query : ' + q19.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q19 : ' + q19.FinalSQL);
        if not q19.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q19 '+#10+#13+q19.sql.Text+#10+#13+'yang diisi pada q19 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q19.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      tempall := tempall + 'q20: ' + #10;
      if q20.SQL.text <> '' then
      begin
        NOURUT := 26;
        q20.Active := false;
        if q20.findparam('USER') <> nil then
          q20.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
        if q20.FindMacro('kondisiCabang') <> nil then
        begin
          if branch_id <> '1001' then
            q20.MacroByName('kondisiCabang').Value := ' and  gggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q20.MacroByName('kondisiCabang').Value := ' ';
        end;
        if q20.FindMacro('kondisiSales') <> nil then
        begin
          if NOT_SALES_all_flag then
            q20.MacroByName('kondisiSales').Value := ' and  ggg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q20.MacroByName('kondisiSales').Value := ' ';
        end;
        if q20.FindMacro('kondisiDivisi') <> nil then
        begin
          if NOT_DIVISI_all_flag then
            q20.MacroByName('kondisiDivisi').Value := ' and  gg.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q20.MacroByName('kondisiDivisi').Value := ' ';
        end;
        if q20.FindMacro('kondisiMerek') <> nil then
        begin
          if NOT_MEREK_all_flag then
            q20.MacroByName('kondisiMerek').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString
          else
            q20.MacroByName('kondisiMerek').Value := ' ';
        end;
        PARENT_SQL_20(NOURUT);
        q20.Macrobyname('Kondisi').Value := skondisiall[NOURUT];

        i := 0;
        while i < JUM_PARAM_YG_DIAMBIL[NOURUT] do
        begin
          tempparam := midstr(PARAM_YG_DIAMBIL[NOURUT][i], 2, length(PARAM_YG_DIAMBIL[NOURUT][i]) - 1);
      {fatra 11 nov 2020
      jika ketemu 13 diganti ke 12
      jika ketemu 23 diganti ke 22
       }
          tempparam_baru := tempparam;
//          if (tempparam = '13') and not variable_periode_flag then
//            tempparam_baru := '12';
//          if (tempparam = '23') and not variable_periode_flag then
//            tempparam_baru := '22';

          if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEARMONTH' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q20.ParamByName(tempparam).asstring := YEARMONTH1
            else
              q20.ParamByName(tempparam).asstring := YEARMONTH2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 9) = 'YEAR_LALU' then
          begin
            q20.ParamByName(tempparam).asstring := YEAR_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_NAME' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q20.ParamByName(tempparam).asstring := PERIOD1
            else
              q20.ParamByName(tempparam).asstring := PERIOD2;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_AWAL' then
          begin
            q20.ParamByName(tempparam).asstring := PERIOD_AWAL;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 11) = 'PERIOD_LALU' then
          begin
            Q20.ParamByName(tempparam).asstring := PERIOD_LALU;
          end
          else if LEFTSTR(UPPERCASE(tempparam), 4) = 'YEAR' then
          begin
            if RIGHTSTR(tempparam, 1) = '1' then
              q20.ParamByName(tempparam).asstring := YEAR1
            else
              q20.ParamByName(tempparam).asstring := YEAR2;
          end
          else
            q20.ParamByName(tempparam).asstring := (v_a_param_yg_dipasing[strtoint(leftstr(tempparam_baru, 1)), strtoint(rightstr(tempparam_baru, length(tempparam_baru) - 1))]);
          inc(i);
        end;
//    try
        tempall := tempall + gantisql(NOURUT, q20.FinalSQL) + #10 + #13;
      //q :=  tmyquery(q20.MasterSource);
      //IF (PARENT_SQL[NOURUT]= AngkaNonAktif ) or (not q.Eof) then
        try
          q20.Active := true;
          Repstart.AddTable('q20', q20);
        except
          on E: exception do
          begin
            Result := False;
            vErr := E.Message;
            logFile('q20 failed open ' + E.Message + '!');
            logFile('q20 query : ' + q20.FinalSQL + '!');
            Exit;
          end;
        end;
        logFile('q20 : ' + q20.FinalSQL);
        if not q20.eof then
          DATAADA := true;
        if PARENT_SQL[NOURUT] <> AngkaNonAktif then
          jumlahrecordcount := jumlahrecordcount + hitungsql(parent_sql[NOURUT]) * hitungsql(NOURUT) * banyakbaris[NOURUT]
        else
          jumlahrecordcount := jumlahrecordcount + hitungsql(NOURUT) * banyakbaris[NOURUT];
//    except
//      messagedlg('maaf terdapat error di q20 '+#10+#13+q20.sql.Text+#10+#13+'yang diisi pada q20 '+#10+#13+skondisiall[NOURUT],mtinformation,[mbOk],0);
//    end;
        tempall := tempall + ' JUMLAH RECORD : ' + inttostr(q20.RecordCount) + #10;
        tempall := tempall + '=====================================' + #10 + #13;
      end;

      Extensi := '.xlsx';

      extensi1 := extensi;

      if not DirectoryExists(patch + 'Report\') then
        mkdir(patch + 'Report\');
      file_excel := AnsiReplaceStr(qSql.fieldbyname('R_DISPLAY').asstring, '\', '');
      file_excel := AnsiReplaceStr(file_excel, '/', '');
      file_excel := AnsiReplaceStr(file_excel, ' ', '_');

      file_excel := patch + 'Report\' + file_excel;
      if not DirectoryExists(patch + 'Report\') then
      begin
        if ForceDirectories(patch + 'Report\') then
          logFile('Folder Report dibuat: ' + patch + 'Report\')
        else
        begin
          logFile('Gagal membuat folder Report: ' + patch + 'Report\');
          BuatFolderUnik(patch + 'Report\');
        end;
      end;
      Filename := file_excel + qReportWebAPI_REPORT_ID.AsString + extensi1;
      if FileExists(Filename) then
        deletefile(Filename);

      for i := 0 to 30 do
      begin
        try
          Repstart.SetValue('Var1_' + IntToStr(i), varTostr(FArray[i]));

        except
          on E: Exception do
            logFile('variant ' + IntToStr(i) + ' err ' + E.Message)
        end;
      end;

      logFile(nama_template + '.xlsx'); // iksan
      logFile(Filename); // iksan
      Repstart.Run(nama_template + '.xlsx', Filename);
      if FileExists(Filename) then
      begin
      // ===== Tentukan folder & nama file PDF tujuan (aman di main thread) =====
        vDbFolder := AnsiReplaceStr(qReportWebDATABASE_NAME.AsString, '/', '');
        vDbFolder := AnsiReplaceStr(vDbFolder, '\', '');
        vDbFolder := AnsiReplaceStr(vDbFolder, ' ', '_');
        vDbFolder := IncludeTrailingPathDelimiter(vOutPutFile + '\' + vDbFolder);

        vReportFolder := IncludeTrailingPathDelimiter(vDbFolder + 'Report');
        logFile('Check Folder ' + vReportFolder);

        if not DirectoryExists(vReportFolder) then
        begin
          if ForceDirectories(vReportFolder) then
            logFile('Folder Report dibuat: ' + vReportFolder)
          else
          begin
            logFile('Gagal membuat folder Report: ' + vReportFolder);
            BuatFolderUnik(vReportFolder);
          end;
        end;
        CopyMyFile(Filename, vReportFolder + ExtractFileName(Filename));

        vFilePdf := ExtractFileName(StringReplace(Filename, '.xlsx', '', [rfReplaceAll, rfIgnoreCase]));
        vFilePdf := vReportFolder + vFilePdf + '.pdf';

        ProcessExcelToPdf(Filename, vFilePdf);
        file_excel := Filename;
      end;
    except
      on E: Exception do
      begin
        Result := False;
        vErr := E.Message;
        logFile(' create excel failed ' + E.Message + '!');
        Exit;
      end;

    end;
  finally
    Repstart.Free;
  end;

  //if not dataada then raise Exception.CreateFmt('Maaf, data tidak ada',[]);
  anu := 'Maaf, data record yang dihasilkan terlalu besar' + #10 + 'harap diperinci lagi';
 // if jumlahrecordcount > 55000 then raise Exception.CreateFmt(anu,[]);
  //Proses_data1 :=  Not qHeader.Eof;
  proses_data1 := DATAADA;
  //proses_data1 := true;
end;

procedure TsvReport.qErpDetailAfterScroll(DataSet: TDataSet);
begin
  logFile('Proses parameter ' + qErpDetailERP_DETAIL_ID.AsString + ' kolom ' + qErpDetailWHERE_COLUMN.AsString + ' value ' + qErpDetailWHERE_VALUE.AsString);
end;

procedure TsvReport.ProcessExcelToPdf(const InFile, OutFile: string);
var
  Xls: TXlsFile;
  PdfExport: TFlexCelPdfExport;
  PdfStream: TFileStream;
  LogPath, ExcelPath, PdfPath: string;
begin
  // Thread-safe and UI-independent allocation
  logFile('Proses convert pdf menjadi ' + OutFile);
  Xls := TXlsFile.Create(InFile, True);
  PdfPath := OutFile;

  try

    PdfExport := TFlexCelPdfExport.Create(Xls, true);
    PdfStream := TFileStream.Create(PdfPath, fmCreate);
    try
      PdfExport.BeginExport(PdfStream);
      PdfExport.ExportAllVisibleSheets(false, 'Service Report');
      PdfExport.EndExport;
    finally
      PdfStream.Free;
      PdfExport.Free;
    end;
  finally
    Xls.Free;
  end;
end;

function TsvReport.createExcel(vReportId: string): Boolean;     //isi_sql
begin
  Result := False;
  logFile('proses createExcel');
  try
    qSql.Active := False;
    qSql.ParamByName('ERP_RPT_ID').AsString := vReportId;
    qSql.Active := True;
  except
    on E: exception do
    begin
      logFile('qSql open Error ' + E.Message);
      Exit;
    end;
  end;

//  logFile('Proses Define_Sql');
  if not Define_sql then
  begin
    Exit;
  end;
//  logFile('selesai Define_Sql');
  nama_template := defa_template + qSql.fieldbyname('R_FILE').asstring;
  if not FileExists(defa_template + qSql.fieldbyname('R_FILE').asstring + '.xlsx') then
  begin
    logFile('Filte template not found in ' + nama_template);
    nama_template := patch + 'default\template\' + qSql.fieldbyname('R_FILE').asstring;
  end;
//  logFile(Repstart.Template);

  if not Run_sql(' set @no = 0') then
    Exit;
  inisialisasi_v;
  logFile('Proses Proses data');
//  Run_queryManage('SELECT API_REPORT_ID, GENERATE_FLAG FROM API_REPORT WHERE API_REPORT_ID=' + qReportWebAPI_REPORT_ID.AsString);
//  if Query2.FieldByName('GENERATE_FLAG').AsString <> '2' then
  begin
//    if Query2.Active then
//      Query2.Close;
    if Proses_data1 then
    begin
      logFile('Proses convert to excel');
    end
    else
    begin
      if vErr = '' then
        logFile('Data Tidak Ada')
      else
      begin
        Exit;
      end;
    end;
  end;
  Result := True;
end;

procedure TsvReport.db1AfterConnect(Sender: TObject);
begin
  db := '''';
  mulai;
  SetInitAwalModule;
end;

procedure TsvReport.ProcessOneQueueItem;
var
  vTemplate, vOutfileExcel, vERP_RPT_DETAIL_ID, vFileFtp, vDatabaseFolder, vHariDel: string;
  js: TJSONValue;
  jsArr, WHERE: TJSONArray;
  Obj, WhereObj: TJSONObject;
  I, J: Integer;

  procedure EnsureDirectoryExists(RemoteDir: string);
  var
    Parts: TStringList;
    CurrentDir: string;
    K: Integer;
  begin
    Parts := TStringList.Create;
    try
      ExtractStrings(['/'], [], PChar(RemoteDir), Parts);
      CurrentDir := '';
      for K := 0 to Parts.Count - 1 do
      begin
        CurrentDir := CurrentDir + '/' + Parts[K];
        try
          IdFTP.ChangeDir(CurrentDir);
        except
          on E: Exception do
          begin
            if pos('Failed', E.Message) > 0 then
            begin
              IdFTP.MakeDir(CurrentDir);
              IdFTP.ChangeDir(CurrentDir);
            end
            else
              raise;
          end;
        end;
      end;
    finally
      Parts.Free;
    end;
  end;

  function UploadSingleFile(aDirectory: string): Boolean;
  var
    aFile, RemoteDir: string;
  begin
    Result := True;
    aFile := ExtractFileName(aDirectory);
    RemoteDir := '/api_report/' + qReportWebDATABASE_NAME.AsString + '/';
    vFileFtp := 'https://seventhsoft.net' + RemoteDir + aFile;
    try
      EnsureDirectoryExists(RemoteDir);
      IdFTP.Put(aDirectory, RemoteDir + aFile);
      logFile('berhasil upload di ' + vFileFtp);
    except
      on E: Exception do
      begin
        logFile(E.Message);
        Result := False;
      end;
    end;
  end;

  function ExecuteCommand(const Command: string; Output: TStrings): Boolean;
  var
    SecurityAttrs: TSecurityAttributes;
    ReadPipe, WritePipe: THandle;
    StartupInfo: TStartupInfo;
    ProcessInfo: TProcessInformation;
    Buffer: array[0..255] of AnsiChar;
    BytesRead: DWORD;
  begin
    Result := False;
    Output.Clear;

    SecurityAttrs.nLength := SizeOf(TSecurityAttributes);
    SecurityAttrs.bInheritHandle := True;
    SecurityAttrs.lpSecurityDescriptor := nil;

    if not CreatePipe(ReadPipe, WritePipe, @SecurityAttrs, 0) then
      Exit;

    FillChar(ProcessInfo, SizeOf(TProcessInformation), 0);
    FillChar(StartupInfo, SizeOf(TStartupInfo), 0);
    StartupInfo.cb := SizeOf(TStartupInfo);
    StartupInfo.hStdError := WritePipe;
    StartupInfo.hStdOutput := WritePipe;
    StartupInfo.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
    StartupInfo.wShowWindow := SW_HIDE;

    if CreateProcess(nil, PChar(Command), @SecurityAttrs, @SecurityAttrs, True, CREATE_NO_WINDOW or NORMAL_PRIORITY_CLASS, nil, nil, StartupInfo, ProcessInfo) then
    begin
      CloseHandle(WritePipe);
      while ReadFile(ReadPipe, Buffer, SizeOf(Buffer), BytesRead, nil) do
      begin
        if BytesRead > 0 then
          Output.Add(Copy(Buffer, 1, BytesRead));
      end;
      Result := True;

      CloseHandle(ProcessInfo.hProcess);
      CloseHandle(ProcessInfo.hThread);
      CloseHandle(ReadPipe);
    end;
  end;

  function connectFtp(): Boolean;
  var
    Command: AnsiString;
    RuleName, AppPath: string;
    Output: TStringList;
  begin
    logFile('Proses connecting server ' + vFTP);
    Result := True;
    Output := TStringList.Create;
    RuleName := LowerCase(ExtractFileName(patch));
    try
      IdFTP.Host := vFTP;
      IdFTP.Username := 'seventhsoft';
      IdFTP.Password := '93vc487c34jfjnfc45n';

      IdFTP.Disconnect;
      IdFTP.Connect;
      try
        if not (ExecuteCommand('netsh advfirewall firewall show rule name="' + RuleName + '"', Output) and (pos(RuleName, Output.Text) > 0)) then
        begin
          Command := AnsiString(Format('netsh advfirewall firewall add rule name="%s" dir=in action=allow program="%s" enable=yes', [RuleName, AppPath]));
          WinExec(PAnsiChar(Command), SW_HIDE);
          Sleep(1000);

          if WinExec(PAnsiChar(Command), SW_SHOWNORMAL) <= 31 then
            raise Exception.Create('Failed to add firewall exception.');
        end;
      finally
        Output.Free;
      end;
      logFile('connected server ' + vFTP);
    except
      on E: Exception do
      begin
        logFile('Error: ' + E.Message);
        Result := True;
      end;
    end;
  end;

begin
  vDatabaseFolder := qReportWebDATABASE_NAME.AsString;
  logFileWeb('Proses antrian id ' + qReportWebAPI_REPORT_ID.AsString);

  try
    // ====== KLAIM JOB SECARA ATOMIK (menggantikan blok SELECT+cek lama) ======
    if not TryClaimReport(qReportWebAPI_REPORT_ID.AsString) then
    begin
      // sudah diklaim / sedang diproses service lain -> lewati item ini
      Exit;
    end;

    try
      qErp_rpt_sql_name.Close;
      qErp_rpt_sql_name.Open;
      while not qErp_rpt_sql_name.Eof do
      begin
        // (blok pembersihan query SQL_NAME, sama seperti kode asli)
        qErp_rpt_sql_name.Next;
      end;
    except
      on E: exception do
      begin
        // diredam, sama seperti kode asli
      end;
    end;

    logFile('Proses data report web id ' + qReportWebAPI_REPORT_ID.AsString);

    // Catatan: blok manual berikut DIHAPUS karena sudah dilakukan atomik
    // oleh TryClaimReport di atas:
    //   qReportWeb.Edit;
    //   qReportWebGENERATE_FLAG.AsString := '2';
    //   qReportWeb.Post;

    if LoginManage(qReportWebDATABASE_NAME.AsString, qReportWebHOSTNAME.AsString, qReportWebPORT.AsString) then
    begin
      try
        vHariDel := '30';
        Run_query('SELECT * FROM SETUP');
        if Query1.FindField('REPORT_HISTORY') <> nil then
          vHariDel := Query1.fieldbyname('REPORT_HISTORY').AsString;

        qReportWebDel.Close;
        qReportWebDel.MacroByName('KONDISI').Value := ' AND a.HOSTNAME = ' + QuotedStr(vHostNames) + ' AND DATEDIFF(SYSDATE(),CREATED_DATE) > ' + vHariDel;
        qReportWebDel.Open;
        while not qReportWebDel.Eof do
        begin
          logFile('hapus report ' + qReportWebDelOUTPUT_FILE.AsString);
          if FileExists(qReportWebDelOUTPUT_FILE.AsString) then
            DeleteFile(qReportWebDelOUTPUT_FILE.AsString);
          qReportWebDel.Next;
        end;

        qUser.Close;
        qUser.ParamByName('ERP_USER_ID').Value := qReportWebERP_USER_ID.AsString;
        qUser.Open;
        logFile('proses laporan pada user id ' + qReportWebERP_USER_ID.AsString);
        Group_id := qUser.fieldbyname('ERP_GROUP_ID').Value;
        qSetup.Close;
        qSetup.Open;
        if qSetup.Eof then
        begin
          defa_template := patch + 'template\';
        end
        else
        begin
          if (pos(':', qsetup.fieldbyname('Template_folder').asstring) > 0) or (pos('\\', qsetup.fieldbyname('Template_folder').asstring) > 0) then
            defa_template := patch + vDatabaseFolder + '\' + vTemplateFolder + '\'
          else
            defa_template := patch + vDatabaseFolder + '\' + qsetup.fieldbyname('Template_folder').asstring + '\';
        end;

        qErpDetail.Close;
        qErpDetail.Open;
        qErpDetail.Filtered := False;
        qErpDetail.Filter := '';

        // ====================================================================
        // BLOK PARSING JSON (System.JSON) - PENGGANTI TlkJSON
        // ====================================================================
        js := TJSONObject.ParseJSONValue(qReportWebJSON.AsString);
        try
          if (js = nil) or (not (js is TJSONArray)) then
            raise Exception.Create('Format JSON tidak valid (bukan array): ' + qReportWebJSON.AsString);

          jsArr := js as TJSONArray;

          for I := 0 to jsArr.Count - 1 do
          begin
            if not (jsArr.Items[I] is TJSONObject) then
              Continue;

            Obj := jsArr.Items[I] as TJSONObject;

            if not Obj.TryGetValue<string>('ERP_RPT_DETAIL_ID', vERP_RPT_DETAIL_ID) then
              Continue;

            if Obj.TryGetValue<TJSONArray>('WHERE', WHERE) then
            begin
              for J := 0 to WHERE.Count - 1 do
              begin
                if not (WHERE.Items[J] is TJSONObject) then
                  Continue;

                WhereObj := WHERE.Items[J] as TJSONObject;

                qErpDetail.Append;
                qErpDetailERP_DETAIL_ID.AsString := vERP_RPT_DETAIL_ID;
                qErpDetailWHERE_COLUMN.AsString := WhereObj.GetValue<string>('COLUMN', '');
                qErpDetailWHERE_VALUE.AsString := WhereObj.GetValue<string>('VALUE', '');
                qErpDetail.Post;

                logFile('isi data  qErpDetail erp_detail_id ' + vERP_RPT_DETAIL_ID + ' kolom ' + qErpDetailWHERE_COLUMN.AsString + ' value ' + qErpDetailWHERE_VALUE.AsString);
              end;
            end;
          end;
        finally
          js.Free; // membebaskan seluruh tree JSON (jsArr, Obj, WHERE, WhereObj ikut terbebas)
        end;
        // ====================================================================
        // AKHIR BLOK PARSING JSON
        // ====================================================================

        logFile('berhasil menambahkan filter ' + IntToStr(qErpDetail.RecordCount) + ' data');
        logFileWeb('Proses Query dan filter');
        vError := '';
        keluar;
        if pertamakali() then
        begin
          logFile('proses data');
          vErr := '';
          file_excel := '';
          if createExcel(qReportWebERP_RPT_ID.AsString) then
          begin
            logFile('proses upload');
            if FileExists(file_excel) and (file_excel <> '') then
            begin
              vFileFtp := '';
              try
                qReportWeb.Edit;
                qReportWebGENERATE_FLAG.AsString := '1';
                qReportWebOUTPUT_FILE.AsString := file_excel;

                if vFtpBol then
                begin
                  if connectFtp then
                  begin
                    if UploadSingleFile(file_excel) then
                    begin
                      qReportWebPATH_FILE.AsString := vFileFtp;
                      logFile('upload file success');
                    end
                    else
                      logFile('upload file gagal, tetap simpan lokal di ' + file_excel);
                    if UploadSingleFile(vFilePdf) then
                    begin
                      qReportWebPDF_FLAG.AsString := '1';
                      qReportWebPATH_FILE.AsString := vFileFtp;
                      logFile('upload file success');
                    end
                    else
                      logFile('upload file gagal, tetap simpan lokal di ' + file_excel);
                  end
                  else
                    logFile('connectFtp gagal, laporan tersimpan lokal di ' + file_excel);
                end
                else
                begin
                  logFile(vReportFolder + ExtractFileName(vFilePdf));
                  if FileExists(vReportFolder + ExtractFileName(vFilePdf)) then
                  begin
                    qReportWebPATH_FILE.AsString := vFTP + '/' + qReportWebDATABASE_NAME.AsString + '/Report/' + ExtractFileName(vFilePdf);
                    qReportWebPDF_FLAG.AsString := '1';
                  end
                  else
                    qReportWebPATH_FILE.AsString := vFTP + '/' + qReportWebDATABASE_NAME.AsString + '/Report/' + ExtractFileName(file_excel);
                  logFile('Mode lokal aktif, file disimpan di ' + qReportWebPATH_FILE.AsString);
                end;

                qReportWeb.Post;
              except
                on E: exception do
                begin
                  logFile(' qreport web err ' + E.Message);
                end;
              end;
              SendNotification(qReportWebAPI_REPORT_ID.AsString);
            end;
          end
          else
          begin
            qReportWeb.Edit;
            qReportWebERROR.AsString := vErr;
            qReportWebERROR_USER.AsString := ' Report bermasalah harap hubungi support!';
            SendNotification(qReportWebAPI_REPORT_ID.AsString);
            qReportWeb.Post;
          end;
        end
        else
        begin
          if vError <> '' then
          begin
            qReportWeb.Edit;
            qReportWebERROR.AsString := vError;
            qReportWebERROR_USER.AsString := ' Report bermasalah harap hubungi support!';
            qReportWeb.Post;
          end;
        end;
        keluar;
      except
        on E: exception do
        begin
          qReportWeb.Edit;
          qReportWebERROR.AsString := E.Message;
          qReportWebERROR_USER.AsString := ' Report bermasalah harap hubungi support!';
          qReportWeb.Post;
          logFile('report bermasalah error erp_rpt_id ' + qReportWebERP_RPT_ID.AsString + ' ' + E.Message);
          SendNotification(qReportWebAPI_REPORT_ID.AsString);
        end;
      end;
    end;
  except
    on E: exception do
    begin
      qReportWeb.Edit;
      qReportWebERROR.AsString := E.Message;
      qReportWebERROR_USER.AsString := ' Report bermasalah harap hubungi support!';
      qReportWeb.Post;
      logFile(qReportWebDATABASE_NAME.AsString + ' on ' + qReportWebHOSTNAME.AsString + ':' + qReportWebPORT.AsString + ' error ' + E.Message);
      SendNotification(qReportWebAPI_REPORT_ID.AsString);
    end;
  end;
end;

procedure TsvReport.SendNotification(param1: string);
begin

  try
    logFile('proses send notif ' + param1);
    RESTRequest1.Params[0].Value := param1;
    RESTRequest1.Execute;
    if RESTResponse1.StatusCode = 200 then
    begin
      logFile(' notif terkirim di api_report_id ' + param1 + ' content ' + sLineBreak + RESTResponse1.Content)
    end
    else
    begin
      logFile(' error code ' + IntToStr(RESTResponse1.StatusCode) + ' api_report_id ' + param1 + ' content ' + sLineBreak + RESTResponse1.Content);
    end;
  except
    on E: Exception do
    begin
      logFile(' error notif api_report_id ' + param1 + ' ' + E.Message)
    end;
  end;
end;

procedure TsvReport.CleanupOldLogs(vHariSimpan: Integer);
var
  SearchRec: TSearchRec;
  vFileDate: TDateTime;
  vFullPath: string;
  vSelisihHari: Integer;
begin
  logFile('Proses cleanup log lama (hapus yg lebih dari ' + IntToStr(vHariSimpan) + ' hari)');
  if FindFirst(patch + 'Log\log*.txt', faAnyFile, SearchRec) = 0 then
  begin
    try
      repeat
        if (SearchRec.Attr and faDirectory) = 0 then
        begin
          vFullPath := patch + SearchRec.Name;
          vFileDate := FileDateToDateTime(SearchRec.Time);
          vSelisihHari := Trunc(Now) - Trunc(vFileDate);

          // Hapus jika file sudah lebih tua dari vHariSimpan hari
          if vSelisihHari > vHariSimpan then
          begin
            try
              if DeleteFile(vFullPath) then
                logFile('Log dihapus (umur ' + IntToStr(vSelisihHari) + ' hari): ' + SearchRec.Name)
              else
                logFile('Gagal hapus log: ' + SearchRec.Name);
            except
              on E: Exception do
                logFile('Error hapus log ' + SearchRec.Name + ': ' + E.Message);
            end;
          end;
        end;
      until FindNext(SearchRec) <> 0;
    finally
      FindClose(SearchRec);
    end;
  end;
  if FindFirst(patch + 'Log\track*.txt', faAnyFile, SearchRec) = 0 then
  begin
    try
      repeat
        if (SearchRec.Attr and faDirectory) = 0 then
        begin
          vFullPath := patch + SearchRec.Name;
          vFileDate := FileDateToDateTime(SearchRec.Time);
          vSelisihHari := Trunc(Now) - Trunc(vFileDate);

          // Hapus jika file sudah lebih tua dari vHariSimpan hari
          if vSelisihHari > vHariSimpan then
          begin
            try
              if DeleteFile(vFullPath) then
                logFile('Log dihapus (umur ' + IntToStr(vSelisihHari) + ' hari): ' + SearchRec.Name)
              else
                logFile('Gagal hapus log: ' + SearchRec.Name);
            except
              on E: Exception do
                logFile('Error hapus log ' + SearchRec.Name + ': ' + E.Message);
            end;
          end;
        end;
      until FindNext(SearchRec) <> 0;
    finally
      FindClose(SearchRec);
    end;
  end;
end;

procedure TsvReport.mulai;
begin
  try
    qSetup.Close;
    qSetup.Open;
  except
    on E: exception do
    begin
      logFile('qSetup open error ' + E.Message);
      Exit;
    end;
  end;

  FArray := VarArrayCreate([0, maxvararray], varVariant);
  FArray[0] := qSetup.fieldbyname('NAME').AsString;
  FArray[1] := qSetup.fieldbyname('ADDRESS1').AsString;
  if qSetup.FieldByName('city').AsString <> '' then
    FArray[2] := qSetup.fieldbyname('city').AsString + ' ' + qSetup.fieldbyname('phone').AsString
  else
    FArray[2] := qSetup.fieldbyname('phone').AsString;
  if not Run_query('SELECT ERP_USER_NAME FROM ERP_USER WHERE ERP_USER_ID=' + qReportWebERP_USER_ID.AsString) then
  begin
    FArray[19] := Query1.Fields[0].AsString;
    FArray[20] := 'Dicetak oleh ' + Query1.Fields[0].AsString;
    Ar1[19] := Query1.Fields[0].AsString;
    Ar1[20] := 'Dicetak oleh ' + Query1.Fields[0].AsString;
  end;
  Ar1[0] := qSetup.fieldbyname('NAME').AsString;
  Ar1[1] := qSetup.fieldbyname('ADDRESS1').AsString;
  if qSetup.FieldByName('city').AsString <> '' then
    Ar1[2] := qSetup.fieldbyname('city').AsString + ' ' + qSetup.fieldbyname('phone').AsString
  else
    Ar1[2] := qSetup.fieldbyname('phone').AsString;
end;

function TsvReport.SetInitAwalModule: Boolean;
begin
  Result := False;
  Run_Query(' SELECT * FROM erp_lookup_value ' + ' where PROGRAM_CODE1 is not null ' + ' And PROGRAM_CODE1<>'''' and Program_Code1 not like ''% %'' ' + ' and Program_Code1 not like ''%-%''');
  while not Query1.Eof do
  begin
    if not Run_sql('SELECT erp_lookup_value_id INTO @' + Query1.Fieldbyname('PROGRAM_CODE1').AsString + ' FROM erp_lookup_value WHERE ' + ' erp_lookup_value_id = ' + quotedstr(Query1.Fieldbyname('erp_lookup_value_id').AsString) + ' limit 1') then
      Exit;
    Query1.Next;
  end;

  Run_Query('SELECT * FROM erp_lookup_set ' + ' where PROGRAM_CODE is not null ' + ' And PROGRAM_CODE<>'''' and Program_Code not like ''% %'' ' + ' and Program_Code not like ''%-%''');
  while not Query1.Eof do
  begin
    if not Run_sql('SELECT erp_lookup_set_id INTO @' + Query1.Fieldbyname('PROGRAM_CODE').AsString + ' FROM erp_lookup_set WHERE ' + ' erp_lookup_set_id = ' + quotedstr(Query1.Fieldbyname('erp_lookup_set_id').AsString) + ' limit 1') then
      Exit;
    Query1.Next;
  end;

  if branch_id = '' then
  begin
    if Run_query('SELECT BRANCH_ID, DATABASE_NAME FROM BRANCH WHERE DATABASE_NAME = (SELECT DATABASE()) AND ACTIVE_FLAG="Y" ORDER BY 1') then
    begin
      while not Query1.Eof do
      begin
        if Query1.Fields[1].AsString = db1.Database then
          branch_id := Query1.Fields[0].AsString;
        Query1.Next;
      end;

      if branch_id = '' then
        branch_id := Query1.Fields[0].AsString;
    end;
  end;

  if Run_Query('SELECT * FROM COA_SETUP GROUP BY COA_ID') then
  begin
    if Query1.FindField('BRANCH_ID') = nil then
    begin
      while not Query1.Eof do
      begin
        if not Run_sql('SELECT COA_ID INTO @' + Query1.Fieldbyname('PROGRAM_ACCOUNT').AsString + ' FROM coa_setup WHERE ' + ' PROGRAM_ACCOUNT = ' + quotedstr(Query1.Fieldbyname('PROGRAM_ACCOUNT').AsString) + ' limit 1') then
          Exit;
        Query1.Next;
      end;
    end
    else
    begin
      if branch_id <> '' then
      begin
        Run_Query('SELECT * FROM COA_SETUP WHERE BRANCH_ID = ' + branch_id + ' GROUP BY COA_ID');
      end
      else
      begin
        if Run_Query('SELECT * FROM COA_SETUP GROUP BY COA_ID') then
        begin
          while not Query1.Eof do
          begin
            if not Run_sql('SELECT COA_ID INTO @' + Query1.Fieldbyname('PROGRAM_ACCOUNT').AsString + ' FROM coa_setup WHERE ' + ' PROGRAM_ACCOUNT = ' + quotedstr(Query1.Fieldbyname('PROGRAM_ACCOUNT').AsString) + ' limit 1') then
              Exit;
            Query1.Next;
          end;
        end;
      end;
    end;
  end;

  NOT_BANK_all_flag := false;
  if qReportWebERP_USER_ID.AsString <> '' then
  begin
    if Run_query('select * from erp_user_d_akun where erp_user_id = ' + qReportWebERP_USER_ID.AsString) then
      if not query1.eof then
        NOT_BANK_all_flag := true;
  end;

  if (CBank.findMacro('kondisiUser') <> nil) then
  begin
    if NOT_BANK_all_flag then
    begin
      CBank.MacroByName('kondisiUser').Value := ' and  eud.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString + ' ';
    end
    else
    begin
      CBank.MacroByName('kondisiUser').Value := '  ';

    end;
  end;

  NOT_REPORT_all_flag := false;
  if Group_id <> '' then
  begin
    if Run_query('select * from  erp_group_rpt  where  erp_rpt_id is not null  and erp_group_id = ' + Group_id) then
      if not query1.eof then
        NOT_REPORT_all_flag := true;
  end;

  NOT_GUDANG_all_flag := false;
  if qReportWebERP_USER_ID.AsString <> '' then
  begin
    if Run_query('select * from erp_warehouse where erp_user_id = ' + qReportWebERP_USER_ID.AsString) then
      if not query1.eof then
        NOT_GUDANG_all_flag := true;
  end;

  try
    cgudang.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
    CGudang.close;
    if NOT_GUDANG_all_flag then
    begin
      CGudang.MacroByName('kondisiGUDANG').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString;
    end
    else
    begin
      CGudang.MacroByName('kondisiGUDANG').Value := '';
    end;
    CGudang.open;
  except
    on E: exception do
    begin
      logFile('open CGudang Bermasalah error ' + E.Message + '!');
    end;
  end;

  NOT_DIVISI_all_flag := false;
  if qReportWebERP_USER_ID.AsString <> '' then
  begin
    if Run_query('select * from erp_group_divisi where erp_user_id = ' + qReportWebERP_USER_ID.AsString) then
      if not query1.eof then
        NOT_DIVISI_all_flag := true;
  end;

  //nass 23 Juni 2022 tambahan buat report
  NOT_SALES_all_flag := false;
  if qReportWebERP_USER_ID.AsString <> '' then
  begin
    if Run_query('select * from erp_user_sales where erp_user_id = ' + qReportWebERP_USER_ID.AsString) then
      if not query1.eof then
        NOT_SALES_all_flag := true;
  end;

  try
    cSales.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
    cSales.Close;
    if NOT_SALES_all_flag then
    begin
      cSales.MacroByName('kondisiUser').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString;
    end
    else
    begin
      cSales.MacroByName('kondisiUser').Value := '';
    end;
    cSales.Open;
  except
    on E: exception do
    begin
      logFile('open cSales Bermasalah error ' + E.Message + '!');
    end;
  end;

  try
    TIPE.close;
    TIPE.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
    if NOT_DIVISI_all_flag then
    begin
      TIPE.MacroByName('kondisiUser').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString;
    end
    else
    begin
      TIPE.MacroByName('kondisiUser').Value := '';
    end;
    TIPE.open;
  except
    on E: exception do
    begin
      logFile('open TIPE Bermasalah error ' + E.Message + '!');
    end;
  end;

  NOT_MEREK_all_flag := false;
  if qReportWebERP_USER_ID.AsString <> '' then
  begin
    if Run_query('select * from erp_group_merek where erp_user_id = ' + qReportWebERP_USER_ID.AsString) then
      if not query1.eof then
        NOT_MEREK_all_flag := true;
  end;

  try
    MEREK.paramByName('USER').asstring := qReportWebERP_USER_ID.AsString;
    MEREK.close;
    if NOT_MEREK_all_flag then
    begin
      MEREK.MacroByName('kondisiUser').Value := ' and  g.ERP_USER_ID= ' + qReportWebERP_USER_ID.AsString;
    end
    else
    begin
      MEREK.MacroByName('kondisiUser').Value := '';
    end;
    MEREK.Open;
  except
    on E: exception do
    begin
      logFile('open MEREK Bermasalah error ' + E.Message + '!');
    end;
  end;

  Result := True;
end;

procedure TsvReport.CopyMyFile(SourceFile, DestFile: string);
begin
  try
    // The third parameter 'True' overwrites the file if it already exists.
    // Set it to 'False' if you want it to raise an exception instead.
    logFile('Copy file ' + SourceFile + ' ke ' + DestFile);
    TFile.Copy(SourceFile, DestFile, True);

    logFile('File copied successfully!');
  except
    on E: Exception do
      logFile('Error copying file: ' + E.Message);
  end;
end;

function TsvReport.DeleteFolderRecursive(const APath: string): Boolean;
var
  SearchRec: TSearchRec;
  vFullPath: string;
begin
  Result := True;
  if not DirectoryExists(APath) then
    Exit;

  if FindFirst(IncludeTrailingPathDelimiter(APath) + '*', faAnyFile, SearchRec) = 0 then
  begin
    try
      repeat
        if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
        begin
          vFullPath := IncludeTrailingPathDelimiter(APath) + SearchRec.Name;
          if (SearchRec.Attr and faDirectory) = faDirectory then
          begin
            // rekursif hapus subfolder dulu
            if not DeleteFolderRecursive(vFullPath) then
              Result := False;
          end
          else
          begin
            // hapus file, pastikan bukan read-only dulu
            try
              if (Integer(SearchRec.Attr) and Integer(faReadOnly)) = Integer(faReadOnly) then
                FileSetAttr(vFullPath, Integer(SearchRec.Attr) and not Integer(faReadOnly));
              if not DeleteFile(vFullPath) then
              begin
                logFile('Gagal hapus file: ' + vFullPath);
                Result := False;
              end;
            except
              on E: Exception do
              begin
                logFile('Error hapus file ' + vFullPath + ': ' + E.Message);
                Result := False;
              end;
            end;
          end;
        end;
      until FindNext(SearchRec) <> 0;
    finally
      FindClose(SearchRec);
    end;
  end;

end;

end.

