program serviceReportGenerate;

uses
  Vcl.SvcMgr,
  System.SysUtils,
  uSvReport in 'uSvReport.pas' {Service1: TService};

{$R *.RES}

const
  INSTANCE_SWITCH = '-instance=';

function GetInstanceName: string;
var
  index: integer;
begin
  Result := '';
  for index := 1 to ParamCount do
  begin
    if SameText(INSTANCE_SWITCH, Copy(ParamStr(index), 1, Length(INSTANCE_SWITCH))) then
    begin
      Result := Copy(ParamStr(index), Length(INSTANCE_SWITCH) + 1, MaxInt);
      Break;
    end;
  end;
  if (Result <> '') and (Result[1] = '"') then
    Result := AnsiDequotedStr(Result, '"');
end;

var
  inst: string;

begin
  // Windows 2003 Server requires StartServiceCtrlDispatcher to be
  // called before CoRegisterClassObject, which can be called indirectly
  // by Application.Initialize. TServiceApplication.DelayInitialize allows
  // Application.Initialize to be called from TService.Main (after
  // StartServiceCtrlDispatcher has been called).
  //
  // Delayed initialization of the Application object may affect
  // events which then occur prior to initialization, such as
  // TService.OnCreate. It is only recommended if the ServiceApplication
  // registers a class object with OLE and is intended for use with
  // Windows 2003 Server.
  //
  // Application.DelayInitialize := True;
  //

  inst := GetInstanceName;

  if not Application.DelayInitialize or Application.Installing then
    Application.Initialize;

  Application.CreateForm(TsvReport, svReport);

  // Jika ingin service mendukung multi-instance,
  // gunakan nilai 'inst' untuk mengubah nama service, contoh:
  // if inst <> '' then
  //   Service1.Name := Service1.Name + '_' + inst;

  if (inst <> '') then
  begin
    svReport.InstanceName := inst;
  end;
  Application.Run;
end.
