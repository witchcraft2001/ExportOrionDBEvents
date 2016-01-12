program EODB;

uses
  Vcl.Forms, SysUtils,
  Main in 'Main.pas' {MainForm};

{$R *.res}

var
  i: integer;

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  i := 1;
  if ParamCount > 1 then
  begin

    try
      while i <= ParamCount do
      begin
        if UpperCase(ParamStr(i)) = '/PATHDB' then
        begin
          MainForm.paramPathDB := ParamStr(i + 1);
          MainForm.textPath.Text := ParamStr(i + 1);
          MainForm.Log(1, 'Param PathToDB = ' + ParamStr(i + 1));
          i := i + 2;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/OUTPUT' then
        begin
          MainForm.paramOutput := ParamStr(i + 1);
          MainForm.textOutput.Text := ParamStr(i + 1);
          MainForm.Log(1, 'Param Output = ' + ParamStr(i + 1));
          i := i + 2;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/PERIOD' then
        begin
          MainForm.paramPeriod := StrToInt(ParamStr(i + 1));
          MainForm.textPeriod.Value := StrToInt(ParamStr(i + 1));
          MainForm.Log(1, 'Param Period = ' + ParamStr(i + 1));
          i := i + 2;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/FILTER' then
        begin
          MainForm.strFilter.Add(ParamStr(i + 1));
          MainForm.strFilter.LoadFromFile(ParamStr(i + 1));
          MainForm.Log(1, 'Param FilterFromFile = ' + ParamStr(i + 1));
          i := i + 2;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/LASTREC' then
        begin
          MainForm.paramLastRecords := True;
          MainForm.radioLastRecords.Checked := True;
          MainForm.Log(1, 'Param LastRecords = True');
          i := i + 1;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/QNAME' then
        begin
          MainForm.strQueryName := ParamStr(i + 1);
//          MainForm.textOutput.Text := ParamStr(i + 1);
          MainForm.Log(1, 'Param QueryName = ' + ParamStr(i + 1));
          i := i + 2;
          continue;
        end;

        if UpperCase(ParamStr(i)) = '/EXPORT' then
        begin
          MainForm.interactiveRun := false;
          MainForm.Log(1, 'Param CMD "Go Export"');
          MainForm.btnExportClick(nil);
          MainForm.Log(1, 'Export CMD ended.');
          i := i + 1;
          continue;
        end;
        i := i + 1;
      end;
    Except
      on E : Exception do MainForm.Log(3,E.ClassName+' ошибка с сообщением : '+E.Message);
    end;
  end
  else
    Application.Run;

end.
