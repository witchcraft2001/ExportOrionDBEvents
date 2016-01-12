unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,  Vcl.FileCtrl,
  Vcl.Samples.Spin, DB, ADODB, Bde.DBTables,DateUtils, inifiles;

const
  //ConnectStr='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Extended Properties=Paradox 5.x';
  PrgName = 'ExportOrionDB';
  PrgFriendName = 'Экспорт данных из БД Орион';
  PrgVersion = 'v.0.1.110116';


type
  TMainForm = class(TForm)
    Label1: TLabel;
    textPath: TButtonedEdit;
    btnPath: TBitBtn;
    btnOutputPath: TBitBtn;
    textOutput: TButtonedEdit;
    Label3: TLabel;
    Label4: TLabel;
    textPeriod: TSpinEdit;
    Bevel1: TBevel;
    btnExport: TButton;
    btnExit: TButton;
    OrionDB: TDatabase;
    MQuery: TQuery;
    Label2: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    textSeparator: TButtonedEdit;
    radioLastRecords: TRadioButton;
    radioPeriod: TRadioButton;
    procedure btnExportClick(Sender: TObject);
    procedure OpenLog;
    procedure Log(level:integer;str:string);
    Procedure CloseLog;
    function GetTempDirectory: String;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnPathClick(Sender: TObject);
    procedure btnOutputPathClick(Sender: TObject);
    procedure textSeparatorChange(Sender: TObject);
    function Text2CSVField(Text: String):String;
    procedure textPathChange(Sender: TObject);
    procedure textOutputChange(Sender: TObject);
    procedure radioLastRecordsClick(Sender: TObject);
    procedure textPeriodChange(Sender: TObject);
  private
    { Private declarations }
    LogLevel:integer;
    logFile:textfile;
    LogFileName:string;
    Logging: boolean;
    AppDataPath: string;
  public
    { Public declarations }
    paramPathDB:string;
    paramOutput:string;
    paramPeriod:integer;
    paramSeparator:string;
    paramLastRecords:boolean;
    interactiveRun:boolean;
    strFilter:TstringList;
    strQueryName:string;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.btnExitClick(Sender: TObject);
begin
    MainForm.Close;
end;

function TMainForm.Text2CSVField(Text: String):String;
var
  i,code:integer;
  isCipher: boolean;
  res:string;
begin
  result:='';
  if Text.Length=0 then Exit;
  res:='';
  Val(Text,i,code);
  if code <> 0 then begin
    for i := 1 to Text.Length do begin
{      if not ((ord(Text[i]) >= ord('1')) and (ord(Text[i]) <= ord('0')))  then begin
        isCipher:=false;
      end;}
      if Text[i] = '"' then begin
        res:=res+'"';
  //      j:=j+1;
      end;
      res:=res + Text[i];
  //    j:=j+1;
    end;
    isCipher := false
  end else begin
    isCipher := true;
    res:=Text;
  end;
//  res[j]:=chr(0);
  if not isCipher then result:= '"' + res+ '"'
    else result:=res;
end;

procedure TMainForm.btnExportClick(Sender: TObject);
var
        //MQuery:TADOQuery;
//        strTimeFrom:string;
        strTime:string;
//        listFields:TStrings;
        I, J:integer;
        fileOut:TextFile;
        timeFrom: TDateTime;
        timeTo: TDateTime;
  X: Integer;
  str1,str2:string;
  flt:boolean;
  iniFile:TIniFile;
  intCount:Integer;
begin
        Log(1,'Подготовка к экспорту');
        if (strQueryName.Length = 0) then strQueryName := 'NoNamedQuery';
        //timeTo:=EncodeDateTime(2015,9,8,8,0,0,0);
        Log(2,'Подключение к БД: '+textPath.Text);
        try
          OrionDB.Params.Add('PATH=' + textPath.Text);
          OrionDB.Params.Add('DEFAULT DRIVER=PARADOX');
          OrionDB.Params.Add('ENABLE BCD=FALSE');
          OrionDB.Connected:=true;
        Except
          on E : Exception do begin
            Log(3,E.ClassName+' ошибка с сообщением : '+E.Message);
            Exit;
          end;
        end;
        Log(2,'Подключено!');
        Log(2,'Подготовка запроса...');
        try
          if paramLastRecords then
          begin
          //Подготовка запроса для выгрузки последних записей
            iniFile := TIniFile.Create(AppDataPath + '\settings.ini');
            intCount:=iniFile.ReadInteger('Counters',strQueryName,0);
            MQuery.SQL.Text := 'SELECT * FROM pLogData, pList WHERE pList.ID = HozOrgan AND Event = 32 AND pLogData.Num > :intCount ORDER BY TimeVal;';
            MQuery.Prepare;
            MQuery.Params[0].AsInteger := intCount;
          end else begin
          //Подготовка запроса для интервала времени
            timeTo:=Now;
            timeTo:=RecodeSecond(timeTo,0);
            timeFrom:=IncMinute(timeTo,textPeriod.Value*-1);
            MQuery.SQL.Text := 'SELECT * FROM pLogData, pList WHERE pList.ID = HozOrgan AND Event = 32 AND TimeVal >= :TimeFrom AND TimeVal <= :TimeTo ORDER BY TimeVal;';
            MQuery.Prepare;
            MQuery.Params[0].AsDateTime:=TimeFrom;
            MQuery.Params[1].AsDateTime:=TimeTo;
          end;
          MQuery.Active := true;
        Except
          on E : Exception do begin
            Log(3,E.ClassName+' ошибка с сообщением : '+E.Message);
            Exit;
          end;
        end;
        Log(2,'Запрос вернул ' + IntToStr(MQuery.RecordCount) + ' строк.');
        if MQuery.RecordCount>0 then begin

          flt:=false;
          //Проверяем, есть ли после фильтра хоть одна запись
          while not MQuery.Eof do begin
              if strFilter.Count>0 then begin
                for X := 0 to strFilter.Count-1 do begin
                  str1 := copy(strFilter[X],1,Pos('=',strFilter[X])-1);
                  str2 := copy(strFilter[X],Pos('=',strFilter[X])+1,Length(strFilter[X])-Pos('=',strFilter[X])+1);
                  if MQuery.FieldByName(str1).AsString = str2 then begin
                    flt:=true;
                    break;
                  end;
                end;
              end else flt := true;
              MQuery.Next;
          end;
          if flt=true then begin
            strTime:=FormatDateTime('ddmmyy_hhmmss', Now);
            Log(2,'Экспорт в файл ' + textOutput.Text + '\orion_' + strTime + '.csv');
            try

            AssignFile(fileOut, textOutput.Text + '\orion_' + strTime + '.csv');
            Rewrite(fileOut);
            Except
              on E : Exception do begin
                Log(3,E.ClassName+' ошибка с сообщением : '+E.Message);
                Exit;
              end;
            end;

            MQuery.First;
            //Записываем имена полей
            for I := 0 to MQuery.Fields.Count-1 do begin
                if I>0 then begin
                  Write(fileOut, paramSeparator);
                end;
                Write(fileOut,MQuery.Fields[I].DisplayName);
            end;
            WriteLn(fileOut);
            J:=0;
            while not MQuery.Eof do begin
                flt:=false;
              //for I := 0 to listFields.Count-1 do begin
                if strFilter.Count>0 then begin
                  for X := 0 to strFilter.Count-1 do begin
                    str1 := copy(strFilter[X],1,Pos('=',strFilter[X])-1);
                    str2 := copy(strFilter[X],Pos('=',strFilter[X])+1,Length(strFilter[X])-Pos('=',strFilter[X])+1);
                    if MQuery.FieldByName(str1).AsString = str2 then begin
                      flt:=true;
                      break;
                    end;
                  end;
                end else flt := true;
                if flt=true then begin

                  if paramLastRecords then
                  begin
                    //Последняя сохраненная запись в intCount
                    if intCount<MQuery.FieldByName('Num').AsInteger then intCount := MQuery.FieldByName('Num').AsInteger;
                  end;

                  for I := 0 to MQuery.Fields.Count-1 do begin
                      if I>0 then begin
                        Write(fileOut,paramSeparator);
                      end;
                      Write(fileOut,Text2CSVField(MQuery.Fields[I].AsString));
                  end;
                    //ShowMessage(MQuery.Fields[I].AsString);
                  WriteLn(fileOut);
                  J:=J+1;
                end;
                MQuery.Next;
            end;
            CloseFile(fileOut);
            if paramLastRecords then
            begin
              iniFile.WriteInteger('Counters',strQueryName,intCount);
              iniFile.Free;
            end;
            Log(2,'Выгружено в файл ' + IntToStr(J) + ' строк.');
          end;
        end;
        //ShowMessage('OK');
        OrionDB.Connected:=False;
        Log(2,'Операция завершена.');
{        MyConnection.ConnectionString:=Format(ConnectStr,[textPath.Text]);
        MyConnection.Connected := true;
        AssignFile(fileOut, textOutput.Text + '\Test.txt');
        Rewrite(fileOut);
        MQuery:=TADOQuery.Create(nil);
        MQuery.Connection:=MyConnection;
        strTimeFrom:='#08/08/2015 08:15:00#';
        strTimeTo:='#08/08/2015 08:20:00#';
        MQuery.SQL.Text:=StringReplace(textQuery.Text,'%timefrom%',strTimeFrom,[rfReplaceAll, rfIgnoreCase]);
        MQuery.SQL.Text:=StringReplace(MQuery.SQL.Text,'%timeto%',strTimeTo,[rfReplaceAll, rfIgnoreCase]);
        //ShowMessage(MQuery.SQL.Text);
        MQuery.Active:=true;
//        MQuery.Fields.GetFieldNames(listFields);
        while not MQuery.Eof do begin
          //for I := 0 to listFields.Count-1 do begin
          for I := 0 to MQuery.Fields.Count-1 do begin
            if I>0 then begin
              Write(fileOut,', ');
            end;
            Write(fileOut,'"',MQuery.Fields[I].AsString,'"');
            //ShowMessage(MQuery.Fields[I].AsString);
          end;
          WriteLn(fileOut);
          MQuery.Next;
        end;
        CloseFile(fileOut);
        MyConnection.Connected:=false;}
end;

procedure TMainForm.btnPathClick(Sender: TObject);
var
  dir:string;
begin
  dir:=textPath.Text;
  if SelectDirectory(dir, [sdPrompt], 1000) then textPath.Text:=dir;
//  'Выберите папку, где расположена БД Орион', textPath.Text, textPath.Text);
end;

procedure TMainForm.btnOutputPathClick(Sender: TObject);
var
  dir:string;
begin
  dir:=textOutput.Text;
//  if GetFolderDialog(Application.Handle, 'Выбрать папку на Delphi (Дельфи)', sFolder) then
  if SelectDirectory(dir, [sdAllowCreate, sdPerformCreate, sdPrompt], 1000) then textOutput.Text:=dir;
end;


function TMainForm.GetTempDirectory: String;
var
  tempFolder: array[0..MAX_PATH] of Char;
begin
  GetTempPath(MAX_PATH, @tempFolder);
  result := StrPas(tempFolder);
end;

procedure TMainForm.OpenLog();
var
  temp:string;
  i: Integer;
begin
        if not Logging then exit;
        if LogLevel=0 then LogLevel := 5;
        temp:=FormatDateTime('yymmdd', Now);
        LogFileName := GetTempDirectory + PrgName+'_'+temp+'.log';
//        showmessage(logfilename);
{        if FileExists(ExtractFilePath(Application.ExeName)+'dbimporter.log') then begin
                AssignFile(LogFile,ExtractFilePath(Application.ExeName)+'dbimporter.log');
                Append(LogFile);
        end else begin
                AssignFile(LogFile,ExtractFilePath(Application.ExeName)+'dbimporter.log');
                Rewrite(LogFile);
        end;          }
        try
          AssignFile(LogFile,LogFileName);
        i:=FileMode;
        FileMode:=fmOpenWrite;
        if FileExists(LogFileName) then Append(LogFile)
          else ReWrite(LogFile);
//        Rewrite(LogFile);
        fileMode:=i;

        Log(1,'#-------------------------------#');
        Log(1,Format('Приложение %s %s запущено',[PrgName,PrgVersion]));
        Except
          Logging:=false;
          Exit;
        end;
end;

procedure TMainForm.radioLastRecordsClick(Sender: TObject);
begin
  paramLastRecords:=radioLastRecords.Checked;
end;

procedure TMainForm.textOutputChange(Sender: TObject);
begin
  paramOutput:=textOutput.Text;
end;

procedure TMainForm.textPathChange(Sender: TObject);
begin
  paramPathDB:=textPath.Text;
end;

procedure TMainForm.textPeriodChange(Sender: TObject);
var
  i,code: integer;
begin
  Val(textPeriod.Text,i,code);
  if code = 0 then paramPeriod := i;
end;

procedure TMainForm.textSeparatorChange(Sender: TObject);
begin
  paramSeparator:=textSeparator.Text;
end;

procedure TMainForm.Log(level:integer;str:string);
begin
        if not Logging then exit;
        if LogLevel>=Level then begin
            try
                Writeln(LogFile,TimeToStr(Time)+' ('+inttostr(level)+'): ',str);
                Flush(LogFile);
            Except

            end;
        end;
end;

Procedure TMainForm.CloseLog;
begin
        if not Logging then exit;
        Log(1,'Приложение остановлено.');
        Flush(LogFile);  { ensures that the text was actually written to file }
        CloseFile(LogFile);
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
  iniFile:TIniFile;
begin
  if interactiveRun then
  begin
    iniFile:=TIniFile.Create(AppDataPath+'\settings.ini');
    iniFile.WriteString('Common','textPath', textPath.Text);
    iniFile.WriteString('Common','textOutput', textOutput.Text);
    iniFile.WriteString('Common','textPeriod',textPeriod.Text);
    iniFile.WriteString('Common','textSeparator',textSeparator.Text);
    iniFile.WriteBool('Common','boolWriteLastRecords',radioLastRecords.Checked);
    iniFile.Free;
    CloseLog();
    strFilter.Free;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  iniFile:TIniFile;
begin
    strFilter := TStringList.Create;
    paramPeriod:=0;
    paramPathDB:='';
    paramOutput:='';
    paramSeparator:=';';

    interactiveRun:=true;
    Logging:= True;
    OpenLog();
    AppDataPath := System.SysUtils.GetEnvironmentVariable('APPDATA')+'\'+PrgName;
    if not DirectoryExists(AppDataPath) then CreateDir(AppDataPath);
    if interactiveRun then
    begin
      iniFile:=TIniFile.Create(AppDataPath+'\settings.ini');
      paramPeriod:=iniFile.ReadInteger('Common','textPeriod',5);
      paramPathDB:=iniFile.ReadString('Common','textPath','');
      paramOutput:=iniFile.ReadString('Common','textOutput','');
      paramSeparator:=iniFile.ReadString('Common','textSeparator',';');
      paramLastRecords:=iniFile.ReadBool('Common','boolWriteLastRecords',true);
      textPath.Text := paramPathDB;
      textOutput.Text := paramOutput;
      textPeriod.Text := IntToStr(paramPeriod);
      textSeparator.Text := paramSeparator;
      if paramLastRecords then radioLastRecords.Checked := true
      else radioPeriod.Checked := true;
    end;

end;

procedure TMainForm.FormShow(Sender: TObject);
begin
    if interactiveRun then Log(1,'Приложение запущено в интерактивном режиме!');

{    if paramPathDB <> '' then begin
        textPath.Text:=paramPathDB;
//        Log(1,'PathToDB = ' + paramPathDB);
    end;
    if paramOutput <> '' then begin
        textOutput.Text:=paramOutput;
//        Log(1,'Output = ' + paramOutput);
    end;
    if paramPeriod <> 0 then begin
        textPeriod.Value:=paramPeriod;
//        Log(1,'Period = ' + IntToStr(paramPeriod));
    end;}
end;

end.
