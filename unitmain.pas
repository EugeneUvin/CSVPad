unit unitmain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, PrintersDlgs, Forms, Controls, Graphics,
  Dialogs,
  Grids, ComCtrls, Menus,
  Printers,
  StdCtrls,
  LConvEncoding,
  LazHelpHTML,
  UTF8Process,
  LCLIntf, LCLType,
  clipbrd, //clipboard
  fpspreadsheet, fpsallformats, //export ods /xls
  inifiles,csvdocument;

type
   TRange = record
     Min : integer;
     Max  : integer;
   end;

type

  { TFormMain }

  TFormMain = class(TForm)
    FindDialog1: TFindDialog;
    FontDialog1: TFontDialog;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem100: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItemAutoResize: TMenuItem;
    MenuItemToggleLabel: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItemColors: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItemRestoreDefaults: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    MenuItem52: TMenuItem;
    MenuItem53: TMenuItem;
    MenuItem54: TMenuItem;
    MenuItem55: TMenuItem;
    MenuItem56: TMenuItem;
    MenuItem57: TMenuItem;
    MenuItem58: TMenuItem;
    MenuItem59: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem60: TMenuItem;
    MenuItem61: TMenuItem;
    MenuItem62: TMenuItem;
    MenuItem63: TMenuItem;
    MenuItem64: TMenuItem;
    MenuItem65: TMenuItem;
    MenuItem66: TMenuItem;
    MenuItem67: TMenuItem;
    MenuItem68: TMenuItem;
    MenuItem69: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem70: TMenuItem;
    MenuItem71: TMenuItem;
    MenuItem72: TMenuItem;
    MenuItem73: TMenuItem;
    MenuItem74: TMenuItem;
    MenuItem75: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem80: TMenuItem;
    MenuItem81: TMenuItem;
    MenuItem82: TMenuItem;
    MenuItem83: TMenuItem;
    MenuItem84: TMenuItem;
    MenuItem85: TMenuItem;
    MenuItem86: TMenuItem;
    MenuItem87: TMenuItem;
    MenuItem88: TMenuItem;
    MenuItem89: TMenuItem;
    MenuItem9: TMenuItem;
    MenuItem90: TMenuItem;
    MenuItem91: TMenuItem;
    MenuItem92: TMenuItem;
    MenuItem93: TMenuItem;
    MenuItem94: TMenuItem;
    MenuItem95: TMenuItem;
    MenuItem96: TMenuItem;
    MenuItem97: TMenuItem;
    MenuItem98: TMenuItem;
    MenuItem99: TMenuItem;
    PopupMenu1: TPopupMenu;
    PrintDialog1: TPrintDialog;
    MenuItemSaveOnExit: TMenuItem;
    StatusBar1: TStatusBar;
    StringGrid1: TStringGrid;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    ToolButton2: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    ToolButton27: TToolButton;
    ToolButton28: TToolButton;
    ToolButton29: TToolButton;
    ToolButton3: TToolButton;
    ToolButton30: TToolButton;
    ToolButton31: TToolButton;
    ToolButton32: TToolButton;
    ToolButton33: TToolButton;
    ToolButton34: TToolButton;
    ToolButton35: TToolButton;
    ToolButton36: TToolButton;
    ToolButton37: TToolButton;
    ToolButton38: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;


    procedure MenuItemSearchOnlineClick(Sender: TObject);
    procedure MenuItemFontClick(Sender: TObject);
    procedure MenuItemDonateClick(Sender: TObject);
    procedure FindDialog1Find(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of String);
    procedure FormShow(Sender: TObject);
    procedure MenuItemOpenCSVClick(Sender: TObject);
    procedure MenuItemSaveClick(Sender: TObject);
    procedure MenuItemSaveAsClick(Sender: TObject);
    procedure MenuItemExportClick(Sender: TObject);
    procedure MenuItemQuitClick(Sender: TObject);
    procedure MenuItemAutoresizeClick(Sender: TObject);
    procedure MenuItemToggleLabelClick(Sender: TObject);
    procedure MenuItemMoveRightToLeftColClick(Sender: TObject);
    procedure MenuItemMoveLeftToRightColClick(Sender: TObject);
    procedure MenuItemMoveEndColClick(Sender: TObject);
    procedure MenuItemMoveRightColClick(Sender: TObject);
    procedure MenuItemMoveLeftColClick(Sender: TObject);
    procedure MenuItemMoveStartColClick(Sender: TObject);
    procedure MenuItemBottomToTopRowClick(Sender: TObject);
    procedure MenuItemTopToBottomRowClick(Sender: TObject);
    procedure MenuItemMoveBottomRowClick(Sender: TObject);
    procedure MenuItemMoveDownRowClick(Sender: TObject);
    procedure MenuItemMoveUpRowClick(Sender: TObject);
    procedure MenuItemMoveTopRowClick(Sender: TObject);
    procedure MenuItemDuplicateColClick(Sender: TObject);
    procedure MenuItemSwapRowsClick(Sender: TObject);
    procedure MenuItemAddColClick(Sender: TObject);
    procedure MenuItemRemoveColsClick(Sender: TObject);
    procedure MenuItemAddRowClick(Sender: TObject);
    procedure MenuItemRemoveRowsClick(Sender: TObject);
    procedure MenuItemToggleToolBarClick(Sender: TObject);
    procedure MenuItemWebpageClick(Sender: TObject);
    procedure MenuItemSwapColumnsClick(Sender: TObject);
    procedure MenuItemPrintClick(Sender: TObject);
    procedure MenuItemTakeSnapshotClick(Sender: TObject);
    procedure MenuItemCopyCellClick(Sender: TObject);
    procedure MenuItemPasteCellClick(Sender: TObject);
    procedure MenuItemSearchClick(Sender: TObject);
    procedure MenuItemDuplicateRowClick(Sender: TObject);
    procedure MenuItemColorLine1Click(Sender: TObject);
    procedure MenuItemColorLine2Click(Sender: TObject);
    procedure MenuItemRestoreDefaultsClick(Sender: TObject);
    procedure MenuItemColorLabelClick(Sender: TObject);
    procedure MenuItemSortColumnClick(Sender: TObject);
    procedure MenuItemDonate2Click(Sender: TObject);
    procedure MenuItemSaveOnExitClick(Sender: TObject);
    procedure MenuItemToggleStatusBarClick(Sender: TObject);
    procedure MenuItemAboutClick(Sender: TObject);
    procedure MenuItemNewClick(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; aCol, aRow: Integer; aRect: TRect; aState: TGridDrawState);
    procedure ToolButtonSortDecreaseClick(Sender: TObject);

   private
     { private declarations }
     procedure LoadStringGrid(csv:string);
     procedure Sg2Csv(sep:char;ff:string);
     procedure Sg2Xml(F:string);
     procedure PrintGrid(var sGrid : TStringGrid);
     procedure Sg2Ods(FilenameODS:string);
     procedure Sg2Xls(FilenameXLS:string);
   public
     { public declarations }
   end;




var
  FormMain: TFormMain;
  tlista   : TStringList;
  FileName : string;

  //Default values :
  sep : char = ',';
  line1    : Tcolor = $00eeeeee;
  line2    : Tcolor = clwhite;
  labelc    : Tcolor = clBtnFace;
  sgfontcolor   : Tcolor = clblack;
  sgfontsize : integer = 10;
  sgfontname : TFontName = 'Arial';
  sgfontstyle : TFontStyles = []; //fsBold,fsItalic,fsUnderline,fsStrikeOut


  eid,sid,lid : integer;


  filterAutoRecognition : string = 'Automatic recognition';
  filterXml : string = 'Xml format';
  filterHtml : string = 'Html documents';
  filterOds : string = 'Open document spreadsheet';
  filterCSV : string = 'Comma-separated file';
  filterCSVsemicolon : string = 'Semicolon-separated file';
  filterCSVpipe : string = 'Pipe-separated file';
  filterCSVasterisk : string = 'Asterisk-separated file';
  filterCSVcolon : string = 'Colon-separated file';
  filterCSVdollar : string = 'Dollar sign-separated file';
  filterCSVtab : string = 'Tab-separated file';
  filterXls : string = 'Excel format 8.0';

  s_true :boolean = false;


const
 conf = 'settings.ini';
 version  = '1.0';

implementation

{$R *.lfm}

//By TrustFm
procedure TFormMain.Sg2Ods(FilenameODS:string);
const OUTPUT_FORMAT = sfOpenDocument;
var i,a : integer;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
begin

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet');

  for i:=0 to stringgrid1.RowCount -1 do begin
      for a:=0 to stringgrid1.ColCount -1 do begin
             MyWorksheet.WriteUTF8Text(i,a,stringgrid1.Cells[a,i]);
      end;
  end;

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(FilenameODS, OUTPUT_FORMAT);
  MyWorkbook.Free;

end;


//By TrustFm
procedure TFormMain.Sg2Xls(FilenameXLS:string);
const OUTPUT_FORMAT = sfExcel8;
var i,a : integer;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
begin

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet');

  for i:=0 to stringgrid1.RowCount -1 do begin
      for a:=0 to stringgrid1.ColCount -1 do begin
             MyWorksheet.WriteUTF8Text(i,a,stringgrid1.Cells[a,i]);
      end;
  end;

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(FilenameXLS, OUTPUT_FORMAT);
  MyWorkbook.Free;

end;

//By TrustFm
procedure GoToURL(URL:string);
var
  v: THTMLBrowserHelpViewer;
  p: LongInt;
  BrowserProcess: TProcessUTF8;
  BrowserPath, BrowserParams : string;
begin
  v:=THTMLBrowserHelpViewer.Create(nil);
  try
    v.FindDefaultBrowser(BrowserPath,BrowserParams);
    p:=System.Pos('%s', BrowserParams);
    System.Delete(BrowserParams,p,2);
    System.Insert(URL,BrowserParams,p);
    BrowserProcess:=TProcessUTF8.Create(nil);
    try
      BrowserProcess.CommandLine:=BrowserPath+' '+BrowserParams;
      BrowserProcess.Execute;
    finally
      BrowserProcess.Free;
    end;
  finally
    v.Free;
  end;

end;

//main functions
//http://www.swissdelphicenter.ch/torry/showcode.php?id=2149
function FontStyletoStr(St: TFontStyles): string;
var
  S: string;
begin
  S := '';
  if St = [fsbold] then S := 'Bold'
  else if St = [fsItalic] then S := 'Italic'
  else if St = [fsStrikeOut] then S := 'StrikeOut'
  else if St = [fsUnderline] then S := 'UnderLine'

  else if St = [fsbold, fsItalic] then S := 'BoldItalic'
  else if St = [fsBold, fsStrikeOut] then S := 'BoldStrike'
  else if St = [fsBold, fsUnderline] then S := 'BoldUnderLine'
  else if St = [fsBold..fsStrikeOut] then S := 'BoldItalicStrike'
  else if St = [fsBold..fsUnderLine] then S := 'BoldItalicUnderLine'
  else if St = [fsbold..fsItalic, fsStrikeOut] then S := 'BoldItalicStrike'
  else if St = [fsBold, fsStrikeOut..fsUnderline] then S := 'BoldStrikeUnderLine'

  else if St = [fsItalic, fsStrikeOut] then S := 'ItalicStrike'
  else if St = [fsItalic..fsUnderline] then S := 'ItalicUnderLine'
  else if St = [fsStrikeOut..fsUnderLine] then S := 'StrikeUnderLine'
  else if St = [fsItalic..fsStrikeOut] then S := 'ItalicUnderLineStrike';
  FontStyletoStr := S;
end;
(*----------------------------------------------------*)

function StrtoFontStyle(St: string): TFontStyles;
var
  S: TfontStyles;
begin
  S  := [];
  St := UpperCase(St);
  if St = 'BOLD' then S := [fsBold]
  else if St = 'ITALIC' then S := [fsItalic]
  else if St = 'STRIKEOUT' then S := [fsStrikeOut]
  else if St = 'UNDERLINE' then S := [fsUnderLine]

  else if St = 'BOLDITALIC' then S := [fsbold, fsItalic]
  else if St = 'BOLDSTRIKE' then S := [fsBold, fsStrikeOut]
  else if St = 'BOLDUNDERLINE' then S := [fsBold, fsUnderLine]
  else if St = 'BOLDITALICSTRIKE' then S := [fsBold..fsStrikeOut]
  else if St = 'BOLDITALICUNDERLINE' then S := [fsBold..fsUnderLine]
  else if St = 'BOLDITALICSTRIKE' then S := [fsbold..fsItalic, fsStrikeOut]
  else if St = 'BOLDSTRIKEUNDERLINE' then S := [fsBold, fsStrikeOut..fsUnderline]

  else if St = 'ITALICSTRIKE' then S := [fsItalic, fsStrikeOut]
  else if St = 'ITALICUNDERLINE' then S := [fsItalic..fsUnderline]
  else if St = 'STRIKEUNDERLINE' then S := [fsStrikeOut..fsUnderLine]
  else if St = 'ITALICUNDERLINESTRIKE' then S := [fsItalic..fsStrikeOut];

  StrtoFontStyle := S;
end;

//By TrustFm
function GetSelectedRange(IsColumn:boolean; Grid : TStringGrid):TRange;
var res : TRange;
begin
     if IsColumn=false then begin
        res.Min := grid.Selection.Top;
        res.Max := grid.Selection.Bottom;
     end else begin //is column
         res.Min := grid.Selection.Left;
         res.Max := grid.Selection.Right;
     end;
     result := res;
end;


//By TrustFm
procedure Swap(IsColumn:boolean; Grid : TStringGrid);
var res : TRange;
begin
     res := GetSelectedRange(IsColumn, Grid);
     if IsColumn=true then begin
        if res.Max-res.Min>=1 then begin
           Grid.MoveColRow(IsColumn,res.Max, res.Min);
           Grid.MoveColRow(IsColumn,res.Min+1, res.Max);
        end else begin
            showmessage('Select a range of columns')
        end;
     end else begin //is row
         if res.Max-res.Min>=1 then begin
            Grid.MoveColRow(IsColumn,res.Max, res.Min);
            Grid.MoveColRow(IsColumn,res.Min+1, res.Max);
         end else begin
             showmessage('Select a range of rows')
         end;
     end;
end;


//By TrustFm
procedure Duplicate(IsColumn:boolean; Grid : TStringGrid);
var i : integer;
begin

     if IsColumn then begin
        Grid.InsertColRow(IsColumn,Grid.Col+1); //add an empty column
        for i:=0 to Grid.RowCount-1 do begin
            Grid.Cells[Grid.Col+1,i]:=Grid.Cells[Grid.Col,i];
        end;
     end else begin //is row
         Grid.InsertColRow(IsColumn,Grid.Row+1); //add an empty row
         for i:=0 to Grid.ColCount-1 do begin
             Grid.Cells[i,Grid.Row+1]:=Grid.Cells[i,Grid.Row];
         end;
     end;


end;



function elvalaszto(const elsosor: string): char; // auto separator
var
   sepa : array [0..6] of char;
   tmpmax,max,i:integer;
   c:char;
begin
     max := 0;
     tmpmax := 0;
     sepa[0] := ',';
     sepa[1] := ';';
     sepa[2] := '*';
     sepa[3] := '|';
     sepa[4] := ':';
     sepa[5] := '$';
     sepa[6] := #9;

     for i:=0 to high(sepa) do begin

         tmpmax := Length(elsosor) - Length(StringReplace(elsosor, sepa[i], '', [rfReplaceAll, rfIgnoreCase]));

         if max < tmpmax then begin
            max := tmpmax;
            c := sepa[i];
         end;

     end;

     Result := c;
end;

//Modded by TrustFm
procedure Sortgrid(Increment:boolean; Grid : TStringGrid; SortCol:integer);
{A simple exchange sort of grid rows}
var
   i,j : integer;
   temp:tstringlist;
begin
  temp:=tstringlist.create;
  with Grid do
       if Increment = true then begin
          for i := FixedRows to RowCount - 2 do  {because last row has no next row}
              for j:= i+1 to rowcount-1 do {from next row to end}
                  if CompareText(Cells[SortCol, i], Cells[SortCol,j]) > 0 then begin
                       temp.assign(rows[j]);
                       rows[j].assign(rows[i]);
                       rows[i].assign(temp);
                  end;
       end else begin //decremental

           for i := FixedRows to RowCount - 2 do  {because last row has no next row}
               for j:= i+1 to rowcount-1 do {from next row to end}
                   if CompareText(Cells[SortCol, i], Cells[SortCol,j]) < 0 then begin
                        temp.assign(rows[j]);
                        rows[j].assign(rows[i]);
                        rows[i].assign(temp);
                   end;
       end; //for
       temp.free;
end;



function htmlcolor(szin: TColor): string;
var
a,b,tmp :string;
begin

  tmp := ColorToString(szin);

    if tmp[1] = '$' then begin

      delete(tmp,1,3);
      a := tmp[1] + tmp[2];
      b := tmp[5] + tmp[6];
      delete(tmp,1,2);
      delete(tmp,3,4);
      tmp := '#'+b+tmp+a;

    end;

    if tmp[1]+tmp[2] = 'cl' then begin

      tmp := stringreplace(tmp,'cl','',[rfReplaceAll]);

    end;
  result := strlower(pchar(tmp));
end;

procedure SGridToHtml(SG: TStringGrid; filenamex:string);
var
  i, p: integer;
   Text,wh: string;
   Dest :tstringlist;

begin
  Dest := tstringlist.Create;
  Dest.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">');
  Dest.Add('<html>');
  Dest.Add('<head>');
  Dest.Add('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">');
  Dest.Add('<title>CSVpad v'+version+' export: ' + ExtractFileName(FileName) + '</title>');
  Dest.Add('<style type="text/css">');
  Dest.Add('body,table{font-family: Verdana, Arial, Helvetica, sans-serif; font-size: ' + inttostr(FormMain.StringGrid1.Font.Size) + 'pt; color: '+htmlcolor(FormMain.StringGrid1.Font.Color)+';}');
  Dest.Add('a,a:hover {text-decoration: underline; color: '+htmlcolor(sgfontcolor)+';}');
  Dest.Add('.center {margin: auto;text-align: left; background: '+htmlcolor(FormMain.Color)+'; border: 1px solid #eee;}');
  Dest.Add('.bg1 {background-color: '+htmlcolor(FormMain.stringgrid1.color)+';}');
  Dest.Add('.bg1:hover {background-color: #eee;}');
  Dest.Add('.bg2 {background: '+htmlcolor(FormMain.stringgrid1.AlternateColor)+';}');
  Dest.Add('.bg2:hover {background-color: #eee;}');
  Dest.Add('.fix {background-color: '+htmlcolor(FormMain.stringgrid1.FixedColor)+'}');
  Dest.Add('</style>');
  Dest.Add('</head>');
  Dest.Add('<body>');
  Dest.Add('<div style="text-align: center; font-size:'+inttostr(FormMain.StringGrid1.Font.Size + 1) +'pt;"><strong>' + ExtractFileName(FileName) + '</strong></div>');
  Dest.Add('<div style="text-align: center;">');
  Dest.Add(' <table class="center" cellpadding="1" cellspacing="1">');

  for i := 0 to SG.RowCount - 1 do
  begin

    if odd(i) then begin
       Dest.Add('  <tr class="bg2">');
    end else begin

    if (sg.FixedRows = 1) and (i = 0) then begin
        Dest.Add('  <tr class="fix">');

      end else begin
        Dest.Add('  <tr class="bg1">');
      end
    end;

    for p := 0 to SG.ColCount - 1 do
    begin

      Text := sg.Cells[p, i];
      if Text = '' then wh := '' else wh := ' width: '+inttostr(sg.ColWidths[p]+50)+'px;';
      Dest.Add('   <td style="height: 15px;'+wh+'">'+Text+'</td>');
    end;
    Dest.Add('  </tr>');

  end;
  Dest.Add('  </table>');
  Dest.Add('<br>');
  Dest.Add('<br>');
  Dest.Add('<div style="text-align: center;">Created using <strong>CSVpad v' + version + '</strong> by: <a href="http://www.trustfm.net/">TrustFm</a>. Based on DMcsvEditor by <a href="http://darhmedia.blogspot.hu/">Darh Media - Tivadar</a></div>');
  Dest.Add('</div>');
  Dest.Add('</body>');
  Dest.Add('</html>');
  Dest.Text := (Dest.Text); //UTF8encode not needed
  Dest.SaveToFile(filenamex);
  Dest.Free;
end;



procedure TFormMain.Sg2Xml(F:string);
var
xml:tstringlist;
i,a:integer;
begin
xml := tstringlist.Create;
{
    An Attribute is something that is self-contained, i.e., a color, an ID, a name.
    An Element is something that does or could have attributes of its own or contain other elements.

}
xml.Add('<?xml version="1.0" encoding="utf-8" standalone="yes"?>');
xml.Add('<!-- Creator: DMcsvEditor v'+version+' (linux) '+datetostr(now)+' -->');
xml.Add('<grid cols="'+inttostr(stringgrid1.ColCount)+'" rows="'+inttostr(stringgrid1.RowCount)+'">');

for i:=0 to stringgrid1.ColCount - 1 do begin

  for a:=0 to stringgrid1.RowCount - 1 do begin

      if stringgrid1.Cells[i,a] <> '' then
         xml.Add(' <cell col="'+inttostr(i+1)+'" row="'+inttostr(a+1)+'">'+stringgrid1.Cells[i,a]+'</cell>')
      else
          xml.Add(' <cell col="'+inttostr(i+1)+'" row="'+inttostr(a+1)+'" />');

  end;

end;

xml.Add('</grid>');
xml.Text := (xml.Text); //not needed UTF8Encode
xml.SaveToFile(F);
xml.Free;
end;

// CSV fájl feldogozása
procedure TFormMain.LoadStringGrid(csv:string);
var
i,a : integer;
csv1 : TCSVDocument;
begin

csv1 := TCSVDocument.Create;
csv1.Delimiter:= sep;
csv1.LoadFromFile(UTF8ToANSI(csv));


stringgrid1.BeginUpdate;
    try
    //stringgrid1.Font.CharSet:= 4;
    stringgrid1.RowCount := csv1.RowCount;
    stringgrid1.ColCount := csv1.ColCount[0];
   for i:=0 to csv1.RowCount -1 do begin

         for a:=0 to csv1.ColCount[i]-1 do begin

                 stringgrid1.Cells[a,i] := (StringReplace(csv1.Cells[a,i],#13#10,' ',[rfReplaceAll])); //sysToutf8 not needed

         end;

   end;

    finally
      stringgrid1.EndUpdate;
      FreeAndNil(csv1);
    end;

end;


//Print the grid
procedure TFormMain.PrintGrid(var sGrid : TStringGrid);
var
  r,c,x,y: Integer;

begin
     Printer.Title := ExtractFileName(FileName);
     printer.BeginDoc;
     Printer.Canvas.Font.Name  := stringgrid1.Font.Name;
     Printer.Canvas.Font.Size  := stringgrid1.Font.Size;

     x:= 10;
     //y1 := 50;
     for r:=0 to sgrid.RowCount - 1 do begin

         if printer.PageHeight <= (x+250) then begin
            printer.NewPage;
            x := 10;
         end;

         if  r = 0 then begin
             if sgrid.FixedRows > 0 then  begin
             Printer.Canvas.Font.Style := [fsBold];

             end;
         end else  begin
             Printer.Canvas.Font.Style := [];
         end;

        y := 20;
        for c := 0 to sgrid.ColCount - 1 do begin
            printer.Canvas.TextOut(y,x,sGrid.Cells[c,r]);
            y :=  sgrid.ColWidths[c] * 5+y;
        end;
     x := x +80;
     end;

  Printer.EndDoc;
end;


procedure TFormMain.Sg2Csv(sep:char;ff:string);
var i,a : integer;
    csv2 : TCSVDocument;
begin

     csv2 := TCSVDocument.Create;
     csv2.Delimiter:= sep;
     csv2.QuoteChar := '"';

     for i:=0 to stringgrid1.RowCount -1 do begin

         for a:=0 to stringgrid1.ColCount -1 do begin

             csv2.Cells[a,i] := stringgrid1.Cells[a,i];

         end;

     end;

     csv2.SaveToFile(ff);
     FreeAndNil(csv2);
end;


///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//EVENTS

//on form create
procedure TFormMain.FormCreate(Sender: TObject);
var folder : string;
    inif : tinifile;
begin
     tlista := tstringlist.Create;

     //restore form
     folder := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)));
     inif := tinifile.Create(folder +conf);
     try
     MenuItemSaveOnExit.Checked := inif.ReadBool('Main','SettingsSave',MenuItemSaveOnExit.Checked);
     lid := inif.ReadInteger('Main','LID',lid);
     sid := inif.ReadInteger('Main','SID',sid);
     eid := inif.ReadInteger('Main','EID',eid);

     MenuItemAutoResize.Checked := inif.ReadBool('Settings','AutoResize',MenuItemAutoResize.Checked);
     MenuItemToggleLabel.Checked := inif.ReadBool('Settings','Label',MenuItemToggleLabel.Checked);

     if MenuItemAutoResize.Checked then
        MenuItemAutoresizeClick(Sender);
     if MenuItemToggleLabel.Checked then
        MenuItemToggleLabelClick(Sender);


     MenuItem24.Checked := inif.ReadBool('View','ToolBar',MenuItem24.Checked);
     toolbar1.Visible := MenuItem24.Checked;
     MenuItem4.Checked := inif.ReadBool('View','StatusBar',MenuItem4.Checked);
     statusbar1.Visible := MenuItem4.Checked;

     stringgrid1.Color:= stringtocolor(inif.ReadString('Settings','Line1',colortostring(line1)));
     stringgrid1.AlternateColor:= stringtocolor(inif.ReadString('Settings','Line2',colortostring(line2)));

     stringgrid1.Font.Color:= stringtocolor(inif.ReadString('Settings','FontColor',colortostring(sgfontcolor)));
     stringgrid1.Font.Size:= strtoint(inif.ReadString('Settings','FontSize',colortostring(sgfontsize)));
     stringgrid1.Font.Name:=inif.ReadString('Settings','FontName',sgfontname);
     stringgrid1.Font.Style:=StrToFontStyle(inif.ReadString('Settings','FontStyle',FontStyleToStr(sgfontstyle)));

     stringgrid1.FixedColor := stringtocolor(inif.ReadString('Settings','labelc',colortostring(labelc)));

     stringgrid1.Repaint;
     finally
            inif.Free;
     end;
// INI READ END

end;

procedure TFormMain.FormDestroy(Sender: TObject);
var folder : string;
    inif : tinifile;
begin

folder := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)));
inif := tinifile.Create(folder + conf );

try
   if MenuItemSaveOnExit.Checked then begin

      inif.WriteBool('Settings','AutoResize',MenuItemAutoResize.Checked);
      inif.WriteBool('Settings','Label',MenuItemToggleLabel.Checked);
      inif.WriteString('Settings','Line1',ColorToString(stringgrid1.Color));
      inif.WriteString('Settings','Line2',ColorToString(stringgrid1.AlternateColor));
      inif.WriteString('Settings','FontColor',ColorToString(stringgrid1.Font.Color));
      inif.WriteString('Settings','FontSize',inttostr(stringgrid1.Font.Size));
      inif.WriteString('Settings','FontName',stringgrid1.Font.Name);
      inif.WriteString('Settings','FontStyle',FontStyleToStr(stringgrid1.Font.Style));
      inif.WriteString('Settings','labelc',ColorToString(stringgrid1.FixedColor));
      inif.WriteBool('View','ToolBar',MenuItem24.Checked);
      inif.WriteBool('View','StatusBar',MenuItem4.Checked);
      inif.WriteInteger('Main','LID',lid);
      inif.WriteInteger('Main','SID',sid);
      inif.WriteInteger('Main','EID',eid);

   end;

   inif.WriteBool('Main','SettingsSave',MenuItemSaveOnExit.Checked);
finally
   inif.Free;
end;

tlista.Free;
end;



//On drop files
procedure TFormMain.FormDropFiles(Sender: TObject; const FileNames: array of String);
var
   txt1:string;
begin

txt1 := copy(FileNames[0],strlen(pchar(FileNames[0]))-3,strlen(pchar(FileNames[0])));

if (AnsiLowerCase(txt1) <> '.csv') and (AnsiLowerCase(txt1) <> '.tab') and (AnsiLowerCase(txt1) <> '.tsv') and (AnsiLowerCase(txt1) <> '.txt') then begin
       application.MessageBox('This is file extension not supported!','Warning',0);
      end else begin

            Filename := FileNames[0];
            tlista.LoadFromFile(FileNames[0]);

            if elvalaszto(tlista.Strings[0]) <> '' then begin

              sep := elvalaszto(tlista.Strings[0]);

            end else begin

              sep := ',';

            end;

            LoadStringGrid(Filename);
            Statusbar1.Panels[2].Text := Filename;
            if sep = #9 then
              Statusbar1.Panels[1].Text := 'Tab'
            else
              Statusbar1.Panels[1].Text := sep;

      end;

end;



//on show form
procedure TFormMain.FormShow(Sender: TObject);
var i:integer;
    txt1:string;
    parameter : string;
begin

     sep := ',';
     Statusbar1.Panels[1].Text := sep;
     if ParamCount > 0 then begin
        for i := 1 to ParamCount do begin
            parameter:=parameter+ParamStr(i)+'';
        end;
     parameter:=''+parameter+'';


     if FileExists(parameter) then begin

        FileName := parameter;
        txt1 := copy(parameter,strlen(pchar(parameter))-3,strlen(pchar(parameter)));

        if (AnsiLowerCase(txt1) = '.csv') or
           (AnsiLowerCase(txt1) = '.tab') or
           (AnsiLowerCase(txt1) = '.tsv') or
           (AnsiLowerCase(txt1) = '.txt') then begin
	   tlista.LoadFromFile(FileName);

           if elvalaszto(tlista.Strings[0]) <> '' then begin

             sep := elvalaszto(tlista.Strings[0]);

          end else begin

            sep := ',';

          end;

            LoadStringGrid(FileName);
            Statusbar1.Panels[2].Text := Filename;
            if sep = #9 then
               Statusbar1.Panels[1].Text := 'Tab'
            else
                Statusbar1.Panels[1].Text := sep;
            end else begin
                application.MessageBox('This file extension is not supported!','Warning',0);
           end;
        end;
   end;
end;



procedure TFormMain.StringGrid1Click(Sender: TObject);
begin
  s_true := false;
end;


procedure TFormMain.StringGrid1DrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
begin
  Statusbar1.Panels[0].Text := 'X:' + IntToStr(stringgrid1.row+1) +
                               '  Y:' + IntToStr(stringgrid1.Col+1);

  with (Sender as TStringGrid) do begin


       if s_true then begin

          if (arow = stringgrid1.row)and (acol =stringgrid1.col) then begin
             Canvas.Brush.Color:= clyellow;
             Canvas.FillRect(aRect);

          end;
       end;
 end;

end;


///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//menu : FILE

//new file [clear stringrid]
procedure TFormMain.MenuItemNewClick(Sender: TObject);
var
  i: integer;
begin
     //sep := ',';
     FileName := '';
     tlista.Clear;
     Statusbar1.Panels[2].Text := '';
     //Statusbar1.Panels[1].Text := ',';
     for i := 0 to StringGrid1.RowCount - 1 do begin
         StringGrid1.Rows[i].Clear();
     end;

     StringGrid1.RowCount := 18;
     StringGrid1.ColCount := 6;
     StringGrid1.FixedRows := 0;
end;

procedure TFormMain.MenuItemOpenCSVClick(Sender: TObject);
var
open : TOpenDialog;
begin
open := TOpenDialog.Create(self);
open.Filter := filterAutoRecognition + ' (*.csv;*.tab;*.tsv;*.txt)|*.csv;*.tab;*.tsv;*.txt|' +
               filterCSV + ' (*.csv)|*.csv|' +
               filterCSVsemicolon + ' (*.csv)|*.csv|' +
               filterCSVpipe + ' (*.csv)|*.csv|' +
               filterCSVasterisk + ' (*.csv)|*.csv|' +
               filterCSVcolon + ' (*.csv)|*.csv|' +
               filterCSVdollar + ' (*.csv)|*.csv|' +
               filterCSVtab + ' (*.tab;*.tsv)|*.tab;*.tsv';
open.FilterIndex := lid;

if open.Execute then begin


    Filename := open.FileName;
    tlista.LoadFromFile(open.FileName);

    if open.FilterIndex = 1 then begin

       sep := elvalaszto(tlista.Strings[0]);

    end else begin

        if open.FilterIndex = 2 then
           sep := ',' ;

        if open.FilterIndex = 3 then
           sep := ';';

        if open.FilterIndex = 4 then
           sep := '|';

        if open.FilterIndex = 5 then
           sep := '*';

        if open.FilterIndex = 6 then
           sep := ':';

        if open.FilterIndex = 7 then
           sep := '$';

        if open.FilterIndex = 8 then
           sep := #9;

    end;
    lid := open.FilterIndex;
    LoadStringGrid(FileName);
    Statusbar1.Panels[2].Text := Filename;
    if sep = #9 then
       Statusbar1.Panels[1].Text := 'Tab'
    else
        Statusbar1.Panels[1].Text := sep;


end;

    open.Free;

end;


//Save
procedure TFormMain.MenuItemSaveClick(Sender: TObject);
begin
    if FileName <> '' then begin
       Sg2Csv(sep,FileName);
       application.MessageBox(pchar(format('File "%s" saved.',[FileName])),pchar('Save'),0);
    end else begin
        MenuItemSaveAsClick(Sender);
    end;
end;


//Save as
procedure TFormMain.MenuItemSaveAsClick(Sender: TObject);
var save : TSaveDialog;
begin
     save := TSaveDialog.Create(self);
     save.Filter := filterCSV + ' (*.csv)|*.csv|' +
                    filterCSVsemicolon + ' (*.csv)|*.csv|' +
                    filterCSVpipe + ' (*.csv)|*.csv|' +
                    filterCSVasterisk + ' (*.csv)|*.csv|' +
                    filterCSVcolon + ' (*.csv)|*.csv|' +
                    filterCSVDollar + ' (*.csv)|*.csv|' +
                    filterCSVtab + ' (*.tab)|*.tab|' +
                    filterCSVtab + ' (*.tsv)|*.tsv';

     save.DefaultExt:='.*.csv|.*.tab|.*.tsv'; //important ! if not set on linux i get empty extensions

     save.FilterIndex := sid;

     if save.Execute then begin
        if save.FilterIndex = 1 then
           sep := ',';
        if save.FilterIndex = 2 then
           sep := ';';
        if save.FilterIndex = 3 then
           sep := '|';
        if save.FilterIndex = 4 then
           sep := '*';
        if save.FilterIndex = 5 then
           sep := ':';
        if save.FilterIndex = 6 then
           sep := '$';
        if save.FilterIndex = 7 then
           sep := #9;
        if save.FilterIndex = 8 then
           sep := #9;

        sid := save.FilterIndex;

        if fileexists(save.FileName) then begin
           if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),pchar('Save as...'),MB_ICONQUESTION + MB_YESNO) = IDYES then begin
              sg2csv(sep,save.FileName);
              FileName := save.FileName;
           end;
        end else begin //file does not exist

            if sep = #9 then begin

               if save.FilterIndex = 7 then begin
                  sg2csv(sep,save.FileName); //'.tab'
                  FileName := save.FileName; //'.tab'
               end else begin
                   sg2csv(sep,save.FileName); //+'.tsv'
                   FileName := save.FileName; //'.tsv'
               end;

            end else begin
                sg2csv(sep,save.FileName); //+'.csv'
                FileName := save.FileName; //+'.csv'
            end;

        end;

        if sep = #9 then
           Statusbar1.Panels[1].Text := 'Tab'
        else
            Statusbar1.Panels[1].Text := sep;

        Statusbar1.Panels[2].Text := FileName;

     end;

     save.Free;

end;


//export
procedure TFormMain.MenuItemExportClick(Sender: TObject);
var save : TSaveDialog;
begin
save := TSaveDialog.Create(self);
save.Title := 'Export';
save.Filter := filterHtml + ' (*.html)|*.html|' +
               filterOds + ' (*.ods)|*.ods|' +
               filterXls + ' (*.xls)|*.xls|' +
               filterXml + ' (*.xml)|*.xml';

save.DefaultExt:='.*.html|.*.ods|.*.xls|.*.xml'; //important ! if not set on linux i get empty extensions

save.FilterIndex := eid;

if save.Execute then begin

   if save.FilterIndex = 1 then begin //html
        if FileExists(save.FileName) then begin
           if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),pchar('Save as...'), MB_ICONQUESTION + MB_YESNO) = IDYES then begin
              SGridToHtml(Stringgrid1,save.FileName);
           end;
        end else begin
            SGridToHtml(Stringgrid1,save.FileName); //+'.html'
        end;
   end else if save.FilterIndex = 2 then begin //ods
      if FileExists(save.FileName) then begin
         if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),pchar('Save as...'),MB_ICONQUESTION + MB_YESNO) = IDYES then begin
            Sg2ods(save.FileName);
         end;
      end else begin //ods does not exist
          Sg2ods(save.FileName);
      end;
   end else if save.FilterIndex = 3 then begin //xls
      if FileExists(save.FileName) then begin
        if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),pchar('Save as...'), MB_ICONQUESTION + MB_YESNO) = IDYES then begin
          Sg2Xls(save.FileName);
        end;
      end else begin
          Sg2Xls(save.FileName);
      end;
   end else if save.FilterIndex = 4 then begin //xml
     if FileExists(save.FileName) then begin
        if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),pchar('Save as...'),MB_ICONQUESTION + MB_YESNO) = IDYES then begin
           sg2xml(save.FileName);
        end;
     end else begin
         sg2xml(save.FileName);
     end;
   end;

  eid := save.FilterIndex;

end;
save.Free;
end;


//Take a snapshot
procedure TFormMain.MenuItemTakeSnapshotClick(Sender: TObject);
var
  MyBitmap: TBitmap;
  MyDC: HDC;
  jpg : TJPEGImage;
  save : TSaveDialog;
begin

  MyDC := GetDC(FormMain.Handle);
  MyBitmap := TBitmap.Create;
  MyBitmap.LoadFromDevice(MyDC);
  MyBitmap.Height:= height-50;

  jpg := TJPEGImage.Create;
  jpg.Assign(MyBitmap);
  save := TSaveDialog.Create(self);
  save.Title := 'Screenshot';
  save.Filter := 'Jpeg format (*.jpg)|*.jpg|' +
                 'Bitmap format (*.bmp)|*.bmp';

  save.DefaultExt:='.*.jpg|.*.bmp'; //important ! if not set on linux i get empty extensions

  if save.Execute then begin

     if save.FilterIndex = 1 then begin  //jpeg
        if FileExists(save.FileName)  then begin

           if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),'Confirmation', MB_ICONQUESTION + MB_YESNO) = IDYES then begin

           jpg.SaveToFile(save.FileName);

           end;

        end else begin

            jpg.SaveToFile(save.FileName);

        end;

     end else begin  //save.FilterIndex <> 1

         if fileexists(save.FileName) then begin

            if application.MessageBox(pchar(format('%s exists. Overwrite this file?',[save.FileName])),'Confirmation', MB_ICONQUESTION + MB_YESNO) = IDYES then begin

               MyBitmap.SaveToFile(save.FileName);

            end;

         end else begin

             mybitmap.SaveToFile(save.FileName);

         end;

     end;

  end;

  ReleaseDC(FormMain.Handle, MyDC);
  FreeAndNil(MyBitmap);
  save.Free;
  jpg.Free;
  //mybitmap.Free;

end;


//print
procedure TFormMain.MenuItemPrintClick(Sender: TObject);
begin

     if printdialog1.Execute then begin
        PrintGrid(StringGrid1);
     end;

end;

//exit program
procedure TFormMain.MenuItemQuitClick(Sender: TObject);
begin
  //application.Terminate;
  FormMain.Close;
end;


///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
///////////////////////////////////////////////////////
//////////////////////////////////////////////////////
//menu : EDIT

//Copy cell
procedure TFormMain.MenuItemCopyCellClick(Sender: TObject);
begin
    clipboard.AsText := stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row];
end;

//Paste cell
procedure TFormMain.MenuItemPasteCellClick(Sender: TObject);
begin
     if Clipboard.HasFormat(CF_TEXT) then
        stringgrid1.Cells[stringgrid1.Col,stringgrid1.Row] :=  clipboard.AsText;
end;


//////////////////////////////////////////////////////////
/////////////COLUMNS OPERATIONS///////////////////////////
//////////////////////////////////////////////////////////
//Add cols
procedure TFormMain.MenuItemAddColClick(Sender: TObject);
begin
    stringgrid1.ColCount:= stringgrid1.ColCount + 1;
    stringgrid1.MoveColRow(true,stringgrid1.ColCount-1, stringgrid1.Col+1);
end;

//duplicate cols
procedure TFormMain.MenuItemDuplicateColClick(Sender: TObject);
begin
     duplicate(true,stringgrid1);
end;

//remove cols
procedure TFormMain.MenuItemRemoveColsClick(Sender: TObject);
var range : trange;
    i: integer;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min>=0 then begin
        if stringgrid1.ColCount <> 0 then begin
           if application.MessageBox(pchar('Are you sure to delete from #'+inttostr(range.Min)+' to #' + inttostr(range.Max) + 'column(s)'),'Confirmation',MB_ICONQUESTION + MB_YESNO) = IDYES then begin

              for i:= range.Max downto range.Min do begin
                  stringgrid1.DeleteColRow(true,i);
              end;

           end;
        end;
     end;
end;

//move start col
procedure TFormMain.MenuItemMoveStartColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min=0 then begin
        stringgrid1.MoveColRow(true,range.Max, 0);
     end else showmessage('Please select only one column');
end;

//move left col
procedure TFormMain.MenuItemMoveLeftColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min=0 then begin
        if range.Max-1>=0 then
           stringgrid1.MoveColRow(true,range.Max, range.Max-1);
     end else showmessage('Please select only one column');
end;

//move right col
procedure TFormMain.MenuItemMoveRightColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min=0 then begin
     if range.Max+1 < stringgrid1.ColCount then
        stringgrid1.MoveColRow(true,range.Max, range.Max+1);
     end else showmessage('Please select only one column');
end;

//move end col
procedure TFormMain.MenuItemMoveEndColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min=0 then begin
        stringgrid1.MoveColRow(true,range.Max, stringgrid1.ColCount-1);
     end else showmessage('Please select only one column');

end;

//move left to right col
procedure TFormMain.MenuItemMoveLeftToRightColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min>0 then begin
        stringgrid1.MoveColRow(true,range.Min, range.Max);
     end else showmessage('Please select a range of columns');

end;

//move right to left col
procedure TFormMain.MenuItemMoveRightToLeftColClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(true,stringgrid1);
     if range.Max-range.Min>0 then begin
        stringgrid1.MoveColRow(true,range.Max, range.Min);
     end else showmessage('Please select a range of columns');
end;


//swap cols
procedure TFormMain.MenuItemSwapColumnsClick(Sender: TObject);
begin
     Swap(true,stringgrid1);
end;


//////////////////////////////////////////////////////////
/////////////ROWS   OPERATIONS////////////////////////////
//////////////////////////////////////////////////////////
//add rows
procedure TFormMain.MenuItemAddRowClick(Sender: TObject);
begin
  stringgrid1.RowCount:=  stringgrid1.RowCount + 1;
  stringgrid1.MoveColRow(false,stringgrid1.RowCount-1, stringgrid1.Row+1);
end;

//duplicate rows
procedure TFormMain.MenuItemDuplicateRowClick(Sender: TObject);
begin
     duplicate(false,stringgrid1);
end;

//remove rows
procedure TFormMain.MenuItemRemoveRowsClick(Sender: TObject);
var range : trange;
    i: integer;
begin

     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min>=0 then begin
        if stringgrid1.RowCount <> 0 then begin
           if application.MessageBox(pchar('Are you sure to delete from #'+inttostr(range.Min)+' to #' + inttostr(range.Max) + 'row(s)'),'Confirmation',MB_ICONQUESTION + MB_YESNO) = IDYES then begin

              for i:= range.Max downto range.Min do begin
                  stringgrid1.DeleteColRow(false,i);
              end;

           end;
        end;
     end;

end;

//move to top row
procedure TFormMain.MenuItemMoveTopRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min=0 then begin
        stringgrid1.MoveColRow(false,range.Max, 0);
     end else showmessage('Please select only one row');

end;

//move up row
procedure TFormMain.MenuItemMoveUpRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min=0 then begin
        if range.Max-1>=0 then
           stringgrid1.MoveColRow(false,range.Max, range.Max-1);
     end else showmessage('Please select only one row');
end;

//move down row
procedure TFormMain.MenuItemMoveDownRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min=0 then begin
     if range.Max+1 < stringgrid1.RowCount then
        stringgrid1.MoveColRow(false,range.Max, range.Max+1);
     end else showmessage('Please select only one row');
end;

//move bottom row
procedure TFormMain.MenuItemMoveBottomRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min=0 then begin
        stringgrid1.MoveColRow(false,range.Max, stringgrid1.RowCount-1);
     end else showmessage('Please select only one row');
end;

//move from top to bottom row
procedure TFormMain.MenuItemTopToBottomRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min>0 then begin
        stringgrid1.MoveColRow(false,range.Min, range.Max);
     end else showmessage('Please select a range of rows');
end;

//move from bottom to top row
procedure TFormMain.MenuItemBottomToTopRowClick(Sender: TObject);
var range : trange;
begin
     range:=GetSelectedRange(false,stringgrid1);
     if range.Max-range.Min>0 then begin
        stringgrid1.MoveColRow(false,range.Max, range.Min);
     end else showmessage('Please select a range of rows');
end;



//swap rows
procedure TFormMain.MenuItemSwapRowsClick(Sender: TObject);
begin
     swap(false,stringgrid1);
end;


//sort increase
procedure TFormMain.MenuItemSortColumnClick(Sender: TObject);
begin
  sortgrid(true,stringgrid1,stringgrid1.Col);
end;

//sort decrease
procedure TFormMain.ToolButtonSortDecreaseClick(Sender: TObject);
begin
  sortgrid(false,stringgrid1,stringgrid1.Col);
end;


//////////////////////////////////////////////////
/////////////////////////////////////////////////
//////////////////////////////////////////////////
/////////////////////////////////////////////////
//menu : SEARCH

//search
procedure TFormMain.MenuItemSearchClick(Sender: TObject);
begin
  FindDialog1.Execute;
end;

//internet search
procedure TFormMain.MenuItemSearchOnlineClick(Sender: TObject);
var cell: string;
begin
  cell := StringGrid1.Cells[StringGrid1.Col,StringGrid1.Row];
  GoToURL('http://www.google.com/search?q=' + cell);
end;



//Find dialog
procedure TFormMain.FindDialog1Find(Sender: TObject);
var
  CurX, CurY, GridWidth, GridHeight: integer;
  X, Y: integer;
  TargetText: string;
  CellText: string;
  i: integer;
  GridRect: TGridRect;
label
  TheEnd;
begin
  CurX := StringGrid1.Selection.Left + 1;
  CurY := StringGrid1.Selection.Top;
  GridWidth := StringGrid1.ColCount;
  GridHeight := StringGrid1.RowCount;
  Y := CurY;
  X := CurX;
  if frMatchCase in FindDialog1.Options then
    TargetText := FindDialog1.FindText
  else
    TargetText := AnsiLowerCase(FindDialog1.FindText);
  while Y < GridHeight do
  begin
    while X < GridWidth do
    begin
      if frMatchCase in FindDialog1.Options then
        CellText := StringGrid1.Cells[X, Y]
      else
        CellText := AnsiLowerCase(StringGrid1.Cells[X, Y]);
      i := Pos(TargetText, CellText) ;

      if i > 0 then
      begin

        GridRect.Left := X;
        GridRect.Right := X;
        GridRect.Top := Y;
        GridRect.Bottom := Y;

          with StringGrid1 do
            Begin
            Col := GridRect.Left;
            Row := GridRect.Bottom;
            Selection := GridRect;
        {GetParentForm(StringGrid1).SetFocus;}
    SetFocus;
    s_true := true;
    //StringGrid1.EditorMode := true;
    {TCustomEdit(Components[0]).SelStart := i - 1;
    TCustomEdit(Components[0]).SelLength := length(TargetText);}

          end;

        goto TheEnd;
      end;
      inc(X);
    end;
    inc(Y);
    X := StringGrid1.FixedCols;

  end;
TheEnd:

end;


////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////
//menu : View
procedure TFormMain.MenuItemToggleToolBarClick(Sender: TObject);
begin
     toolbar1.Visible := not toolbar1.Visible;
end;

procedure TFormMain.MenuItemToggleStatusBarClick(Sender: TObject);
begin
  statusbar1.Visible := not statusbar1.Visible;
end;


////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////
//menu : Settings

//Autoresize
procedure TFormMain.MenuItemAutoresizeClick(Sender: TObject);
begin
     stringgrid1.AutoFillColumns := not stringgrid1.AutoFillColumns;
end;

//Toggle Label
procedure TFormMain.MenuItemToggleLabelClick(Sender: TObject);
begin
  if stringgrid1.Fixedrows > 0 then begin
         stringgrid1.Fixedrows := 0;
  end else begin
        stringgrid1.Fixedrows := 1;
  end;

end;

//select fonts
procedure TFormMain.MenuItemFontClick(Sender: TObject);
begin

  FontDialog1.Font:=StringGrid1.Font;
  if FontDialog1.Execute then begin
     StringGrid1.Font:=FontDialog1.Font;
     stringgrid1.Font.Color:=FontDialog1.Font.Color;
     stringgrid1.Font.Size:=FontDialog1.Font.Size;
     stringgrid1.Font.Name:=FontDialog1.Font.Name;
     stringgrid1.Font.Style:=FontDialog1.Font.Style;
  end;
end;

//Color line 1
procedure TFormMain.MenuItemColorLine1Click(Sender: TObject);
var colorb :TColorDialog;
begin
     colorb := TColorDialog.Create(self);
     if colorb.Execute then begin
        stringgrid1.Color:= colorb.Color;
     end;
     colorb.Free;
end;

//Color line 2
procedure TFormMain.MenuItemColorLine2Click(Sender: TObject);
var colorb :TColorDialog;
begin
     colorb := TColorDialog.Create(self);
     if colorb.Execute then begin
        stringgrid1.AlternateColor:= colorb.Color;
     end;
     colorb.Free;
end;

//Color label
procedure TFormMain.MenuItemColorLabelClick(Sender: TObject);
var colorb :TColorDialog;
begin
     colorb := TColorDialog.Create(self);
     if colorb.Execute then begin
        stringgrid1.FixedColor := colorb.Color;
     end;
     colorb.Free;
end;

//on MenuItemSaveOnExit toggle MenuItemSaveOnExit
procedure TFormMain.MenuItemSaveOnExitClick(Sender: TObject);
begin
  MenuItemSaveOnExit.Checked := not MenuItemSaveOnExit.Checked;
end;

//restore default values
procedure TFormMain.MenuItemRestoreDefaultsClick(Sender: TObject);
begin

   MenuItemAutoResize.Checked:=false;
   stringgrid1.AutoFillColumns:=false;
   MenuItemToggleLabel.Checked:=true;
   stringgrid1.Fixedrows := 1;
   stringgrid1.Font.Size := sgfontsize;
   stringgrid1.Font.Color := sgfontcolor;
   stringgrid1.Font.Name := sgfontname;
   stringgrid1.Font.Style := sgfontstyle;
   stringgrid1.AlternateColor := line1;
   stringgrid1.Color := line2;
   stringgrid1.FixedColor := labelc;
   MenuItemSaveOnExit.Checked:=true;

end;


////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////
//menu : Help


//About box
procedure TFormMain.MenuItemAboutClick(Sender: TObject);
var w:tform;
    a,b,c,d,e:tlabel;

begin

w := tform.Create(self);

  with w do begin
    caption := 'About Box';
    position:= pomainformcenter;
    borderstyle:= bsdialog;
    width:= 500;
    height:= 160;
    color:= $00eeeeee;

  end;

  a := tlabel.Create(self);
  b := tlabel.Create(self);
  c := tlabel.Create(self);
  d := tlabel.Create(self);
  e := tlabel.Create(self);

  with a do begin
    parent := w;
    top := 10;
    alignment := tacenter;
    font.Size := 12;
    font.Style := [fsbold];
    width:= 490;
    height:=30;
    autosize:=false;
    font.color:=$00c76001;
  end;

  with b do begin
    parent := w;
    top := 40;
    alignment := tacenter;
    font.Size := 10;
    font.Style := [fsbold];
    width:= 490;
    autosize:=false;
    font.color:=$000174f0;
  end;

  with c do begin
    parent := w;
    top := 60;
    alignment := tacenter;
    font.Size := 10;
    font.Style := [fsunderline];
    width:= 490;
    autosize:=false;
    font.color:=$00000000;
  end;

  with d do begin
    parent := w;
    top := 80;
    alignment := tacenter;
    font.Size := 8;
    font.Style := [fsbold];
    width:= 490;
    autosize:=false;
    font.color:=$00666666;
    //wordwrap := true;
  end;

   with e do begin
    parent := w;
    top := 100;
    alignment := tacenter;
    font.Size := 8;
    font.Style := [fsbold];
    width:= 490;
    autosize:=false;
    font.color:=$00666666;
    wordwrap := true;
  end;


  a.Caption := 'CSVpad v' + version + ' By TrustFm [www.trustfm.net]';
  b.Caption := 'Original code by: Tivadar (Darh Media) 2003 - 2013 [darhmedia.blogspot.hu]';
  c.Caption := 'Thanks to:';
  d.Caption := 'Tivadar (Darh Media) DMcsvEditor, Vladimir Zhirov, Christian Ebenegger, Pinvoke';
  e.caption := 'CodeTyphon, Lazarus, FreePascal, UPX team';
  w.ShowModal;

  a.Free;
  b.Free;
  c.Free;
  d.Free;
  e.free;
  w.Free;
end;


//Go to website
procedure TFormMain.MenuItemWebpageClick(Sender: TObject);
begin
  GoToURL('http://www.trustfm.net');
end;


//Donate trustfm
procedure TFormMain.MenuItemDonateClick(Sender: TObject);
begin
  GoToURL('https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=M52J6BGABLMQU&lc=US');
end;


//donate 2
procedure TFormMain.MenuItemDonate2Click(Sender: TObject);
begin
    GoToURL('https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=WMTEWGNAQ4Y9J');
end;


end.

