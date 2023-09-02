unit AppointmentUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, dsCalendar, DBGridEhGrouping,
  ToolCtrlsEh, Halcn6db, DB, DBTables, GridsEh, DBGridEh, Mask, IniFiles,
  ExtCtrls, Menus, DateUtils, ComObj;


type
  TAppointmentForm = class(TForm)
    dsCalendar1: TdsCalendar;
    PageControl1: TPageControl;
    Tab1: TTabSheet;
    Tab2: TTabSheet;
    Tab3: TTabSheet;
    BitBtn7: TBitBtn;
    DBGrid2: TDBGridEh;
    BitBtn1: TBitBtn;
    DataSource1: TDataSource;
    DayTable: TTable;
    DataSource2: TDataSource;
    AppTable: THalcyonDataSet;
    DBGrid1: TDBGridEh;
    AppTable_forCreate: TTable;
    AppTablePOSTCODE: TStringField;
    AppTableDATE_APP: TDateField;
    AppTableTIME1: TStringField;
    AppTableTIME2: TStringField;
    AppTableK_KL: TStringField;
    AppTableCAR_N_ZAP: TStringField;
    AppTableD_UDL: TDateField;
    AppTableIO_UDL: TStringField;
    AppTableIO_ADD: TStringField;
    AppTableD_ADD: TDateField;
    AppTableT_ADD: TStringField;
    AppTableGOS_N: TStringField;
    AppTableVIN: TStringField;
    AppTableMARKA: TStringField;
    AppTableOBR_KL: TMemoField;
    AppTableCOMM: TMemoField;
    AppTableTAB_N: TStringField;
    AppTableN_ZAJ: TStringField;
    Panel0: TPanel;
    Panel1: TPanel;
    Label3: TLabel;
    Panel2: TPanel;
    Label4: TLabel;
    Panel3: TPanel;
    Label5: TLabel;
    AppTableNAVL_IO: TStringField;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    AppTableTEL: TStringField;
    AppTableDOP_PR: TStringField;
    AppTableREC_ID: TStringField;
    PopupMenu2: TPopupMenu;
    N3: TMenuItem;
    DetailPanel: TPanel;
    Panel5: TPanel;
    SaveBitBtn: TBitBtn;
    CancelBitBtn: TBitBtn;
    SpeedButton1: TSpeedButton;
    Bevel1: TBevel;
    PageControl2: TPageControl;
    Tab10: TTabSheet;
    Tab20: TTabSheet;
    CommentMemo: TMemo;
    ReasonMemo: TMemo;
    ConfirmSMSCheckBox: TCheckBox;
    Label6: TLabel;
    Label7: TLabel;
    ClientNameMaskEdit: TMaskEdit;
    ClientSearchBitBtn: TBitBtn;
    MarkaMaskEdit: TMaskEdit;
    Label8: TLabel;
    telCountryCodeMaskEdit: TMaskEdit;
    Label88: TLabel;
    telCodeMaskEdit: TMaskEdit;
    Label126: TLabel;
    telNumberMaskEdit: TMaskEdit;
    ModelMaskEdit: TMaskEdit;
    StartTimeMaskEdit: TMaskEdit;
    Label9: TLabel;
    EndTimeMaskEdit: TMaskEdit;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    CarNumberMaskEdit: TMaskEdit;
    Label13: TLabel;
    CardNumberMaskEdit: TMaskEdit;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    AppTableNAVL_PR: TStringField;
    AppTable_forCreateDATE_APP: TDateField;
    AppTable_forCreateTIME1: TStringField;
    AppTable_forCreateTIME2: TStringField;
    AppTable_forCreatePOSTCODE: TStringField;
    AppTable_forCreateGOS_N: TStringField;
    AppTable_forCreateVIN: TStringField;
    AppTable_forCreateMARKA: TStringField;
    AppTable_forCreateK_KL: TStringField;
    AppTable_forCreateTEL: TStringField;
    AppTable_forCreateCAR_N_ZAP: TStringField;
    AppTable_forCreateOBR_KL: TMemoField;
    AppTable_forCreateCOMM: TMemoField;
    AppTable_forCreateTAB_N: TStringField;
    AppTable_forCreateN_ZAJ: TStringField;
    AppTable_forCreateNAVL_IO: TStringField;
    AppTable_forCreateNAVL_PR: TStringField;
    AppTable_forCreateDOP_PR: TStringField;
    AppTable_forCreateD_UDL: TDateField;
    AppTable_forCreateIO_UDL: TStringField;
    AppTable_forCreateIO_ADD: TStringField;
    AppTable_forCreateD_ADD: TDateField;
    AppTable_forCreateT_ADD: TStringField;
    AppTable_forCreateREC_ID: TStringField;
    DataSource3: TDataSource;
    AppTableTime: THalcyonDataSet;
    AppTableTimeDATE_APP: TDateField;
    AppTableTimeTIME1: TStringField;
    AppTableTimeTIME2: TStringField;
    AppTableTimePOSTCODE: TStringField;
    AppTableTimeGOS_N: TStringField;
    AppTableTimeVIN: TStringField;
    AppTableTimeMARKA: TStringField;
    AppTableTimeK_KL: TStringField;
    AppTableTimeTEL: TStringField;
    AppTableTimeCAR_N_ZAP: TStringField;
    AppTableTimeOBR_KL: TMemoField;
    AppTableTimeCOMM: TMemoField;
    AppTableTimeTAB_N: TStringField;
    AppTableTimeN_ZAJ: TStringField;
    AppTableTimeNAVL_IO: TStringField;
    AppTableTimeNAVL_PR: TStringField;
    AppTableTimeDOP_PR: TStringField;
    AppTableTimeD_UDL: TDateField;
    AppTableTimeIO_UDL: TStringField;
    AppTableTimeIO_ADD: TStringField;
    AppTableTimeD_ADD: TDateField;
    AppTableTimeT_ADD: TStringField;
    AppTableTimeREC_ID: TStringField;
    Panel6: TPanel;
    DBGrid3: TDBGridEh;
    NewRecBitBtn: TBitBtn;
    CreateOrderBitBtn: TBitBtn;
    DeleteBitBtn: TBitBtn;
    Panel7: TPanel;
    SpeedButton2: TSpeedButton;
    SelectedTimeLAbel: TLabel;
    Label1: TLabel;
    ScaleComboBox: TComboBox;
    Label2: TLabel;
    AppTable_forCreateBRIGHT: TStringField;
    AppTableBRIGHT: TStringField;
    AppTableTimeBRIGHT: TStringField;
    N4: TMenuItem;
    N5: TMenuItem;
    EditOrderBitBtn2: TBitBtn;
    N6: TMenuItem;
    N7: TMenuItem;
    CreateOrderBitBtn3: TBitBtn;
    EditRecBitBtn: TBitBtn;
    DelRecBitBtn: TBitBtn;
    PopupMenu3: TPopupMenu;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    Panel8: TPanel;
    Label20: TLabel;
    TelMaskEditCountryCode: TMaskEdit;
    Label22: TLabel;
    TelMaskEditCode: TMaskEdit;
    Label23: TLabel;
    TelMaskEditNumber: TMaskEdit;
    CarNumberFindEdit: TMaskEdit;
    Label21: TLabel;
    Label19: TLabel;
    Label233: TLabel;
    MonMaskEdit: TMaskEdit;
    Label263: TLabel;
    Label265: TLabel;
    TueMaskEdit: TMaskEdit;
    WedMaskEdit: TMaskEdit;
    Label266: TLabel;
    Label267: TLabel;
    ThuMaskEdit: TMaskEdit;
    FriMaskEdit: TMaskEdit;
    Label264: TLabel;
    Label268: TLabel;
    SatMaskEdit: TMaskEdit;
    SunMaskEdit: TMaskEdit;
    Label270: TLabel;
    Bevel2: TBevel;
    SearchByCardBitBtn: TBitBtn;
    Rec_IDLAbel: TLabel;
    Label26: TLabel;
    DateTimePicker1: TDateTimePicker;
    WorkPostsComboBox: TComboBox;
    k_klLabel: TLabel;
    Car_n_zapLabel: TLabel;
    StoPostComboBox: TComboBox;
    AppTable_forCreateFIO: TStringField;
    AppTableFIO: TStringField;
    AppTableTimeFIO: TStringField;
    ClientSpravBitBtn: TBitBtn;
    Panel9: TPanel;
    ClientCarsListBox: TListBox;
    Panel10: TPanel;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    CarsQuery: TQuery;
    CarsQuerymarka: TStringField;
    CarsQuerymodel: TStringField;
    CarsQueryAUTO: TStringField;
    CarsQueryGOS_N: TStringField;
    CarsQueryK_KL: TStringField;
    CarsQueryN_ZAP: TStringField;
    CarsQueryVIN: TStringField;
    CarsQueryYEAR: TStringField;
    AppTable_forCreateMODEL: TStringField;
    AppTableMODEL: TStringField;
    AppTableTimeMODEL: TStringField;
    AppTableForAdd: THalcyonDataSet;
    AppTableForAddDATE_APP: TDateField;
    AppTableForAddTIME1: TStringField;
    AppTableForAddTIME2: TStringField;
    AppTableForAddPOSTCODE: TStringField;
    AppTableForAddGOS_N: TStringField;
    AppTableForAddVIN: TStringField;
    AppTableForAddMARKA: TStringField;
    AppTableForAddMODEL: TStringField;
    AppTableForAddK_KL: TStringField;
    AppTableForAddFIO: TStringField;
    AppTableForAddTEL: TStringField;
    AppTableForAddCAR_N_ZAP: TStringField;
    AppTableForAddOBR_KL: TMemoField;
    AppTableForAddCOMM: TMemoField;
    AppTableForAddTAB_N: TStringField;
    AppTableForAddN_ZAJ: TStringField;
    AppTableForAddNAVL_IO: TStringField;
    AppTableForAddNAVL_PR: TStringField;
    AppTableForAddDOP_PR: TStringField;
    AppTableForAddD_UDL: TDateField;
    AppTableForAddIO_UDL: TStringField;
    AppTableForAddIO_ADD: TStringField;
    AppTableForAddD_ADD: TDateField;
    AppTableForAddT_ADD: TStringField;
    AppTableForAddBRIGHT: TStringField;
    AppTableForAddREC_ID: TStringField;
    Label17: TLabel;
    SQLQuery1: TQuery;
    DateTimePicker2: TDateTimePicker;
    DateTimePicker3: TDateTimePicker;
    Label25: TLabel;
    SpeedButton5: TSpeedButton;
    PopupMenu4: TPopupMenu;
    N14: TMenuItem;
    D_offTableForCreate: TTable;
    D_offTableForCreateKOD: TStringField;
    D_offTableForCreateD1: TDateField;
    D_offTableForCreateIO_ADD: TStringField;
    D_offTableForCreateD_ADD: TDateField;
    D_offTable: THalcyonDataSet;
    D_offTableForCreatePR: TStringField;
    D_offTableKOD: TStringField;
    D_offTableD1: TDateField;
    D_offTablePR: TStringField;
    D_offTableIO_ADD: TStringField;
    D_offTableD_ADD: TDateField;
    DBGrid4: TDBGridEh;
    D_offTableSTO: THalcyonDataSet;
    DataSource4: TDataSource;
    D_offTableSTOKOD: TStringField;
    D_offTableSTOD1: TDateField;
    D_offTableSTOPR: TStringField;
    D_offTableSTOIO_ADD: TStringField;
    D_offTableSTOD_ADD: TDateField;
    dsCalendar2: TdsCalendar;
    ResourceItemsComboBox: TComboBox;
    FindByCarNumberBitBtn: TBitBtn;
    BitBtn15: TBitBtn;
    Label24: TLabel;
    MonthsComboBox: TComboBox;
    YearComboBox: TComboBox;
    DBGrid7: TDBGridEh;
    MonthTable: TTable;
    MonthTableD1: TStringField;
    MonthTableD2: TStringField;
    MonthTableD3: TStringField;
    MonthTableD4: TStringField;
    MonthTableD5: TStringField;
    MonthTableD6: TStringField;
    MonthTableD7: TStringField;
    DataSource5: TDataSource;
    PostCodeLAbel: TLabel;
    MonthTablePosts: TTable;
    DataSource6: TDataSource;
    MonthTablePostsD1: TStringField;
    MonthTablePostsD2: TStringField;
    MonthTablePostsD3: TStringField;
    MonthTablePostsD4: TStringField;
    MonthTablePostsD5: TStringField;
    MonthTablePostsD6: TStringField;
    MonthTablePostsD7: TStringField;
    MonthTablePostsD8: TStringField;
    MonthTablePostsD9: TStringField;
    MonthTablePostsD10: TStringField;
    MonthTablePostsD11: TStringField;
    MonthTablePostsD12: TStringField;
    MonthTablePostsD13: TStringField;
    MonthTablePostsD14: TStringField;
    MonthTablePostsD15: TStringField;
    MonthTablePostsD16: TStringField;
    MonthTablePostsD17: TStringField;
    MonthTablePostsD18: TStringField;
    MonthTablePostsD19: TStringField;
    MonthTablePostsD20: TStringField;
    MonthTablePostsD21: TStringField;
    MonthTablePostsD22: TStringField;
    MonthTablePostsD23: TStringField;
    MonthTablePostsD24: TStringField;
    MonthTablePostsD25: TStringField;
    MonthTablePostsD26: TStringField;
    MonthTablePostsD27: TStringField;
    MonthTablePostsD28: TStringField;
    MonthTablePostsD29: TStringField;
    MonthTablePostsD30: TStringField;
    MonthTablePostsD31: TStringField;
    DBGrid9: TDBGridEh;
    MonthTablePostsPOSTCODE: TStringField;
    CreateOrderBitBtn2: TBitBtn;
    Tab4: TTabSheet;
    AppTAbleMonth: THalcyonDataSet;
    AppTAbleMonthDATE_APP: TDateField;
    AppTAbleMonthTIME1: TStringField;
    AppTAbleMonthTIME2: TStringField;
    AppTAbleMonthPOSTCODE: TStringField;
    AppTAbleMonthGOS_N: TStringField;
    AppTAbleMonthVIN: TStringField;
    AppTAbleMonthMARKA: TStringField;
    AppTAbleMonthMODEL: TStringField;
    AppTAbleMonthK_KL: TStringField;
    AppTAbleMonthFIO: TStringField;
    AppTAbleMonthTEL: TStringField;
    AppTAbleMonthCAR_N_ZAP: TStringField;
    AppTAbleMonthOBR_KL: TMemoField;
    AppTAbleMonthCOMM: TMemoField;
    AppTAbleMonthTAB_N: TStringField;
    AppTAbleMonthN_ZAJ: TStringField;
    AppTAbleMonthNAVL_IO: TStringField;
    AppTAbleMonthNAVL_PR: TStringField;
    AppTAbleMonthDOP_PR: TStringField;
    AppTAbleMonthD_UDL: TDateField;
    AppTAbleMonthIO_UDL: TStringField;
    AppTAbleMonthIO_ADD: TStringField;
    AppTAbleMonthD_ADD: TDateField;
    AppTAbleMonthT_ADD: TStringField;
    AppTAbleMonthBRIGHT: TStringField;
    AppTAbleMonthREC_ID: TStringField;
    DataSource7: TDataSource;
    DBGrid20: TDBGridEh;
    AllPostsComboBox: TComboBox;
    BitBtn56: TBitBtn;
    BitBtn57: TBitBtn;
    Tab_nTable: THalcyonDataSet;
    Tab_nTableTAB_N: TStringField;
    Tab_nTableFIO: TStringField;
    Tab_nTableMARKA: TStringField;
    Tab_nTablePROC: TFloatField;
    Tab_nTableC_NCH: TFloatField;
    Tab_nTableVALUTA: TStringField;
    Tab_nTableD_I: TDateField;
    Tab_nTableIO: TStringField;
    MonthTableStaff: TTable;
    DataSource8: TDataSource;
    MonthTableStaffSTAFFCODE: TStringField;
    MonthTableStaffD1: TStringField;
    MonthTableStaffD2: TStringField;
    MonthTableStaffD3: TStringField;
    MonthTableStaffD4: TStringField;
    MonthTableStaffD5: TStringField;
    MonthTableStaffD6: TStringField;
    MonthTableStaffD7: TStringField;
    MonthTableStaffD8: TStringField;
    MonthTableStaffD9: TStringField;
    MonthTableStaffD10: TStringField;
    MonthTableStaffD11: TStringField;
    MonthTableStaffD12: TStringField;
    MonthTableStaffD13: TStringField;
    MonthTableStaffD14: TStringField;
    MonthTableStaffD15: TStringField;
    MonthTableStaffD16: TStringField;
    MonthTableStaffD17: TStringField;
    MonthTableStaffD18: TStringField;
    MonthTableStaffD19: TStringField;
    MonthTableStaffD20: TStringField;
    MonthTableStaffD21: TStringField;
    MonthTableStaffD22: TStringField;
    MonthTableStaffD23: TStringField;
    MonthTableStaffD24: TStringField;
    MonthTableStaffD25: TStringField;
    MonthTableStaffD26: TStringField;
    MonthTableStaffD27: TStringField;
    MonthTableStaffD28: TStringField;
    MonthTableStaffD29: TStringField;
    MonthTableStaffD30: TStringField;
    MonthTableStaffD31: TStringField;
    StaffListBox: TListBox;
    Label27: TLabel;
    BitBtn17: TBitBtn;
    N15: TMenuItem;
    N16: TMenuItem;
    addRecordCheckBox: TCheckBox;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    RezhimLabel: TLabel;
    Label28: TLabel;
    OrderNumberMaskEdit: TMaskEdit;
    N21: TMenuItem;
    N22: TMenuItem;
    OpenOrderBitBtn: TBitBtn;
    N23: TMenuItem;
    N24: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure dsCalendar1DateChange(Sender: TObject; FromDate,
      ToDate: TDateTime);
    procedure ScaleComboBoxChange(Sender: TObject);
    procedure MonMaskEditChange(Sender: TObject);
    procedure SatMaskEditChange(Sender: TObject);
    procedure MonMaskEditClick(Sender: TObject);
    procedure TueMaskEditClick(Sender: TObject);
    procedure TueMaskEditChange(Sender: TObject);
    procedure MonMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TueMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure WedMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ThuMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FriMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SatMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SunMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure DBGrid1DrawDataCell(Sender: TObject; const Rect: TRect;
      Field: TField; State: TGridDrawState);
    procedure PageControl1Change(Sender: TObject);
    procedure DBGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumnEh);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumnEh; State: TGridDrawState);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SaveBitBtnClick(Sender: TObject);
    procedure CancelBitBtnClick(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure telCountryCodeMaskEditChange(Sender: TObject);
    procedure telCodeMaskEditChange(Sender: TObject);
    procedure telNumberMaskEditChange(Sender: TObject);
    procedure telCountryCodeMaskEditClick(Sender: TObject);
    procedure telCodeMaskEditClick(Sender: TObject);
    procedure telNumberMaskEditClick(Sender: TObject);
    procedure NewRecBitBtnClick(Sender: TObject);
    procedure DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumnEh; State: TGridDrawState);
    procedure SpeedButton2Click(Sender: TObject);
    procedure DeleteBitBtnClick(Sender: TObject);
    procedure PopupMenu2Popup(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure EditOrderBitBtn2Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure TelMaskEditCountryCodeChange(Sender: TObject);
    procedure TelMaskEditCountryCodeClick(Sender: TObject);
    procedure TelMaskEditCodeChange(Sender: TObject);
    procedure TelMaskEditCodeClick(Sender: TObject);
    procedure TelMaskEditNumberClick(Sender: TObject);
    procedure TelMaskEditNumberChange(Sender: TObject);
    procedure CarNumberFindEditClick(Sender: TObject);
    procedure CarNumberFindEditKeyPress(Sender: TObject; var Key: Char);
    procedure EditRecBitBtnClick(Sender: TObject);
    procedure DelRecBitBtnClick(Sender: TObject);
    procedure DetailPanelExit(Sender: TObject);
    procedure DBGrid3DblClick(Sender: TObject);
    procedure DBGrid2DblClick(Sender: TObject);
    procedure ClientSearchBitBtnClick(Sender: TObject);
    procedure ClientSpravBitBtnClick(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure ClientCarsListBoxMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ClientCarsListBoxClick(Sender: TObject);
    procedure SearchByCardBitBtnClick(Sender: TObject);
    procedure CardNumberMaskEditClick(Sender: TObject);
    procedure CardNumberMaskEditKeyPress(Sender: TObject; var Key: Char);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure StartTimeMaskEditClick(Sender: TObject);
    procedure EndTimeMaskEditClick(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure CarNumberMaskEditClick(Sender: TObject);
    procedure ClientNameMaskEditClick(Sender: TObject);
    procedure FindByCarNumberBitBtnClick(Sender: TObject);
    procedure CarNumberMaskEditKeyPress(Sender: TObject; var Key: Char);
    procedure CarNumberMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure MarkaMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ModelMaskEditKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtn15Click(Sender: TObject);
    procedure dsCalendar2DateSelect(Sender: TObject);
    procedure ResourceItemsComboBoxChange(Sender: TObject);
    procedure MonthsComboBoxChange(Sender: TObject);
    procedure YearComboBoxChange(Sender: TObject);
    procedure DBGrid7DrawDataCell(Sender: TObject; const Rect: TRect;
      Field: TField; State: TGridDrawState);
    procedure DBGrid7CellClick(Column: TColumnEh);
    procedure DBGrid7MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure DBGrid9DrawDataCell(Sender: TObject; const Rect: TRect;
      Field: TField; State: TGridDrawState);
    procedure DBGrid9MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure DBGrid9CellClick(Column: TColumnEh);
    procedure CreateOrderBitBtn2Click(Sender: TObject);
    procedure CreateOrderBitBtn3Click(Sender: TObject);
    procedure CreateOrderBitBtnClick(Sender: TObject);
    procedure DBGrid20DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumnEh; State: TGridDrawState);
    procedure ClientNameMaskEditDblClick(Sender: TObject);
    procedure BitBtn56Click(Sender: TObject);
    procedure BitBtn57Click(Sender: TObject);
    procedure SunMaskEditChange(Sender: TObject);
    procedure FriMaskEditChange(Sender: TObject);
    procedure ThuMaskEditChange(Sender: TObject);
    procedure WedMaskEditChange(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure BitBtn17Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure AllPostsComboBoxChange(Sender: TObject);
    procedure WorkPostsComboBoxChange(Sender: TObject);
    procedure OrderNumberMaskEditClick(Sender: TObject);
    procedure OrderNumberMaskEditChange(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure OpenOrderBitBtnClick(Sender: TObject);
    procedure N24Click(Sender: TObject);
    procedure PopupMenu3Popup(Sender: TObject);
  private
    { Private declarations }

    //Она нужна, что бы для формы создавалась иконка в трее
    procedure CreateParams(var Params: TCreateParams); override;

  public
    { Public declarations }

  protected

  end;

var
  AppointmentForm: TAppointmentForm;
  Global_X, Global_Y, LastPaint_X, LastPaint_Y, FieldNAIM_Righ_X: Integer;
  Global_ColumnName, RealColumnNAme, key1a, key2a: String;
  FormForSendSMS_addRec, FormForSendSMS_editRec: Boolean;


implementation

uses Unit1, Unit_install, navigator_unit, Unit2, findUnit, zajavki,
     CcfUnit, SMSUnit;


{$R *.dfm}


procedure createNewZajProc(Table: THalcyonDataSet; k_kl,car_n_zap: string);
var
  n_dok: string;
  ccf_za_nal, ccf_za_b_nal: extended;
begin

  ZajavkiForm.Close;



  InsertForm.ZajQuery.SQL.Clear;
  InsertForm.ZajQuery.SQL.Add('SELECT * FROM "'+Path_base+'zakaz\20'+FormatDateTime('yy',date)+'\zaj_o.dbf" WHERE K_KL=:K_KL');
  InsertForm.ZajQuery.ParamByName('K_KL').AsString:=k_kl;
  InsertForm.ZajQuery.Open;
  InsertForm.ZajQuery.LAst;

  if (InsertForm.ZajQuery.FieldBYNAme('N_DOK').AsString<>'') and  (InsertForm.ZajQuery.FieldBYNAme('SOST').AsString<>'2')
      and  (InsertForm.ZajQuery.FieldBYNAme('SOST').AsString<>'3') and  (InsertForm.ZajQuery.FieldBYNAme('SOST').AsString<>'A') then
      begin
      n_dok:=InsertForm.ZajQuery.FieldBYNAme('N_DOK').AsString;
      if my_dlg('Внимание!',PChar('Для клиента уже есть открытая заявка №'+n_dok+' от '+InsertForm.ZajQuery.FieldBYNAme('D_O').AsString+'.'+#13+'Все равно создать новую?'),clYellow)=FAlse then
         begin
         ZajavkiForm.Show;
         if ZajavkiForm.Zaj_oQuery.Locate('N_DOK',n_dok,[])=True then
            begin
            ZajavkiForm.DBGrid2DBLClick(ZajavkiForm.DBGrid2);
            AppointmentForm.Close;
            end
         else
            my_messageTime('Внимание!','Документ не найден.',clYellow,3000);


         exit;
         end;


      end;




  if ZajavkiForm.Zaj_oTable_for_add.TableName<>path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.dbf' then
     begin
     ZajavkiForm.Zaj_oTable_for_add.Close;
     ZajavkiForm.Zaj_oTable_for_add.IndexFiles.Clear;
     ZajavkiForm.Zaj_oTable_for_add.TableName:=path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.dbf';
     ZajavkiForm.Zaj_oTable_for_add.IndexFiles.Add(path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.cdx');
     end;

  if ZajavkiForm.Zaj_oTable_forFind.TableName<>path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.dbf' then
     begin
     ZajavkiForm.Zaj_oTable_forFind.Close;
     ZajavkiForm.Zaj_oTable_forFind.IndexFiles.Clear;
     ZajavkiForm.Zaj_oTable_forFind.TableName:=path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.dbf';
     ZajavkiForm.Zaj_oTable_forFind.IndexFiles.Add(path_base+'ZAKAZ\'+FormatDateTime('YYYY',date)+'\zaj_o.cdx');
     end;




  Form1.LastNum.Open;
  Form1.LastNum.IndexName:='MX';
  Form1.LastNum.SetKey;
  Form1.LastNumMX.AsString:=Mx;
  Form1.LastNumKOD.AsString:='N_ZAJ';
  if Form1.LastNum.GoToKey=False then
     Form1.LastNum.AppendRecord([Mx,'N_ZAJ',FormatDateTime('yy',date)+MX+'J'+'00000','Номер заявки']);

  if StrToInt(FormatDateTime('yy',date))>StrToInt(Copy(Form1.LastNumPEREM.AsString,1,2)) then
     begin
     Form1.Lastnum.edit;
     Form1.LastNumPEREM.AsString:=FormatDateTime('yy',date)+MX+'J'+'00000';
     Form1.LastNum.Edit;
     Form1.LastNum.Post;
     end;

  if (StrToInt(FormatDateTime('yy',date))<StrToInt(Copy(Form1.LastNumPEREM.AsString,1,2))) then
     begin
     my_messageTime('Внимание','На Вашем компьютере установлена неправильная дата.'+#13+'Так работать нельзя!', clYellow,20000);
     exit;
     end;


   // Заносим в ZAJ_O
   InsertForm.ZajQuery.SQL.Clear;
   InsertForm.ZajQuery.SQL.Add('INSERT INTO "'+Path_base+'zakaz\20'+FormatDateTime('yy',date)+'\zaj_o.dbf" '+
                               '(IO,N_DOK,K_KL,KLIENT,D_O,KURS_D,V_UE,K_SKID_D,SKIDKA_D,SO,CAR,N_CAR,K_SKID_R,SKIDKA_R,C_NCH,NSP,NDS,OKRUGL,TIME_O,MANAGER,KOD_SUBD,OTDEL,CAR_N_ZAP,HIDDEN,RUB_CL) VALUES '+
                               '(:IO,:N_DOK,:K_KL,:KLIENT,:D_O,:KURS_D,:V_UE,:K_SKID_D,:SKIDKA_D,:SO,:CAR,:N_CAR,:K_SKID_R,:SKIDKA_R,:C_NCH,:NSP,:NDS,:OKRUGL,:TIM,:MANAGER,:KOD_SUBD,:OTDEL,:CAR_N_ZAP,:HIDDEN,:RUB_CL)');



   InsertForm.ue.Open;
   InsertForm.ue.IndexName:='Mx';
   InsertForm.ue.SetRange(Mx,Mx+'9');
   InsertForm.ue.Last;


  try
  ccf_za_nal:=StrToFloat(Float_point(ccfForm.MaskEdit9.Text));
  except
  ccf_za_nal:=0.;
  end;

  try
  ccf_za_b_nal:=StrToFloat(Float_point(ccfForm.MaskEdit10.Text));
  except
  ccf_za_b_nal:=0.;
  end;


 if Copy(k_kl,2,1)='F' then
    begin
    Form1.cl_f.Open;
    Form1.cl_f.IndexNAme:='K_KL';
    if Form1.cl_f.FindKey([k_kl]) then
       begin
       InsertForm.ZajQuery.Params.ParamByName('KLIENT').AsString:=UncryptString(Form1.cl_fNAME.AsString,key1c,key2c);
       InsertForm.ZajQuery.Params.ParamByName('RUB_CL').AsString:=Form1.cl_fRUB_CL.AsString;

       InsertForm.ZajQuery.Params.ParamByName('K_SKID_D').AsString:=FloatToStr(Form1.cl_FK_N_SKID_Z.AsFloat);
       InsertForm.ZajQuery.Params.ParamByName('SKIDKA_D').AsFloat:=Form1.cl_FSKIDKA_Z.AsFloat;
       InsertForm.ZajQuery.Params.ParamByName('K_SKID_R').AsString:=FloatToStr(Form1.cl_FK_N_SKID_R.AsFloat);
       InsertForm.ZajQuery.Params.ParamByName('SKIDKA_R').AsFloat:=Form1.cl_FSKIDKA_R.AsFloat;
       InsertForm.ZajQuery.Params.ParamByName('C_NCH').AsFloat:=Form1.cl_FC_NCH.AsFloat;


       InsertForm.ZajQuery.Params.ParamByName('SO').AsString:='01';
       InsertForm.ZajQuery.Params.ParamByName('KURS_D').Value:=InsertForm.ueCURS.AsFloat+(InsertForm.ueCURS.AsFloat/100)*ccf_za_nal;

       InsertForm.ZajQuery.Params.ParamByName('MANAGER').AsString:=InsertForm.cl_FMANAGER.AsString;

       end
    else
       begin
       my_messageTime('Внимание!','Клиент не найден в справочнике.',clYellow,3000);
       exit;
       end;


    end
else
 if Copy(k_kl,2,1)='U' then
    begin
    Form1.cl_u.Open;
    Form1.cl_u.IndexNAme:='K_KL';
    if Form1.cl_u.FindKey([k_kl]) then
       begin
       InsertForm.ZajQuery.Params.ParamByName('KLIENT').AsString:=UncryptString(Form1.cl_uFS.AsString,key1c,key2c)+' '+UncryptString(Form1.cl_uORG.AsString,key1c,key2c);
       InsertForm.ZajQuery.Params.ParamByName('RUB_CL').AsString:=Form1.cl_uRUB_CL.AsString;
       InsertForm.ZajQuery.Params.ParamByName('K_SKID_D').AsString:=FloatToStr(Form1.Cl_uK_N_SKID_Z.AsFloat);
       InsertForm.ZajQuery.Params.ParamByName('SKIDKA_D').AsFloat:=Form1.Cl_uSKIDKA_Z.AsFloat;
       InsertForm.ZajQuery.Params.ParamByName('K_SKID_R').AsString:=FloatToStr(Form1.Cl_uK_N_SKID_R.AsFloat);
       InsertForm.ZajQuery.Params.ParamByName('SKIDKA_R').AsFloat:=Form1.Cl_uSKIDKA_R.AsFloat;
       InsertForm.ZajQuery.Params.ParamByName('C_NCH').AsFloat:=Form1.Cl_uC_NCH.AsFloat;

       InsertForm.ZajQuery.Params.ParamByName('SO').AsString:='01';
       InsertForm.ZajQuery.Params.ParamByName('KURS_D').Value:=InsertForm.ueCURS.AsFloat+(InsertForm.ueCURS.AsFloat/100)*ccf_za_nal;

       InsertForm.ZajQuery.Params.ParamByName('MANAGER').AsString:=Form1.Cl_uMANAGER.AsString;


       end
    else
       begin
       my_messageTime('Внимание!','Клиент не найден в справочнике.',clYellow,3000);
       exit;
       end;

    end
 else
     begin
     my_messageTime('Внимание!','Клиент не найден в справочнике.',clYellow,3000);
     exit;
     end;




   Form1.Lastnum.edit;
   Form1.LastNumPEREM.AsString:=FormatDateTime('yy',date)+MX+'J'+New_number(Copy(Form1.LastNumPEREM.AsString,5,5));
   Form1.LastNum.Edit;
   Form1.LastNum.Post;

   n_dok:=FormatDateTime('yy',date)+MX+'J'+Copy(Form1.LastNumPEREM.AsString,5,5);



   if ZajavkiForm.MaskEdit9.Text<>FormatDateTime('yyyy',date) then
      begin
      ZajavkiForm.MaskEdit9.Text:=FormatDateTime('yyyy',date);
      ZajavkiForm.PAgeControl1.ActivePage:=ZajavkiForm.Tab1;
      Zajavki.Change_Year;
      end;

   ZajavkiForm.Zaj_oTable.Open;
   ZajavkiForm.Zaj_oTable.IndexName:='N_DOK';
   if ZajavkiForm.Zaj_oTable.FindKey([n_dok])=True then
      begin
      Form1.Lastnum.edit;
      Form1.LastNumPEREM.AsString:=FormatDateTime('yy',date)+MX+'J'+New_number(Copy(Form1.LastNumPEREM.AsString,5,5));
      Form1.LastNum.Edit;
      Form1.LastNum.Post;

      n_dok:=FormatDateTime('yy',date)+MX+'J'+Copy(Form1.LastNumPEREM.AsString,5,5);
      end;


   ZajavkiForm.Zaj_oTable.Open;
   ZajavkiForm.Zaj_oTable.IndexName:='N_DOK';
   if ZajavkiForm.Zaj_oTable.FindKey([n_dok])=True then
      begin
      my_messageTime('Внимание!',PChar('Запрос отклонен! Документ с номером '+n_dok+' уже есть. Сообщите старшему администратору.'),clYellow,30000);
      exit;
      end;



   if Install.ComboBox3.ItemIndex=1 then InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=1
   else
   if Install.ComboBox3.ItemIndex=0 then InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=2
   else
   if Install.ComboBox3.ItemIndex=2 then InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=0
   else
   if Install.ComboBox3.ItemIndex=3 then InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=5 //до десятков рублей
   else
   if Install.ComboBox3.ItemIndex=4 then InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=6 //до сотен рублей
   else InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat:=2;

   try
   InsertForm.ZajQuery.ParamByName('NSP').AsFloat:=StrToFloat(Float_point(ccfForm.MaskEdit12.Text));
   except
    InsertForm.ZajQuery.ParamByName('NSP').AsFloat:=0;
   end;

   try
   InsertForm.ZajQuery.ParamByName('NDS').AsFloat:=StrToFloat(Float_point(ccfForm.MaskEdit4.Text));
   except
    InsertForm.ZajQuery.ParamByName('NDS').AsFloat:=20;
   end;

   InsertForm.ZajQuery.Params.ParamByName('V_UE').AsString:=Form1.Valuta_ue.Caption;

   InsertForm.Cars.Open;
   InsertForm.Cars.IndexName:='N_ZAP';
   InsertForm.Cars.SetKey;
   InsertForm.CarsK_KL.AsString:=k_kl;
   InsertForm.CarsN_ZAP.AsString:=car_N_ZAP;
   if InsertForm.Cars.GoToKey then
      begin
      InsertForm.ZajQuery.Params.ParamByName('CAR').AsString:=InsertForm.CarsAUTO.AsString;
      InsertForm.ZajQuery.Params.ParamByName('N_CAR').AsString:=UnCryptString(InsertForm.CarsGOS_N.AsString,key1a,key2a);
      end
   else
      begin
      my_messageTime('Внмание!','Присвойте автомобиль клиенту в справочнике',clYellow,30000);
      exit;
      end;


   try
   hist_o_tableName_proc(InsertForm.hist_o,n_dok,'20'+Copy(n_dok,1,2),FormatDateTime('MM',Date));
   except
   end;

   try
   InsertForm.hist_o.Open;
   InsertForm.hist_o.AppendRecord([kod_operatora,
                                   Date,
                                   TimeToStr(Time),
                                   n_dok,
                                   NULL,
                                   k_kl,
                                   InsertForm.ZajQuery.Params.ParamByName('KLIENT').AsString,
                                   InsertForm.ZajQuery.Params.ParamByName('CAR').AsString,
                                   InsertForm.ZajQuery.Params.ParamByName('N_CAR').AsString,
                                   Date,
                                   NULL,
                                   NULL,
                                   NULL,
                                   InsertForm.ZajQuery.Params.ParamByName('SO').AsString,
                                   NULL,
                                   InsertForm.ZajQuery.Params.ParamByName('KURS_D').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('V_UE').AsString,
                                   InsertForm.ZajQuery.Params.ParamByName('K_SKID_D').AsString,
                                   InsertForm.ZajQuery.Params.ParamByName('SKIDKA_D').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('K_SKID_R').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('SKIDKA_R').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('C_NCH').AsFloat,
                                   NULL,NULL,NULL,NULL,NULL,NULL,
                                   InsertForm.ZajQuery.Params.ParamByName('NSP').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('NDS').AsFloat,
                                   InsertForm.ZajQuery.Params.ParamByName('OKRUGL').AsString]);
   except
   end;




   ZajavkiForm.Zaj_oTable_for_add.Open;
   ZajavkiForm.Zaj_oTable_for_add.AppendRecord([kod_operatora,
                                                n_dok,
                                                NULL, //variant
                                                k_kl,
                                                InsertForm.ZajQuery.Params.ParamByName('KLIENT').AsString,
                                                InsertForm.ZajQuery.Params.ParamByName('CAR').AsString,
                                                InsertForm.ZajQuery.Params.ParamByName('N_CAR').AsString,
                                                date,
                                                NULL, //d_opl
                                                NULL, //d_otsr
                                                NULL, //SUM_PR
                                                InsertForm.ZajQuery.Params.ParamByName('SO').AsString,
                                                NULL,  //sost
                                                InsertForm.ZajQuery.Params.ParamByName('KURS_D').AsFloat,
                                                InsertForm.ZajQuery.ParamByName('V_UE').AsString,
                                                InsertForm.ZajQuery.Params.ParamByName('K_SKID_D').AsString,
                                                InsertForm.ZajQuery.Params.ParamByName('SKIDKA_D').AsFloat,
                                                InsertForm.ZajQuery.Params.ParamByName('K_SKID_R').AsString,
                                                InsertForm.ZajQuery.Params.ParamByName('SKIDKA_R').AsFloat,
                                                InsertForm.ZajQuery.Params.ParamByName('C_NCH').AsFloat,
                                                NULL,//pr_pech
                                                NULL,//block
                                                NULL, //n_recl
                                                NULL, //pr_otvet
                                                NULL, //prober
                                                NULL,//comm1
                                                NULL,//comm2
                                                NULL,//kod_subd
                                                InsertForm.ZajQuery.ParamByName('NSP').AsFloat,
                                                InsertForm.ZajQuery.ParamByName('NDS').AsFloat,
                                                InsertForm.ZajQuery.ParamByName('OKRUGL').AsFloat,   //S_OKRUGL - N -1
                                                NULL,//N_fac
                                                NULL,//D_PR
                                                FormatDateTime('hh:nn',now),
                                                NULL,//car_in
                                                NULL,//hidden
                                                NULL,//d_end
                                                NULL,//t_end
                                                NULL,// k_plat
                                                InsertForm.ZajQuery.Params.ParamByName('MANAGER').AsString, //manager
                                                NULL, //otdel
                                                NULL, //мастер откр
                                                NULL, //мастер закр
                                                car_n_zap,
                                                Date,
                                                FormatDateTime('hh:nn',now),
                                                NULL,
                                                NULL,
                                                NULL,
                                                NULL,
                                                NULL,
                                                NULL,
                                                NULL,//OKRUGL - c -1
                                                NULL, //inc
                                                InsertForm.ZajQuery.Params.ParamByName('RUB_CL').AsString]);




   ZajavkiForm.Zaj_oTable_for_add.Close;


   try
   Table.Edit;
   Table.FieldByName('N_ZAJ').AsString:=n_dok;
   Table.Edit;
   Table.Post;
   except
   end;

   ZajavkiForm.Show;
   if ZajavkiForm.Zaj_oQuery.Locate('N_DOK',n_dok,[])=True then
      ZajavkiForm.DBGrid2DBLClick(ZajavkiForm.DBGrid2)
   else
      my_messageTime('Внимание!','Документ не найден.',clYellow,3000);


   AppointmentForm.Close;
end;



procedure create_d_offTable(yyyy: string);
begin
with AppointmentForm do
begin
  if FileExists(path_base+'HIST\'+yyyy+'\D_offTable.dbf')=False then
     begin
     d_offTableSTO.Close;
     d_offTableForCreate.Close;
     d_offTableForCreate.TableName:=path_base+'HIST\'+yyyy+'\D_offTable.dbf';
     d_offTableForCreate.IndexFiles.Clear;
     d_offTableForCreate.CreateTable;

     d_offTableForCreate.Close;

     d_offTableSTO.Close;
     d_offTableSTO.TableName:=path_base+'HIST\'+yyyy+'\D_offTable.dbf';
     d_offTableSTO.IndexFiles.Clear;

     d_offTableSTO.Open;
     d_offTableSTO.IndexOn(path_base+'HIST\'+yyyy+'\d_offTable.cdx','KOD','KOD+DTOS(D1)','',Duplicates,Ascending);
     d_offTableSTO.Reindex;
     d_offTableSTO.Close;
     end;
end;
end;




//показываем иконку окна в трее
procedure TAppointmentForm.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);

  if Install.ComboBox13.ItemIndex=1 then
     begin

     with Params do
          begin
          Params.ExStyle := ExStyle or WS_EX_APPWINDOW;
          end;
     end;
end;




procedure CreateMonthTableProc(manth,_year: integer);
var
  i, count_month_day, day, first, j: integer;
  first_day_of_week, manth_number, year: string;
begin
with appointmentForm do
  begin

  MonthTable.Close;
  MonthTable.TableName:=tmp_path+'MonthTable.dbf';
  MonthTable.CreateTable;

  MonthTable.Open;
  for i:=1 to 6 do
      MonthTable.AppendRecord([]);

  //определяем день недели первого числа выбранного месяца
  manth_number:=IntToStr(manth);
  if Length(manth_number)=1 then manth_number:='0'+manth_number;
  year:=IntToStr(_year);


                                  //Получаем дату в формате YYYYMMDD
  first_day_of_week:=DayOfTheWeek(year+manth_number+'01');



  count_month_day:=DaysInAMonth(_year,manth);


  if first_day_of_week='Понедельник' then first:=1;
  if first_day_of_week='Вторник' then first:=2;
  if first_day_of_week='Среда' then first:=3;
  if first_day_of_week='Четверг' then first:=4;
  if first_day_of_week='Пятница' then first:=5;
  if first_day_of_week='Суббота' then first:=6;
  if first_day_of_week='Воскресенье' then first:=7;


  MonthTable.First;
  day:=1;

  for i:=first to 7 do
         begin
         MonthTable.Edit;
         MonthTable.FieldByName('D'+IntToStr(i)).AsString:=IntToStr(day);
         MonthTable.Edit;
         MonthTable.Post;

         inc(day);
         end;


   MonthTable.Next;
   for j:=1 to 5 do
       begin
       for i:=1 to 7 do
           begin
           MonthTable.Edit;
           MonthTable.FieldByName('D'+IntToStr(i)).AsString:=IntToStr(day);
           MonthTable.Edit;
           MonthTable.Post;

           inc(day);
           if day>count_month_day then break;
           end;

       if day>count_month_day then break;
       MonthTable.Next;
       end;

end;
end;




procedure CreateMonthTableProcPosts(manth,_year: integer);
var
  i, count_month_day, day, first, j: integer;
  first_day_of_week, manth_number, year, str: string;
begin
with appointmentForm do
begin


 MonthTablePosts.Close;
 MonthTablePosts.TableName:=tmp_path+'postsMonthTable.dbf';
 MonthTablePosts.CreateTable;

 MonthTablePosts.Open;

  Form1.Spr00.Open;
  Form1.Spr00.IndexName:='NAIM';
  Form1.Spr00.SetRange('30'+'','30'+'яя');
  Form1.Spr00.First;
  while Form1.Spr00.Eof=FAlse do
        begin
        MonthTablePosts.AppendRecord([Form1.Spr00KOD.ASString]);
        Form1.Spr00.Next;
        end;



 //определяем день недели первого числа выбранного месяца
  manth_number:=IntToStr(manth);
  if Length(manth_number)=1 then manth_number:='0'+manth_number;
  year:=IntToStr(_year);



                                  //Получаем дату в формате YYYYMMDD
  first_day_of_week:=DayOfTheWeek(year+manth_number+'01');



  count_month_day:=DaysInAMonth(_year,manth);


  MonthTablePostsD29.Visible:=True;
  MonthTablePostsD30.Visible:=True;
  MonthTablePostsD31.Visible:=True;


  if count_month_day=28 then
     begin
     MonthTablePostsD29.Visible:=False;
     MonthTablePostsD30.Visible:=False;
     MonthTablePostsD31.Visible:=False;
     end
  else
  if count_month_day=30 then
     MonthTablePostsD31.Visible:=False;



  if first_day_of_week='Понедельник' then first:=1;
  if first_day_of_week='Вторник' then first:=3;
  if first_day_of_week='Среда' then first:=5;
  if first_day_of_week='Четверг' then first:=7;
  if first_day_of_week='Пятница' then first:=9;
  if first_day_of_week='Суббота' then first:=11;
  if first_day_of_week='Воскресенье' then first:=13;

  str:='ПнВтСрЧтПтСбВс';


  MonthTablePosts.First;
  day:=first;

  for i:=1 to MonthTablePosts.Fields.Count-1 do
      begin

      MonthTablePosts.Fields[i].DisplayLabel:=Copy(str,day,2)+' '+Copy(MonthTablePosts.Fields[i].FieldName,2,10);

      day:=day+2;
      if day>13 then day:=1;
      end;



end;
end;



procedure CreateMonthTableProcStaff(manth,_year: integer);
var
  i, count_month_day, day, first, j: integer;
  first_day_of_week, manth_number, year, str: string;
begin
with appointmentForm do
begin


 MonthTableStaff.Close;
 MonthTableStaff.TableName:=tmp_path+'StaffMonthTable.dbf';
 MonthTableStaff.CreateTable;

 MonthTableStaff.Open;

 if Tab_nTable.TAbleName<>path_base+'\sprav\tab_n.dbf' then
    begin
    Tab_nTable.Close;
    Tab_nTable.TAbleName:=path_base+'\sprav\tab_n.dbf';
    Tab_nTable.IndexFiles.Clear;
    Tab_nTable.IndexFiles.Add(path_base+'\sprav\tab_n.cdx');
    end;




  Form1.Spr00.Open;
  Form1.Spr00.IndexName:='NAIM';
  Form1.Spr00.SetRange('32'+'','32'+'яя');
  Form1.Spr00.First;
  while Form1.Spr00.Eof=FAlse do
        begin
        if Form1.Spr00PR.ASString='M' then
           MonthTableStaff.AppendRecord(['MM'+Form1.Spr00KOD.ASString]);

        Form1.Spr00.Next;
        end;


  SqlQuery1.SQL.Clear;
  SqlQuery1.SQL.Add('SELECT DISTINCT TAB_N,FIO FROM "'+tab_nTable.TableName+'"');
  SqlQuery1.Open;
  SqlQuery1.First;



  While SqlQuery1.Eof=False do
        begin
        if Copy(SqlQuery1.FieldByName('FIO').AsString,1,1)<>'*' then
           MonthTableStaff.AppendRecord(['ST'+SqlQuery1.FieldByName('TAB_N').AsString]);

        SqlQuery1.Next;
        end;





 //определяем день недели первого числа выбранного месяца
  manth_number:=IntToStr(manth);
  if Length(manth_number)=1 then manth_number:='0'+manth_number;
  year:=IntToStr(_year);



                                  //Получаем дату в формате YYYYMMDD
  first_day_of_week:=DayOfTheWeek(year+manth_number+'01');



  count_month_day:=DaysInAMonth(_year,manth);


  MonthTableStaffD29.Visible:=True;
  MonthTableStaffD30.Visible:=True;
  MonthTableStaffD31.Visible:=True;


  if count_month_day=28 then
     begin
     MonthTableStaffD29.Visible:=False;
     MonthTableStaffD30.Visible:=False;
     MonthTableStaffD31.Visible:=False;
     end
  else
  if count_month_day=30 then
     MonthTableStaffD31.Visible:=False;



  if first_day_of_week='Понедельник' then first:=1;
  if first_day_of_week='Вторник' then first:=3;
  if first_day_of_week='Среда' then first:=5;
  if first_day_of_week='Четверг' then first:=7;
  if first_day_of_week='Пятница' then first:=9;
  if first_day_of_week='Суббота' then first:=11;
  if first_day_of_week='Воскресенье' then first:=13;

  str:='ПнВтСрЧтПтСбВс';


  MonthTableStaff.First;
  day:=first;

  for i:=1 to MonthTableStaff.Fields.Count-1 do
      begin

      MonthTableStaff.Fields[i].DisplayLabel:=Copy(str,day,2)+' '+Copy(MonthTableStaff.Fields[i].FieldName,2,10);

      day:=day+2;
      if day>13 then day:=1;
      end;



end;
end;








procedure fill_sto_days_off(dsCalendar: TDSCAlendar; kod: string);
var
 _date: TDate;
 _year: string;
begin
with AppointmentForm do
begin

     if StrToInt(FormatDAteTime('YYYY',dsCAlendar.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',dsCAlendar.DAte);


   if d_offTableSTO.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
      begin
      d_offTableSTO.Close;
      d_offTableSTO.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
      d_offTableSTO.IndexFiles.Clear;
      d_offTableSTO.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');
      end;

    d_offTableSTO.Open;
    d_offTableSTO.IndexName:='KOD';
    d_offTableSTO.SetRange(Copy(kod+'      ',1,6)+'',Copy(kod+'      ',1,6)+'яя');

    dsCAlendar.Holidays.Clear;
    d_offTableSTO.First;
    while d_offTableSTO.Eof=False do
          begin
          dsCAlendar.Holidays.Add(StrToInt(FormatDAteTime('dd',d_offTableSTOD1.VAlue)),StrToInt(FormatDAteTime('mm',d_offTableSTOD1.VAlue)));

          d_offTableSTO.Next;
          end;


    dsCalendar.Repaint;
end;
end;


procedure erace_fields;
begin
with AppointmentForm do
     begin
     StoPostComboBox.Visible:=FAlse;

     addRecordCheckBox.Visible:=FAlse;
     addRecordCheckBox.Checked:=FAlse;

     ConfirmSMSCheckBox.Checked:=False;

     BitBtn15.Visible:=False;
     CreateOrderBitBtn2.Enabled:=False;
     CreateOrderBitBtn.Enabled:=False;

     CardNumberMaskEdit.Text:='';

     OrderNumberMaskEdit.Text:='';
     StartTimeMaskEdit.Text:='';
     EndTimeMaskEdit.Text:='';
     CarNumberMaskEdit.Text:='';
     ClientNameMaskEdit.Text:='';
     MarkaMaskEdit.Text:='';
     ModelMaskEdit.Text:='';
     telCountryCodeMaskEdit.Text:='+7';
     telCodeMaskEdit.Text:='';
     telNumberMaskEdit.Text:='';
     ReasonMemo.Text:='';
     CommentMemo.Text:='';

     Rec_idLAbel.CAption:='';
     Car_n_zapLAbel.CAption:='';
     K_KLLAbel.CAption:='';
     ClientSpravBitBtn.Visible:=FAlse;

     DAteTimePicker1.Visible:=FAlse;
     LAbel26.Visible:=FAlse;

     end;

end;




procedure TAppointmentForm.FormCreate(Sender: TObject);
begin
   RezhimLabel.Caption:='';

   FormForSendSMS_addRec:=False;
   FormForSendSMS_editRec:=False;

   Key1a:='y';
   Key2a:='7';

   AppointmentForm.Icon:=Form1.Icon;
   Global_X:=0; Global_Y:=0; LastPaint_X:=0; LastPaint_Y:=0; FieldNAIM_Righ_X:=0;
   Global_ColumnName:='';
   RealColumnNAme:='';


   PAnel9.Left:=ClientSearchBitBtn.Left;

   DAteTimePicker2.Date:=date;
   DAteTimePicker3.Date:=date;

   YearComboBox.Items.Clear;
   YearComboBox.Items.Add(FormatDateTime('YYYY',date));
   YearComboBox.Items.Add(IntToStr(StrToInt(FormatDateTime('YYYY',date))+1));
   YearComboBox.ItemIndex:=0;

end;


procedure zapolnenie_kl_proc;
var
  tel: string;
begin
with AppointmentForm do
begin

       if (Copy(k_klLAbel.CAption,2,1)='F') then
          begin
          Form1.cl_f.Open;
          Form1.cl_f.IndexName:='K_KL';
          if Form1.cl_f.FindKey([k_klLAbel.CAption]) then
             begin
             ClientNameMaskEdit.Text:=UnCryptString(Form1.cl_fNAME.AsString,key1c,key2c);

             tel:=Remove_tire(UnCryptString(Form1.cl_fTEL3.AsString,key1c,key2c));

             if Length(tel)=7 then tel:='+7000'+tel;

             if (Copy(tel,1,1)<>'+') and (tel<>'') then tel:='+'+tel;
             if (Copy(tel,1,2)='+8')  then tel:='+7'+Copy(tel,3,20);
             if (Copy(tel,1,2)='+9')  then tel:='+7'+Copy(tel,2,20);
             if (Copy(tel,1,2)='+(')  then tel:='+7'+Copy(tel,2,20);


             if Pos('(',tel)=0 then
                begin
                tel:=Copy(tel,1,2)+'('+Copy(tel,3,3)+')'+Copy(tel,6,10);
                end;


             telCountryCodeMaskEdit.Text:='';
             telCodeMaskEdit.Text:='';
             telNumberMaskEdit.Text:='';

             if Pos('(',tel)>0 then
                Copy(telCountryCodeMaskEdit.Text,1,Pos('(',tel)-1)
             else
                telCountryCodeMaskEdit.Text:='+7';

             if Pos(')',tel)>0 then
                telCodeMaskEdit.Text:=Copy(tel,Pos('(',tel)+1,Pos(')',tel)-4);

             if Pos(')',tel)>0 then
                telNumberMaskEdit.Text:=Remove_tire(Copy(tel,Pos(')',tel)+1,100));


             end;
          end
       else
       if Copy(k_klLAbel.CAption,2,1)='U' then
          begin
          Form1.cl_u.Open;
          Form1.cl_u.IndexName:='K_KL';
          if Form1.cl_u.FindKey([k_klLAbel.CAption]) then
             begin
             ClientNameMaskEdit.Text:=UnCryptString(Form1.cl_uFS.AsString,key1c,key2c)+' '+UnCryptString(Form1.cl_uORG.AsString,key1c,key2c);

             tel:=Remove_tire(UnCryptString(Form1.cl_uTEL.AsString,key1c,key2c));

             if Length(tel)=7 then tel:='+7000'+tel;

             if (Copy(tel,1,1)<>'+') and (tel<>'') then tel:='+'+tel;
             if (Copy(tel,1,2)='+8')  then tel:='+7'+Copy(tel,3,20);
             if (Copy(tel,1,2)='+9')  then tel:='+7'+Copy(tel,2,20);
             if (Copy(tel,1,2)='+(')  then tel:='+7'+Copy(tel,2,20);


             if Pos('(',tel)=0 then
                begin
                tel:=Copy(tel,1,2)+'('+Copy(tel,3,3)+')'+Copy(tel,6,10);
                end;


             telCountryCodeMaskEdit.Text:='';
             telCodeMaskEdit.Text:='';
             telNumberMaskEdit.Text:='';

             if Pos('(',tel)>0 then
                telCountryCodeMaskEdit.Text:=Copy(tel,1,Pos('(',tel)-1)
             else
                telCountryCodeMaskEdit.Text:='+7';

             if Pos(')',tel)>0 then
                telCodeMaskEdit.Text:=Copy(tel,Pos('(',tel)+1,Pos(')',tel)-4);

             if Pos(')',tel)>0 then
                telNumberMaskEdit.Text:=Remove_tire(Copy(tel,Pos(')',tel)+1,100));

             end;
          end
end;
end;



procedure car_from_sprav_proc;
var
   str: string;
   finded: Boolean;
   i: integer;
begin
with AppointmentForm do
begin



     if Trim(k_klLAbel.CAption)<>'' then
        begin
        carsQuery.SQL.Clear;
        carsQuery.SQL.Add('SELECT * FROM "'+path_base+'sprav\cars.dbf" WHERE (K_KL LIKE :K_KL) ORDER BY 2,10');
        carsQuery.ParamByName('K_KL').AsString:=k_klLAbel.CAption;

        carsQuery.Open;
        carsQuery.First;

        finded:=False;

        if (carsQuery.RecordCount>1) and (trim(CarNumberMaskEdit.Text)<>'') then
           begin

           while carsQuery.Eof=False do
                 begin
                 if Pos(CarNumberMaskEdit.Text,UnCryptString(CarsQUERY.FieldByName('GOS_N').AsString,key1a,key2a))>0 then
                    begin
                    finded:=True;
                    break;
                    end;

                 carsQuery.Next;
                 end;
          end;



        if (carsQuery.RecordCount>1) and (finded=False) then
           begin  //выбор машины

           ClientCarsListBox.Items.Clear;
           carsQuery.First;

           while carsQuery.Eof=FAlse do
                 begin
                 str:=Copy(CarsQUERY.FieldByName('N_ZAP').AsString+'          ',1,10);

                 Form1.Spr01.Open;
                 Form1.Spr01.IndexName:='KOD';
                 Form1.Spr01.SetKey;
                 Form1.Spr01GR.AsString:='11';
                 Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,2);
                 if Form1.Spr01.GoToKey then
                    str:=str+Copy(Trim(Form1.Spr01NAIM.AsString)+'                    ',1,20);

                 Form1.Spr01.SetKey;
                 Form1.Spr01GR.AsString:='12';
                 Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,4);
                 if Form1.Spr01.GoToKey then
                    str:=str+Copy(Trim(Form1.Spr01NAIM.AsString)+'                    ',1,20);

                 str:=str+' Гос.№: '+UnCryptString(CarsQUERY.FieldByName('GOS_N').AsString,key1a,key2a);

                 ClientCarsListBox.Items.Add(str);

                 carsQuery.Next;
                 end;

           PAnel10.Caption:='Выберите автомобиль';
           PAnel9.Visible:=True;
           end
        else
           begin //заполняем
           CarNumberMaskEdit.Text:=UnCryptString(CarsQUERY.FieldByName('GOS_N').AsString,key1a,key2a);

           CAr_n_zapLAbel.Caption:=CarsQUERY.FieldByName('N_ZAP').AsString;

           if trim(car_n_zapLAbel.CAption)<>'' then
              begin
              BitBtn15.Visible:=True;
              BitBtn15.Enabled:=True;
              end;

           Form1.Spr01.Open;
           Form1.Spr01.IndexName:='KOD';
           Form1.Spr01.SetKey;
           Form1.Spr01GR.AsString:='11';
           Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,2);
           if Form1.Spr01.GoToKey then
              MarkaMaskEdit.Text:=Trim(Form1.Spr01NAIM.AsString);

           Form1.Spr01.SetKey;
           Form1.Spr01GR.AsString:='12';
           Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,4);
           if Form1.Spr01.GoToKey then
              ModelMaskEdit.Text:=Trim(Form1.Spr01NAIM.AsString);


           Form1.Spr01.SetKey;
           Form1.Spr01GR.AsString:='13';
           Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,6);
           if Form1.Spr01.GoToKey then
              ModelMaskEdit.Text:=TrimRight(ModelMaskEdit.Text)+', '+Trim(Form1.Spr01NAIM.AsString);


           i:=Pos('  ',ModelMaskEdit.Text);
           while i>0 do
                 begin
                 ModelMaskEdit.Text:=Copy(ModelMaskEdit.Text,1,i-1)+' '+Copy(ModelMaskEdit.Text,i+2,100);
                 i:=Pos('  ',ModelMaskEdit.Text);
                 end;

           end;

        end;

end;

end;



procedure find_by_tel_proc;
var
  rec_id: string;
begin

with AppointmentForm do
     begin
     if trim(TelMaskEditNumber.Text)='' then exit;

     SqlQuery1.SQL.Clear;
     SqlQuery1.SQL.Add('SELECT * FROM "'+AppTAble.TAbleNAme+'" WHERE TEL LIKE :TEL');
     if trim(TelMaskEditCode.Text)<>'' then
        SqlQuery1.ParamByNAme('TEL').ASString:='%'+CryptString(TelMaskEditCountryCode.Text+TelMaskEditCode.Text+TelMaskEditNumber.Text, key1c, key2c)+'%'
     else
        SqlQuery1.ParamByNAme('TEL').ASString:='%'+CryptString(TelMaskEditNumber.Text, key1c, key2c)+'%';

     SqlQuery1.Open;
     SqlQuery1.LAst;

     rec_id:=SqlQuery1.FieldByNAme('REC_ID').AsString;


     if SqlQuery1.FieldByNAme('DATE_APP').AsString<>'' then
        begin
        if my_dlg('Внимание!',PChar('Для номера: '+TelMaskEditCountryCode.Text+'('+TelMaskEditCode.Text+')'+TelMaskEditNumber.Text+#13+'найдено записей: '+IntToStr(SqlQuery1.RecordCount)+#13+'Последняя запись на '+SqlQuery1.FieldByNAme('DATE_APP').AsString+' - '+SqlQuery1.FieldByNAme('FIO').AsString+#13+'Показать эту запись?'),clYellow)=True then
           begin
           dsCAlendar1.Date:=SqlQuery1.FieldByNAme('DATE_APP').Value;
           dsCalendar1DateChange(dsCAlendar1,SqlQuery1.FieldByNAme('DATE_APP').Value-1,SqlQuery1.FieldByNAme('DATE_APP').Value);
           PageControl1.ActivePAge:=TAb2;
           PageControl1Change(PageControl1);
           AppTAble.Locate('REC_ID',rec_id,[]);
           end;



        end
     else
        my_messageTime('Внимание!','Ничего не найдено.',clYellow,3000);

     end;






end;


procedure find_by_gosn_proc;
var
  rec_id: string;
begin

with AppointmentForm do
     begin
     if trim(CarNumberFindEdit.Text)='' then exit;

     SqlQuery1.SQL.Clear;
     SqlQuery1.SQL.Add('SELECT * FROM "'+AppTAble.TAbleNAme+'" WHERE GOS_N LIKE :GOS_N');
     SqlQuery1.ParamByNAme('GOS_N').ASString:='%'+CryptString(CarNumberFindEdit.Text,key1c,key2c)+'%';
     SqlQuery1.Open;
     SqlQuery1.LAst;

     rec_id:=SqlQuery1.FieldByNAme('REC_ID').AsString;

     if SqlQuery1.FieldByNAme('DATE_APP').AsString<>'' then
        begin
        if my_dlg('Внимание!',PChar('Для номера: '+CarNumberFindEdit.Text+#13+'найдено записей: '+IntToStr(SqlQuery1.RecordCount)+#13+'Последняя запись на '+SqlQuery1.FieldByNAme('DATE_APP').AsString+' - '+SqlQuery1.FieldByNAme('FIO').AsString+#13+'Показать эту запись?'),clYellow)=True then
           begin
           dsCAlendar1.Date:=SqlQuery1.FieldByNAme('DATE_APP').Value;
           dsCalendar1DateChange(dsCAlendar1,SqlQuery1.FieldByNAme('DATE_APP').Value-1,SqlQuery1.FieldByNAme('DATE_APP').Value);
           PageControl1.ActivePAge:=TAb2;
           PageControl1Change(PageControl1);
           AppTAble.Locate('REC_ID',rec_id,[]);
           end;



        end
     else
        my_messageTime('Внимание!','Ничего не найдено.',clYellow,3000);

     end;


end;



procedure fill_form(Table: THalcyonDataSet); //zapolnenie
var
  tel: string;
begin
with AppointmentForm do
begin

     LAbel13.Caption:='Запись на '+FormatDAteTime('MMMM DD',Table.FieldByName('DATE_APP').Value)+'-е, '+DayOfTheWeek(FormatDateTime('YYYYMMDD',Table.FieldByName('DATE_APP').VAlue));

     StartTimeMaskEdit.ReadOnly:=FAlse;

     CardNumberMaskEdit.Text:='';

     if Form1.SpeedButton3.Visible=True then
        CreateOrderBitBtn2.Enabled:=True
     else
        CreateOrderBitBtn2.Enabled:=False;


     addRecordCheckBox.Checked:=FAlse;


     Form1.Spr00.Open;
     Form1.Spr00.IndexName:='KOD';
     Form1.Spr00.SetRange('','');
     Form1.Spr00.SetKey;
     Form1.Spr00GR.AsString:='30';
     Form1.Spr00KOD.AsString:=Table.FieldByName('POSTCODE').ASString;
     if Form1.Spr00.GoToKey then
        begin
        StoPostComboBox.ItemIndex:=StoPostComboBox.Items.IndexOf(Form1.Spr00NAIM.ASString);
        StoPostComboBox.Visible:=True;
        addRecordCheckBox.Visible:=True;
        addRecordCheckBox.Checked:=FAlse;
        end;


      Rec_IDLAbel.CAption:=Table.FieldByName('REC_ID').ASString;


      DAteTimePicker1.Date:=Table.FieldByName('DATE_APP').VAlue;
      DAteTimePicker1.Visible:=True;
      LAbel26.Visible:=True;

      ClientNameMaskEdit.Text:=Table.FieldByName('FIO').AsString;

      StartTimeMaskEdit.Text:=Copy(Table.FieldByName('TIME1').AsString,1,2)+':'+Copy(Table.FieldByName('TIME1').AsString,3,2);

      if Trim(Table.FieldByName('TIME2').AsString)<>'' then
         EndTimeMaskEdit.Text:=Copy(Table.FieldByName('TIME2').AsString,1,2)+':'+Copy(Table.FieldByName('TIME2').AsString,3,2)
      else
         EndTimeMaskEdit.Text:='';


      CarNumberMaskEdit.Text:=UnCryptString(Table.FieldByName('GOS_N').AsString,key1c,key2c); //gos_n

      MarkaMaskEdit.Text:=Table.FieldByName('MARKA').AsString; //marka
      ModelMaskEdit.Text:=Table.FieldByName('MODEL').AsString;
      OrderNumberMaskEdit.Text:=Table.FieldByName('N_ZAJ').AsString;
      k_klLabel.Caption:=Table.FieldByName('K_KL').AsString;

      if trim(k_klLabel.Caption)<>'' then
         ClientSpravBitBtn.Visible:=True
      else
         ClientSpravBitBtn.Visible:=False;

      tel:=UnCryptString(Table.FieldByName('TEL').AsString,key1c,key2c);

      if Pos('+',tel)=1 then
         telCountryCodeMaskEdit.Text:=Copy(tel,1,2)
      else
         telCountryCodeMaskEdit.Text:='+7';

      telCodeMaskEdit.Text:=Copy(tel,3,3);
      telNumberMaskEdit.Text:=Copy(tel,6,10);

      car_n_zapLAbel.CAption:=Table.FieldByName('CAR_N_ZAP').AsString;

      if trim(car_n_zapLAbel.CAption)<>'' then
         begin
         BitBtn15.Visible:=True;
         BitBtn15.Enabled:=True;
         end;

      ReasonMemo.Text:=UnCryptString(Table.FieldByName('OBR_KL').AsString,key1c,key2c);
      CommentMemo.Text:=UnCryptString(Table.FieldByName('COMM').AsString,key1c,key2c);

 end;
end;




procedure TAppointmentForm.FormShow(Sender: TObject);
var
  IniFile: TIniFile;
  _year:   string;
begin
  Tab20.TabVisible:=False;

  if ccfForm.CheckBox53.Checked=True then
     ConfirmSMSCheckBox.Enabled:=True
  else
     ConfirmSMSCheckBox.Enabled:=False;

  AppointmentForm.Color:=Form1.Color;

  dsCalendar1.Colors.BackGround:=Form1.Color;
  dsCalendar1.Colors.Title:=Form1.PAnel1.Color;
  dsCalendar1.Colors.Selected:=Form1.PAnel1.Color;
  dsCalendar1.Colors.Circle:=Form1.Label2.Color;
  dsCalendar1.Colors.SelectedFont:=clBlack;
  dsCalendar1.Colors.Holidays:=$00AAAAFF;

  dsCalendar2.Colors.BackGround:=Form1.Color;
  dsCalendar2.Colors.Title:=Form1.PAnel1.Color;
  dsCalendar2.Colors.Selected:=Form1.PAnel1.Color;
  dsCalendar2.Colors.Circle:=Form1.Label2.Color;
  dsCalendar2.Colors.SelectedFont:=clBlack;
  dsCalendar2.Colors.Holidays:=Form1.Color;


  dsCalendar2.Date:=Date;

  PAnel1.Color:=Form1.PAnel1.Color;
  PAnel2.Color:=$0080FFFF;
  PAnel3.Color:=$00AAAAFF;

  DetailPanel.Color:=Form1.Color;
  PAnel6.Color:=Form1.Color;
  PAnel9.Color:=Form1.Color;

  DBGrid1.Color:=Form1.PAnel1.Color;
  DBGrid1.FixedColor:=Form1.Color;

  DBGrid2.Color:=Form1.PAnel1.Color;
  DBGrid2.FixedColor:=Form1.Color;

  DBGrid3.Color:=Form1.PAnel1.Color;
  DBGrid3.FixedColor:=Form1.Color;

  DBGrid4.Color:=Form1.PAnel1.Color;
  DBGrid4.FixedColor:=Form1.Color;

  DBGrid7.Color:=Form1.Color;
  DBGrid7.FixedColor:=Form1.Color;

  DBGrid9.Color:=Form1.Color;
  DBGrid9.FixedColor:=Form1.Color;

  DBGrid20.Color:=Form1.PAnel1.Color;
  DBGrid20.FixedColor:=Form1.Color;

  LAbel2.Font.Color:=Form1.LAbel3.Font.Color;
  LAbel18.Font.Color:=Form1.LAbel3.Font.Color;
  LAbel13.Font.Color:=Form1.LAbel3.Font.Color;

  ScaleComboBox.Color:=Form1.PAnel1.Color;
  ResourceItemsComboBox.Color:=Form1.PAnel1.Color;
  MonthsComboBox.Color:=Form1.PAnel1.Color;
  YearComboBox.Color:=Form1.PAnel1.Color;
  WorkPostsComboBox.Color:=Form1.PAnel1.Color;
  AllPostsComboBox.Color:=Form1.PAnel1.Color;

  MonMaskEdit.Color:=Form1.PAnel1.Color;
  TueMaskEdit.Color:=Form1.PAnel1.Color;
  WedMaskEdit.Color:=Form1.PAnel1.Color;
  ThuMaskEdit.Color:=Form1.PAnel1.Color;
  FriMaskEdit.Color:=Form1.PAnel1.Color;
  SatMaskEdit.Color:=Form1.PAnel1.Color;
  SunMaskEdit.Color:=Form1.PAnel1.Color;
  TelMaskEditCountryCode.Color:=Form1.PAnel1.Color;
  TelMaskEditCode.Color:=Form1.PAnel1.Color;
  TelMaskEditNumber.Color:=Form1.PAnel1.Color;
  CarNumberFindEdit.Color:=Form1.PAnel1.Color;

  ClientCarsListBox.Color:=Form1.PAnel1.Color;

  DateTimePicker2.Color:=Form1.PAnel1.Color;
  DateTimePicker3.Color:=Form1.PAnel1.Color;

  PageControl1.ActivePage:=TAb1;

  IniFile:=TIniFile.Create(path_base+'\PEREM\appoint.ini');

  MonMaskEdit.Text:=IniFile.ReadString('WorkTime','1','с 08:00 до 20:00');
  TueMaskEdit.Text:=IniFile.ReadString('WorkTime','2','с 08:00 до 20:00');
  WedMaskEdit.Text:=IniFile.ReadString('WorkTime','3','с 08:00 до 20:00');
  ThuMaskEdit.Text:=IniFile.ReadString('WorkTime','4','с 08:00 до 20:00');
  FriMaskEdit.Text:=IniFile.ReadString('WorkTime','5','с 08:00 до 20:00');
  SatMaskEdit.Text:=IniFile.ReadString('WorkTime','6','с 08:00 до 20:00');
  SunMaskEdit.Text:=IniFile.ReadString('WorkTime','7','с 08:00 до 20:00');

  ScaleComboBox.ItemIndex:=StrToInt(IniFile.ReadString('Scale','value','0'));

  IniFile.Free;

  LAbel2.Caption:='Запись на '+FormatDAteTime('MMMM DD',dsCAlendar1.DAte)+'-е, '+DayOfTheWeek(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte));


  dsCalendar1.Date:=DAte;
  dsCalendar1DateChange(dsCalendar1, date-1, date);

  if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
     _year:='NY'
  else
     _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


  d_offTable.Close;
  d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
  d_offTable.IndexFiles.Clear;
  d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');


  if FileExists(d_offTable.TableName)=False then
     begin
     Create_d_offTable(_year);
     Create_d_offTable('NY');
     end;

  MonthsComboBox.ItemIndex:=StrToInt(FormatDateTime('MM',date))-1;
  MonthsComboBoxChange(MonthsComboBox);


  fill_sto_days_off(dsCAlendar1,'STO');

  if SunMaskEdit.Text='с   :   до   :  ' then
     SunMaskEdit.Color:=$00AAAAFF
  else
     SunMaskEdit.Color:=Form1.PAnel1.Color;

  if SatMaskEdit.Text='с   :   до   :  ' then
     SatMaskEdit.Color:=$00AAAAFF
  else
     SatMaskEdit.Color:=Form1.PAnel1.Color;

  if MonMaskEdit.Text='с   :   до   :  ' then
     MonMaskEdit.Color:=$00AAAAFF
  else
     MonMaskEdit.Color:=Form1.PAnel1.Color;

  if TueMaskEdit.Text='с   :   до   :  ' then
     TueMaskEdit.Color:=$00AAAAFF
  else
     TueMaskEdit.Color:=Form1.PAnel1.Color;

  if WedMaskEdit.Text='с   :   до   :  ' then
     WedMaskEdit.Color:=$00AAAAFF
  else
     WedMaskEdit.Color:=Form1.PAnel1.Color;

  if ThuMaskEdit.Text='с   :   до   :  ' then
     ThuMaskEdit.Color:=$00AAAAFF
  else
     ThuMaskEdit.Color:=Form1.PAnel1.Color;

  if FriMaskEdit.Text='с   :   до   :  ' then
     FriMaskEdit.Color:=$00AAAAFF
  else
     FriMaskEdit.Color:=Form1.PAnel1.Color;


  Panel6.Width:=Tab1.Width-BitBtn1.Width;
  DBGrid3.Width:=PAnel6.Width-DBGrid3.Left*2;
  SpeedButton2.Left:=PAnel7.Width-SpeedButton2.Width-3;


end;

procedure TAppointmentForm.BitBtn7Click(Sender: TObject);
begin
   if DetailPanel.Visible=True then
      begin
      DetailPanel.Visible:=FAlse;
      exit;
      end;

   if PAnel6.Visible=True then
      begin
      PAnel6.Visible:=FAlse;
      BitBtn1.Click;
      exit;
      end;

   if PAnel9.Visible=True then
      begin
      PAnel9.Visible:=FAlse;
      exit;
      end;


   AppointmentForm.Close;

end;

procedure TAppointmentForm.FormResize(Sender: TObject);
begin
  DBGrid9.Left:=DBGrid7.Left;

  PageControl1.Top:=dsCalendar1.Top+dsCalendar1.Height+5;
  PageControl1.Width:=AppointmentForm.ClientWidth-PageControl1.Left-10;
  PageControl1.Height:=AppointmentForm.ClientHeight-PageControl1.Top-BitBtn7.Height;
  BitBtn7.Left:=PageControl1.Left+PageControl1.Width-BitBtn7.Width;


  PAnel8.Top:=dsCAlendar1.Top+dsCalendar1.Height-Panel8.Height;
  PAnel8.Left:=PageControl1.Left+PageControl1.Width-PAnel8.Width;


  DBGrid1.Width:=TAb1.ClientWidth-2*DBGrid1.Left;
  DBGrid2.Width:=TAb2.ClientWidth-2*DBGrid2.Left;
  DBGrid20.Width:=TAb2.ClientWidth-2*DBGrid2.Left;

  DBGrid1.Height:=TAb1.ClientHeight-DBGrid1.Top-Panel0.Height-10;
  DBGrid2.Height:=TAb2.ClientHeight-DBGrid2.Top-DeleteBitBtn.Height-10;
  DBGrid20.Height:=TAb2.ClientHeight-DBGrid2.Top-DeleteBitBtn.Height-10;

  Panel0.Top:=DBGrid1.Top+DBGrid1.Height;

  DeleteBitBtn.Top:=DBGrid2.Top+DBGrid2.Height+5;
  EditOrderBitBtn2.Top:=DBGrid2.Top+DBGrid2.Height+5;
  CreateOrderBitBtn3.Top:=DBGrid2.Top+DBGrid2.Height+5;
  BitBtn17.Top:=DBGrid2.Top+DBGrid2.Height+5;


  DetailPanel.Left:=(AppointmentForm.ClientWidth div 2) - (DetailPanel.Width div 2);
  DetailPanel.Top:=(AppointmentForm.ClientHeight div 2) - (DetailPanel.Height div 2);


  Panel6.Width:=Tab1.Width-BitBtn1.Width;
  DBGrid3.Width:=PAnel6.Width-DBGrid3.Left*2;
  SpeedButton2.Left:=PAnel7.Width-SpeedButton2.Width-3;

  PAnel6.Left:=(AppointmentForm.ClientWidth div 2) - (PAnel6.Width div 2);
  PAnel6.Top:=(AppointmentForm.ClientHeight div 2) - (PAnel6.Height div 2) -  BitBtn7.Height;

  LAbel2.Left:=dsCAlendar1.Left+dsCAlendar1.Width+20;
  LAbel2.Top:=dsCAlendar1.Top+dsCAlendar1.Height-(dsCAlendar1.Height div 9)*2;

  label27.left:=label2.left;
  StaffListBox.left:=Label2.Left;

  ScaleComboBox.Left:=LAbel1.Left+LAbel1.Width+10;

  DBGrid9.Height:=Tab3.ClientHeight-DBGrid9.Top-10;
  DBGrid9.Width:=Tab3.ClientWidth-DBGrid9.Left-10;

  StaffListBox.Height:=LAbel2.Top-5-StaffListBox.Top;
  StaffListBox.Height:=(StaffListBox.Height div StaffListBox.Canvas.TextHeight('Aa'))*StaffListBox.Canvas.TextHeight('Aa');

end;

procedure TAppointmentForm.dsCalendar1DateChange(Sender: TObject; FromDate,
  ToDate: TDateTime);
var
  _year: string;
begin
  if Loader_key=false then exit;

  if AppointmentForm.Showing=True then
     begin
     LAbel2.Caption:='Запись на '+FormatDAteTime('MMMM DD',dsCAlendar1.DAte)+'-е, '+DayOfTheWeek(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte));
     BitBtn1.Click;
     end;



  if FormatDAteTime('YYYY',fromDate)<>FormatDAteTime('YYYY',toDate) then
     begin
     if StrToInt(FormatDAteTime('YYYY',toDate))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',toDate);



     if d_offTableSTO.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
        begin
        d_offTableSTO.Close;
        d_offTableSTO.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
        d_offTableSTO.IndexFiles.Clear;
        d_offTableSTO.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');
        end;

     d_offTableSTO.Open;
     d_offTableSTO.IndexName:='KOD';
     d_offTableSTO.SetRange('STO   '+'','STO   '+'яя');

     fill_sto_days_off(dsCalendar1,'STO');

     end;



  Tab2.Caption:='   Все записи на '+Copy(LAbel2.Caption,11,100)+'   ';
  Tab4.Caption:='   Все записи на '+FormatDateTime('mmmm yyyy',dsCalendar1.Date)+'   ';
end;

procedure TAppointmentForm.ScaleComboBoxChange(Sender: TObject);
begin
 if (AppointmentForm.Showing=True) then
     BitBtn1.Click;
end;

procedure TAppointmentForm.MonMaskEditChange(Sender: TObject);
begin
  if MonMaskEdit.Focused then
     begin
     TueMaskEdit.Text:=MonMaskEdit.Text;
     WedMaskEdit.Text:=MonMaskEdit.Text;
     ThuMaskEdit.Text:=MonMaskEdit.Text;
     FriMaskEdit.Text:=MonMaskEdit.Text;
     SatMaskEdit.Text:=MonMaskEdit.Text;
     SunMaskEdit.Text:=MonMaskEdit.Text;
     end;


  if MonMaskEdit.Text='с   :   до   :  ' then
     MonMaskEdit.Color:=$00AAAAFF
  else
     MonMaskEdit.Color:=Form1.Panel2.Color;


end;

procedure TAppointmentForm.SatMaskEditChange(Sender: TObject);
begin
  if SatMaskEdit.Focused then
     SunMaskEdit.Text:=MonMaskEdit.Text;

  if SatMaskEdit.Text='с   :   до   :  ' then
     SatMaskEdit.Color:=$00AAAAFF
  else
     SatMaskEdit.Color:=Form1.Panel2.Color;

end;

procedure TAppointmentForm.MonMaskEditClick(Sender: TObject);
begin
  MonMaskEdit.SelectAll;
end;

procedure TAppointmentForm.TueMaskEditClick(Sender: TObject);
begin
  TueMaskEdit.SelectAll;
end;

procedure TAppointmentForm.TueMaskEditChange(Sender: TObject);
begin
  if TueMaskEdit.Focused then
     begin
     WedMaskEdit.Text:=MonMaskEdit.Text;
     ThuMaskEdit.Text:=MonMaskEdit.Text;
     FriMaskEdit.Text:=MonMaskEdit.Text;
     SatMaskEdit.Text:=MonMaskEdit.Text;
     SunMaskEdit.Text:=MonMaskEdit.Text;
     end;


  if TueMaskEdit.Text='с   :   до   :  ' then
     TueMaskEdit.Color:=$00AAAAFF
  else
     TueMaskEdit.Color:=Form1.Panel2.Color;

end;

procedure TAppointmentForm.MonMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_DOWN then TueMaskEdit.SetFocus;

end;

procedure TAppointmentForm.TueMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then MonMaskEdit.SetFocus;
  if key=VK_DOWN then WedMaskEdit.SetFocus;

end;

procedure TAppointmentForm.WedMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then TueMaskEdit.SetFocus;
  if key=VK_DOWN then ThuMaskEdit.SetFocus;

end;

procedure TAppointmentForm.ThuMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then WedMaskEdit.SetFocus;
  if key=VK_DOWN then FriMaskEdit.SetFocus;

end;

procedure TAppointmentForm.FriMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then ThuMaskEdit.SetFocus;
  if key=VK_DOWN then SatMaskEdit.SetFocus;

end;

procedure TAppointmentForm.SatMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then FriMaskEdit.SetFocus;
  if key=VK_DOWN then SunMaskEdit.SetFocus;

end;

procedure TAppointmentForm.SunMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=VK_UP then SatMaskEdit.SetFocus;

end;

procedure TAppointmentForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
var
  IniFile: TIniFile;
begin
  RezhimLabel.Caption:='';

  FormForSendSMS_addRec:=FAlse;
  FormForSendSMS_editRec:=FAlse;

  IniFile:=TIniFile.Create(path_base+'\PEREM\appoint.ini');

  IniFile.WriteString('WorkTime','1',MonMaskEdit.Text);
  IniFile.WriteString('WorkTime','2',TueMaskEdit.Text);
  IniFile.WriteString('WorkTime','3',WedMaskEdit.Text);
  IniFile.WriteString('WorkTime','4',ThuMaskEdit.Text);
  IniFile.WriteString('WorkTime','5',FriMaskEdit.Text);
  IniFile.WriteString('WorkTime','6',SatMaskEdit.Text);
  IniFile.WriteString('WorkTime','7',SunMaskEdit.Text);

  IniFile.WriteString('Scale','value',IntToStr(ScaleComboBox.ItemIndex));

  IniFile.Free;

  DetailPanel.Visible:=False;
  Panel6.Visible:=False;

end;


function createDayTableProc(): Boolean;
var
  day_of_week, i: integer;
  end_time, last_name_time, NewName: string;
begin
with AppointmentForm do
begin

  day_of_week:=DayOfWeek(dsCAlendar1.Date)-1;

  DayTable.Active:=False;
  DayTable.TableName := tmp_path+'AppDayTable.dbf';

  with DayTable do
       begin
       Active := False;
       TableType := ttFoxPro;
       DayTable.FieldDefs.Clear;


       with DayTable.FieldDefs.AddFieldDef do
            begin
            Name := 'POSTCODE';
            DataType := ftString;
            Size := 2;
            end;

       with DayTable.FieldDefs.AddFieldDef do
            begin
            Name := 'NAIM';
            DataType := ftString;
            Size := 40;
            end;


       with DayTable.FieldDefs.AddFieldDef do
            begin

            //первое доступное время
            case day_of_week of
               1: begin
                  Name := 'H'+Copy(MonMaskEdit.Text,3,2);
                  end_time:='H'+Copy(MonMaskEdit.Text,12,2);
                  last_name_time:=Copy(MonMaskEdit.Text,3,2);
                  end;

               2: begin
                  Name := 'H'+Copy(TueMaskEdit.Text,3,2);
                  end_time:='H'+Copy(TueMaskEdit.Text,12,2);
                  last_name_time:=Copy(TueMaskEdit.Text,3,2);
                  end;

               3: begin
                  Name := 'H'+Copy(WedMaskEdit.Text,3,2);
                  end_time:='H'+Copy(WedMaskEdit.Text,12,2);
                  last_name_time:=Copy(WedMaskEdit.Text,3,2);
                  end;

               4: begin
                  Name := 'H'+Copy(ThuMaskEdit.Text,3,2);
                  end_time:='H'+Copy(ThuMaskEdit.Text,12,2);
                  last_name_time:=Copy(ThuMaskEdit.Text,3,2);
                  end;

               5: begin
                  Name := 'H'+Copy(FriMaskEdit.Text,3,2);
                  end_time:='H'+Copy(FriMaskEdit.Text,12,2);
                  last_name_time:=Copy(FriMaskEdit.Text,3,2);
                  end;

               6: begin
                  Name := 'H'+Copy(SatMaskEdit.Text,3,2);
                  end_time:='H'+Copy(SatMaskEdit.Text,12,2);
                  last_name_time:=Copy(SatMaskEdit.Text,3,2);
                  end;

               0: begin
                  Name := 'H'+Copy(SunMaskEdit.Text,3,2);
                  end_time:='H'+Copy(SunMaskEdit.Text,12,2);
                  last_name_time:=Copy(SunMaskEdit.Text,3,2);
                  end;
               end;

            DataType := ftString;
            Size := 5;


            NewName := Name;
            last_name_time:=Copy(NewName,2,2);
            end;

                 //+ 1/2 часа
                 if ScaleComboBox.ItemIndex=1 then
                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'30';
                         DataType := ftString;
                         Size := 5;
                         end;

                 //+ 1/4 часа
                 if ScaleComboBox.ItemIndex=2 then
                    begin
                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'15';
                         DataType := ftString;
                         Size := 5;
                         end;

                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'30';
                         DataType := ftString;
                         Size := 5;
                         end;

                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'45';
                         DataType := ftString;
                         Size := 5;
                         end;

                    end;
            //3108



            //следующие до закрытия
            for  i:=1 to 100 do
                 begin


                 if StrToInt(last_name_time)+1>=10 then
                    begin
                    NewName := 'H'+IntToStr(StrToInt(last_name_time)+1);
                    last_name_time:=Copy(NewName,2,2);
                    end
                  else
                    begin
                    NewName := 'H0'+IntToStr(StrToInt(last_name_time)+1);
                    last_name_time:=Copy(NewName,3,2);
                    end;

                if NewName=end_time then break;

                 //час
                 with DayTable.FieldDefs.AddFieldDef do
                      begin
                      Name := NewName;
                      DataType := ftString;
                      Size := 5;
                      end;

                 //+ 1/2 часа
                 if ScaleComboBox.ItemIndex=1 then
                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'30';
                         DataType := ftString;
                         Size := 5;
                         end;

                 //+ 1/4 часа
                 if ScaleComboBox.ItemIndex=2 then
                    begin
                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'15';
                         DataType := ftString;
                         Size := 5;
                         end;

                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'30';
                         DataType := ftString;
                         Size := 5;
                         end;

                    with DayTable.FieldDefs.AddFieldDef do
                         begin
                         Name := NewName+'45';
                         DataType := ftString;
                         Size := 5;
                         end;

                    end;


                 end;


       IndexDefs.Clear;
       end; // with dayTable



    try
    DayTable.CreateTable;

    createDayTableProc:=True;
    result:=True;
    except
      createDayTableProc:=False;
    end;

 end;//with AppointmentForm


end;


procedure TAppointmentForm.BitBtn1Click(Sender: TObject);
var
  i, col_index1, col_index2: integer;
  pole, last_time, last_color, _year, gos_n: string;
  holiday_key, car_here: Boolean;
begin

   PAnel6.Visible:=False;
   DetailPanel.Visible:=False;



     if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


   if AppTAble.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
      begin
      AppTAble.Close;
      AppTAble.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
      AppTAble.IndexFiles.Clear;
      AppTAble.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
      end;

  if FileExists(AppTAble.TableName)=False then
     begin
     AppTAble.Close;
     AppTAble_ForCreate.Close;

     AppTAble_ForCreate.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
     AppTAble_ForCreate.CreateTable;

     AppTAble.Close;
     AppTAble_ForCreate.Close;

     AppTAble.Close;
     AppTAble.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
     AppTAble.IndexFiles.Clear;

     AppTAble.Open;
     AppTAble.IndexOn(path_base+'HIST\'+_year+'\apptable.cdx','DATE_APP','DTOS(DATE_APP)+POSTCODE+TIME1','',Duplicates,Ascending);
     AppTAble.IndexOn(path_base+'HIST\'+_year+'\apptable.cdx','DATE_APP1','DTOS(DATE_APP)+NAVL_PR+TIME1','',Duplicates,Ascending);
     AppTAble.IndexOn(path_base+'HIST\'+_year+'\apptable.cdx','REC_ID','REC_ID','',Duplicates,Ascending);

     AppTAble.Reindex;
     AppTAble.Close;

     //В NY еще добавляем
     if FileExists(path_base+'HIST\NY\apptable.dbf')=False then
        begin
        AppTAble_ForCreate.TableName:=path_base+'HIST\NY\apptable.dbf';
        AppTAble_ForCreate.CreateTable;

        AppTAble.Close;
        AppTAble_ForCreate.Close;

        AppTAble.Close;
        AppTAble.TableName:=path_base+'HIST\NY\apptable.dbf';
        AppTAble.IndexFiles.Clear;

        AppTAble.Open;
        AppTAble.IndexOn(path_base+'HIST\NY\apptable.cdx','DATE_APP','DTOS(DATE_APP)+POSTCODE+TIME1','',Duplicates,Ascending);
        AppTAble.IndexOn(path_base+'HIST\NY\apptable.cdx','DATE_APP1','DTOS(DATE_APP)+NAVL_PR+TIME1','',Duplicates,Ascending);
        AppTAble.IndexOn(path_base+'HIST\NY\apptable.cdx','REC_ID','REC_ID','',Duplicates,Ascending);


        AppTAble.Reindex;
        AppTAble.Close;

        AppTAble.IndexFiles.Clear;
        AppTAble.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
        end;
      end;





  Form1.Spr00.Open;
  Form1.Spr00.IndexName:='NAIM';
  Form1.Spr00.SetRange('30'+'','30'+'яя');
  if Form1.Spr00.RecordCount=0 then
     begin
     my_messageTime('Внимание!','Задайте ресурсы автосервиса в "Общем справочнике".',clYellow,5000);
     exit;
     end;

  DAyTAble.Close;


  //создаем дневную таблицу
  if createDayTableproc()=False then
     begin
     my_messageTime('Внимание!','Ошибка создания временного файла.',clYellow,5000);
     exit;
     end;



  DAyTAble.Open;

  try
  DayTable.Fields[0].Visible:=False;
  DayTable.Fields[1].DisplayLabel:='Список постов:';

  for i:=2 to 100 do
      begin
      if (Length(DayTable.Fields[i].DisplayName)<4) then
          DayTable.Fields[i].DisplayLAbel:=Copy(DayTable.Fields[i].DisplayName,2,2)+'.00'
      else
          DayTable.Fields[i].DisplayLAbel:=Copy(DayTable.Fields[i].DisplayName,2,2)+'.'+Copy(DayTable.Fields[i].DisplayName,4,2);

      if ScaleComboBox.ItemIndex>0 then
         DayTable.Fields[i].DisplayWidth:=5
      else
         DayTable.Fields[i].DisplayWidth:=7;
      end;
  except
  end;




    try
    holiday_key:=FAlse;

    if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
       begin
       d_offTable.Close;
       d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
       d_offTable.IndexFiles.Clear;
       d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');


      if FileExists(d_offTable.TableName)=False then
         begin
         Create_d_offTable(_year);
         Create_d_offTable('NY');
         end;
      end;


    d_offTable.Open;
    d_offTable.Setrange('','');
    d_offTable.IndexName:='KOD';
    d_offTable.SetKey;
    d_offTableKOD.AsString:='STO';
    d_offTableD1.Value:=dsCalendar1.Date;
    if d_offTable.GotoKey then
       holiday_key:=True;

    except
    end;






  //заполняем дневную таблицу по рабочему справочнику (ресурсы автосервиса)
  Form1.Spr00.First;

  StoPostComboBox.Items.Clear;

  ResourceItemsComboBox.Items.Clear;
  ResourceItemsComboBox.Items.Add('ВСЕ ресурсы СТО');
  ResourceItemsComboBox.Items.Add('Персонал СТО');
  ResourceItemsComboBox.Items.Add('Нерабочие дни СТО');

  WorkPostsComboBox.Items.Clear;
  WorkPostsComboBox.Items.Add('Все ресурсы СТО');

  AllPostsComboBox.Items.Clear;
  AllPostsComboBox.Items.Add('Все ресурсы СТО');

  while Form1.Spr00.Eof=False do
        begin
        DAyTAble.AppendRecord([Form1.Spr00KOD.AsString,
                               Form1.Spr00NAIM.AsString]);


        StoPostComboBox.Items.Add(Form1.Spr00NAIM.AsString);
        ResourceItemsComboBox.Items.Add(Form1.Spr00NAIM.AsString);

        WorkPostsComboBox.Items.Add(Form1.Spr00NAIM.AsString);
        AllPostsComboBox.Items.Add(Form1.Spr00NAIM.AsString);


        Form1.Spr00.Next;
        end;


  WorkPostsComboBox.ItemIndex:=0;
  AllPostsComboBox.ItemIndex:=0;


  ResourceItemsComboBox.ItemIndex:=0;
  ResourceItemsComboBoxChange(ResourceItemsComboBox);


  if DAyTAble.RecordCount<=5 then DBGrid1.RowLines:=3
     else DBGrid1.RowLines:=2;


  //заполняем занятое и свободное время

  DAtaSource2.Enabled:=False;

  AppTAble.Open;

  AppTAble.IndexName:='DATE_APP';

  DAyTAble.First;
  while DAyTAble.Eof=FAlse do
        begin
        if DAyTAble.FieldByName('POSTCODE').AsString<>'' then
           begin

           if holiday_key=True then
              begin  //СТО не работает
              for i:=2 to DAyTAble.Fields.Count-1 do
                  begin
                  DAyTAble.Edit;
                  DAyTAble.Fields[i].Value:='X';
                  DAyTAble.Edit;
                  DAyTAble.Post;
                  end;
              end;


        d_offTable.Open;  //пост не работает
        d_offTable.Setrange('','');
        d_offTable.IndexName:='KOD';
        d_offTable.SetKey;
        d_offTableKOD.AsString:=DAyTAble.FieldByName('POSTCODE').AsString;
        d_offTableD1.Value:=dsCalendar1.DAte;
        if d_offTable.GotoKey then
           for i:=2 to DAyTAble.Fields.Count-1 do
               begin
               DAyTAble.Edit;
               DAyTAble.Fields[i].Value:='X';
               DAyTAble.Edit;
               DAyTAble.Post;
               end;


        AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+'яя');
        AppTAble.First;
        gos_n:='';
        while AppTAble.Eof=FAlse do
              begin


                 if AppTAbleD_UDL.ASString='' then
                    begin

                    try
                    try
                    pole:=DAyTAble.FieldByNAme('H'+AppTAbleTIME1.ASString).FieldNAme;
                    except
                      pole:=DAyTAble.FieldByNAme('H'+Copy(AppTAbleTIME1.ASString,1,2)).FieldNAme;
                    end;

                    except
                      Pole:=DAyTAble.Fields[2].FieldNAme;
                    end;


                    col_index1:=DAyTAble.FieldByNAme(pole).Index;


                    DAyTAble.Edit;

                    if AppTAbleNAVL_PR.ASString='' then
                       begin
                       try
                       if DAyTAble.FieldByNAme(pole).ASString='' then
                          begin
                          gos_n:='№'+Copy(remove_all_except_digits(UnCryptString(AppTAbleGOS_N.ASString,key1c,key2c)),1,3);
                          if trim(AppTAbleN_zaj.ASString)<>'' then
                             gos_n:='+'+gos_n;

                          if trim(gos_n)<>'' then
                             DAyTAble.FieldByNAme(pole).ASString:=gos_n//'1';
                          else
                             DAyTAble.FieldByNAme(pole).ASString:='1';
                          end
                       else
                          begin
                          car_here:=False;
                          if (Copy(DAyTAble.FieldByNAme(pole).ASString,1,1)='+') then car_here:=True;

                          DAyTAble.FieldByNAme(pole).ASString:=IntToStr(StrToInt(DAyTAble.FieldByNAme(pole).ASString)+1);

                           if (car_here=True) and (trim(AppTAbleN_zaj.ASString)<>'') then
                               DAyTAble.FieldByNAme(pole).ASString:='+'+DAyTAble.FieldByNAme(pole).ASString;

                          end;

                       except
                           if (Copy(DAyTAble.FieldByNAme(pole).ASString,1,1)='+') and (trim(AppTAbleN_zaj.ASString)<>'') then
                               DAyTAble.FieldByNAme(pole).ASString:='+2'
                           else
                               DAyTAble.FieldByNAme(pole).ASString:='2'

                       end;
                       end
                    else
                       DAyTAble.FieldByNAme(pole).ASString:='X';

                    DAyTAble.Edit;
                    DAyTAble.Post;



                    if AppTAbleTIME2.ASString<>'' then
                       begin //задано время окончания, значит надо закраcить все до времени окончания

                       try
                       try
                       pole:=DAyTAble.FieldByNAme('H'+AppTAbleTIME2.ASString).FieldNAme;
                       except
                         pole:=DAyTAble.FieldByNAme('H'+Copy(AppTAbleTIME2.ASString,1,2)).FieldNAme;
                       end;

                       col_index2:=DAyTAble.FieldByNAme(pole).Index;
                       except
                         col_index2:=DAyTAble.Fields.Count;
                       end;



                       inc(col_index1);// т.к. начинаем со следующей колонки, в этой уже все проставили выше
                       while col_index1<col_index2 do
                             begin

                             DAyTAble.Edit;
                             try
                             if AppTAbleNAVL_PR.ASString='' then
                                begin

                                try
                                if DAyTAble.Fields[col_index1].ASString='' then
                                   begin
                                   if trim(gos_n)<>'' then
                                      DAyTAble.Fields[col_index1].ASString:=gos_n//'1';
                                   else
                                      DAyTAble.Fields[col_index1].ASString:='1';
                                   end

                                else
                                   begin
                                   car_here:=False;
                                   if (Copy(DAyTAble.Fields[col_index1].ASString,1,1)='+') then car_here:=True;

                                   DAyTAble.Fields[col_index1].ASString:=IntToStr(StrToInt(DAyTAble.Fields[col_index1].ASString)+1);

                                   if (car_here=True) and (trim(AppTAbleN_zaj.ASString)<>'') then
                                       DAyTAble.Fields[col_index1].ASString:='+'+DAyTAble.Fields[col_index1].ASString;

                                   end;

                                except

                                   if (Copy(DAyTAble.Fields[col_index1].ASString,1,1)='+') and (trim(AppTAbleN_zaj.ASString)<>'') then
                                       DAyTAble.Fields[col_index1].ASString:='+2'
                                   else
                                       DAyTAble.Fields[col_index1].ASString:='2'

                                end;

                                end
                             else
                                DAyTAble.Fields[col_index1].ASString:='X';

                             except
                             end;

                             DAyTAble.Edit;
                             DAyTAble.Post;


                             inc(col_index1);
                             end;


                       end;


                    end; //не UDL


                 AppTAble.Next;
                 end;
           end;

        DAyTAble.Next;
        end;


  AppTAble.Open;
  AppTAble.IndexName:='DATE_APP1';
  AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+''+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+'"'+'яя');
  AppTAble.First;

  //раскрашиваем по одинаковому времени
  last_time:=AppTableTIME1.AsString;
  Last_color:='';
  AppTAble.Next;
  While AppTAble.Eof=FAlse do
        begin
        if last_time<>AppTableTIME1.AsString then
           begin
           last_time:=AppTableTIME1.AsString;

           if Last_color='' then
               Last_color:='1'
           else
               Last_color:='';

           end;

        try
        AppTAble.Edit;
        AppTAbleBRIGHT.AsString:=Last_Color;
        AppTAble.Edit;
        AppTAble.Post;
        except
        end;


        AppTAble.Next;
        end;



  AppTAble.First;

  DAtaSource2.Enabled:=True;


  if PageControl1.ACtivePage=Tab1 then
     dbGrid1.SetFocus
  else
  if PageControl1.ACtivePage=Tab1 then
     dbGrid2.SetFocus;


   if AppTAbleMonth.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
      begin
      AppTAbleMonth.Close;
      AppTAbleMonth.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
      AppTAbleMonth.IndexFiles.Clear;
      AppTAbleMonth.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
      end;



     AppTAbleMonth.Open;
     AppTAbleMonth.IndexName:='DATE_APP1';
     AppTAbleMonth.SetRange(FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'01'+''+'',FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'31'+'"'+'яя');
     AppTAbleMonth.First;




  Label27.Caption:='Персонал СТО на '+FormatDAteTime('MMMM DD',dsCAlendar1.DAte)+'-е, '+DayOfTheWeek(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte));
  StaffListBox.Clear;

  d_offTable.Open;
  d_offTable.Setrange('','');
  d_offTable.IndexName:='KOD';
  d_offTable.SetKey;
  d_offTableKOD.AsString:='STO';
  d_offTableD1.Value:=dsCalendar1.DAte;
  if d_offTable.GotoKey=True then
     exit; // сечрвис не работает

  Form1.Spr00.Open;
  Form1.Spr00.IndexName:='NAIM';
  Form1.Spr00.SetRange('32'+'','32'+'яя');
  Form1.Spr00.First;
  while Form1.Spr00.Eof=False do
        begin
        if Form1.Spr00PR.AsString='M' then
           begin
           d_offTable.Open;  //пост не работает
           d_offTable.Setrange('','');
           d_offTable.IndexName:='KOD';
           d_offTable.SetKey;
           d_offTableKOD.AsString:='MM'+Form1.Spr00KOD.AsString;
           d_offTableD1.Value:=dsCalendar1.DAte;
           if d_offTable.GotoKey=False then
              StaffListBox.Items.Add(Copy('Мастер'+'          ',1,10)+' - '+Form1.Spr00NAIM.AsString);
           end;



        Form1.Spr00.Next;
        end;


  SqlQuery1.SQL.Clear;
  SqlQuery1.SQL.Add('SELECT DISTINCT TAB_N,FIO FROM "'+tab_nTable.TableName+'"');
  SqlQuery1.Open;
  SqlQuery1.First;



  while SqlQuery1.Eof=False do
        begin

        if Copy(SqlQuery1.FieldByName('FIO').AsString,1,1)<>'*' then
           begin
           d_offTable.Open;  //пост не работает
           d_offTable.Setrange('','');
           d_offTable.IndexName:='KOD';
           d_offTable.SetKey;
           d_offTableKOD.AsString:='ST'+SqlQuery1.FieldByName('TAB_N').ASString;
           d_offTableD1.Value:=dsCalendar1.DAte;
           if d_offTable.GotoKey=False then
              StaffListBox.Items.Add(Copy('Таб.№ '+SqlQuery1.FieldByName('TAB_N').ASString+'          ',1,10)+' - '+SqlQuery1.FieldByName('FIO').AsString);
           end;

        SqlQuery1.Next;
        end;



end;

procedure TAppointmentForm.DBGrid1DrawDataCell(Sender: TObject;
  const Rect: TRect; Field: TField; State: TGridDrawState);

procedure my_BrushProc(canvas: TCanvas;Rect: TRect; step: integer);
var
  x,y,x1,y1: integer;
begin

  x:= Rect.Left; y :=Rect.Top;

  x1:=Rect.Left; y1:=Rect.Top;


  x1:=x1+step;
  y:=y+step;

  canvas.Pen.Color:=clBlack;
  canvas.Pen.Width:=1;

  while (x<Rect.Right) do
        begin
        Canvas.MoveTo(x,y);
        Canvas.LineTo(x1,y1);

        if y<Rect.Bottom then y:=y+step
           else x:=x+step;

        if x1<Rect.Right then x1:=x1+step
           else y1:=y1+step
        end;

end;

var
  gos_n: string;
  Bitmap: TBitmap;
begin
   with DBGrid1.Canvas do
        begin
        if (Global_X<Rect.Right) and (Global_X>Rect.Left) then
            RealColumnNAme:=Field.FieldName;


        if (Global_X<Rect.Right) and (Global_X>Rect.Left) and (Global_x>FieldNAIM_Righ_X) then
             begin
             Brush.Color:=Install.PAnel2.Color;
             FillRect(Rect);


             Global_ColumnNAme:=Copy(Field.FieldName,2,10);
             if Length(Global_ColumnNAme)>2 then
                Global_ColumnNAme:=Copy(Global_ColumnNAme,1,2)+':'+Copy(Global_ColumnNAme,3,2)
             else
                Global_ColumnNAme:=Copy(Global_ColumnNAme,1,2)+':00';


             end;

        if ((Global_Y>Rect.Top) and (Global_Y<Rect.Bottom)) then
             begin
             Brush.Color:=Install.PAnel2.Color;
             FillRect(Rect);
             end;



        if ((Global_X<Rect.Right) and (Global_X>Rect.Left)) and
           ((Global_Y>Rect.Top) and (Global_Y<Rect.Bottom))  and (Global_x>FieldNAIM_Righ_X) then
             begin
             Brush.Color:=Install.PAnel2.Color;
             FillRect(Rect);
             Pen.Color:=clYellow;
             Pen.Width:=2;
             Rectangle(Rect);
             end;


        try
        if (Copy(Field.FieldName,1,1)='H') and (trim(Field.Value)<>'') then
           begin

           brush.Color:=PAnel2.Color;
           fillRect(Rect);
           Font.Size:=8;

           gos_n:=Field.Value;

           if Copy(gos_n,1,1)='+' then
              begin
              brush.Color:=PAnel2.Color;
              fillRect(Rect);

              gos_n:=Copy(gos_n,2,10);
              TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(gos_n) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), gos_n);

              my_BrushProc(DBGrid1.Canvas,Rect,5);

              end
           else
              begin
              TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(Field.Value) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), Field.Value);
              end;


           end;
        except
        end;


        try
        if (Field.Value>=1) then
           begin

           brush.Color:=PAnel2.Color;
           fillRect(Rect);
           Font.Size:=10;

           gos_n:=Field.Value;

           if Copy(gos_n,1,1)='+' then
              begin
              brush.Color:=PAnel2.Color;
              fillRect(Rect);

              gos_n:=Copy(gos_n,2,10);
              TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(gos_n) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), gos_n);

              my_BrushProc(DBGrid1.Canvas,Rect,5);


              end
           else
              begin
              TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(Field.Value) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), Field.Value);
              end;


           end;
        except
        end;

        if (Field.Value='X') then
           begin
           brush.Color:=PAnel3.Color;
           fillRect(Rect);
           Font.Size:=10;
           TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(Field.Value) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), Field.Value);
           end;



        if (Trim(Field.FieldName)='NAIM') then
            begin
            FieldNAIM_Righ_X:=Rect.Right;
            FillRect(Rect);
            TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(Field.Value) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), Field.Value);
            end;

        end;
end;

procedure TAppointmentForm.PageControl1Change(Sender: TObject);
var
  _year: string;
begin

  Global_X:=0;
  Global_Y:=0;


  DetailPanel.Visible:=False;
  Panel6.Visible:=False;


  if PageControl1.ActivePage=Tab1 then
     begin
     BitBtn1.Click;
     BitBtn7.SetFocus;
     end;

  if PageControl1.ActivePage=Tab2 then
     begin
     WorkPostsComboBox.ItemIndex:=0;


     DBGrid2.SetFocus;

     if (AppTAbleDate_APP.AsString<>'') then
         begin
         if Form1.SpeedButton3.Visible=True then
            EditOrderBitBtn2.Enabled:=True
         end
     else
         EditOrderBitBtn2.Enabled:=False;

     end;


  if PageControl1.ActivePage=Tab4 then
     begin
     DBGrid20.SetFocus;

     if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


   if AppTAbleMonth.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
      begin
      AppTAbleMonth.Close;
      AppTAbleMonth.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
      AppTAbleMonth.IndexFiles.Clear;
      AppTAbleMonth.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
      end;


   AllPostsComboBox.ItemIndex:=0;

   AppTAbleMonth.Open;
   AppTAbleMonth.IndexName:='DATE_APP1';
   AppTAbleMonth.SetRange(FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'01'+''+'',FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'31'+'"'+'яя');
   AppTAbleMonth.First;


   end;


end;

procedure TAppointmentForm.DBGrid1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin



  Global_X:=X;
  Global_Y:=Y;

  if (abs(Global_X-LastPaint_X)>=15) or (abs(Global_Y-LastPaint_Y)>=15) then
      begin
      dbGrid1.Repaint;
      LastPaint_X:=Global_X;
      LastPaint_Y:=Global_Y;


      if  (RealColumnName='NAIM') then
          dbGrid1.Cursor:=crDefault
      else
          dbGrid1.Cursor:=crHandPoint;

      end;

end;

procedure TAppointmentForm.PopupMenu1Popup(Sender: TObject);
var
   _year: string;
begin

   N1.Visible:=True;
   N2.Visible:=True;


   N1.Caption:='Закрыть запись на '+DateToStr(dsCAlendar1.Date);
   N2.Caption:='Закрыть запись на '+Global_ColumnName;


   AppTAble.Open;
   AppTAble.IndexName:='DATE_APP1';
   AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+''+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+'я'+'яя');
   AppTAble.First;




   if AppTAble.Locate('POSTCODE;TIME1;NAVL_PR',VarArrayOf([DayTAble.FieldByNAme('POSTCODE').ASString,Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),'1']),[])=True then
      begin
      N1.Visible:=FAlse;
      N2.Caption:='Открыть запись на '+Global_ColumnName;
      end
   else //ищем закрытие на весь день
      begin

      if StrToInt(FormatDateTime('YYYY',dsCalendar1.Date))>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=FormatDateTime('YYYY',dsCalendar1.Date);

      if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
         begin
         d_offTable.Close;
         d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
         d_offTable.IndexFiles.Clear;
         d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');

         if FileExists(d_offTable.TableName)=False then
            begin
            Create_d_offTable(_year);
            Create_d_offTable('NY');
            end;
         end;



      d_offTable.Open;
              d_offTable.IndexName:='KOD';
              d_offTable.SetKey;
              d_offTableKOD.AsString:=DayTAble.FieldByNAme('POSTCODE').ASString;
              d_offTableD1.Value:=dsCalendar1.Date;
              if d_offTable.GotoKey then
                 begin
                 N1.Caption:='Открыть запись на '+DateToStr(dsCAlendar1.Date);
                 N2.Visible:=FAlse;
                 end;
      end;



   if AppTAble.Locate('POSTCODE;TIME1;NAVL_PR;IO_UDL',VarArrayOf([DayTAble.FieldByNAme('POSTCODE').ASString,Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),NULL,NULL]),[])=True then
      begin //уже есть записи на ремонт
      N1.Visible:=FAlse;
      N2.Visible:=FAlse;
      end;


   AppTAble.IndexName:='DATE_APP1';
   AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+''+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+'"'+'яя');


end;

procedure TAppointmentForm.DBGrid1CellClick(Column: TColumnEh);
var
   _year: string;
begin

if DetailPanel.Visible=True then exit;


   DetailPanel.Visible:=False;
   Panel6.Visible:=False;

   Rec_IDLabel.Caption:='';

   if (RealColumnName='NAIM') then
      begin
      //первое доступное время
      //Global_ColumnNAme:=Copy(DayTAble.Fields[2].FieldName,2,10);
//      my_messageTime('Внимание!',PChar('Вы можете изменить этот список: '+#13+'Общий справочник -> Ресурсы автосервиса.'),clYellow,10000);

      exit;
      end;


   try
   if DAyTable.FieldByNAme('H'+Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2)).AsString='X' then
      begin
      my_messageTime('Внимание!','Запись закрыта!',clYellow,2000);
      exit;
      end;
   except
   end;

   if DAyTable.FieldByNAme('H'+Copy(Global_ColumnNAme,1,2)).AsString='X' then
      begin
      my_messageTime('Внимание!','Запись закрыта!',clYellow,2000);
      exit;
      end;



   if AppTAble.Locate('POSTCODE;TIME1;NAVL_PR;IO_UDL',VarArrayOf([DayTAble.FieldByNAme('POSTCODE').ASString,Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),NULL,NULL]),[])=True then
      begin //уже есть записи на ремонт

      if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


      if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
         begin
         AppTAbleTime.Close;
         AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
         AppTAbleTime.IndexFiles.Clear;
         AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
         end;

      AppTAbleTime.Open;
      AppTAbleTime.IndexName:='DATE_APP';
      AppTAbleTime.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2));

      Panel7.CAption:='Запись на '+DateToStr(dsCAlendar1.DAte)+' - '+Global_ColumnNAme;

      SelectedTimeLAbel.CAption:=Global_ColumnNAme;

      PAnel6.Left:=(AppointmentForm.ClientWidth div 2) - (PAnel6.Width div 2);
      PAnel6.Top:=(AppointmentForm.ClientHeight div 2) - (PAnel6.Height div 2) -  BitBtn7.Height;
      PAnel6.Visible:=True;
      Panel6.Width:=Tab1.Width-BitBtn1.Width;
      DBGrid3.Width:=PAnel6.Width-DBGrid3.Left*2;
      SpeedButton2.Left:=PAnel7.Width-SpeedButton2.Width-3;



      CreateOrderBitBtn.Enabled:=FAlse;
      if Form1.SpeedButton3.Visible=True then
         CreateOrderBitBtn.Enabled:=True;

      DBGrid3.SetFocus;



      exit;
      end;



//19112021
//есть записи на это время, но с другими минутами и может они скрыты по сетке. Например сетка по 1 часу а запись на 16:30
   try
   if (Trim(DAyTable.FieldByNAme('H'+Copy(Global_ColumnNAme,1,2)).AsString)<>'') then
      begin


      if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


      if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
         begin
         AppTAbleTime.Close;
         AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
         AppTAbleTime.IndexFiles.Clear;
         AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
         end;

      AppTAbleTime.Open;
      AppTAbleTime.IndexName:='DATE_APP';
      AppTAbleTime.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+Copy(Global_ColumnNAme,1,2),FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+Copy(Global_ColumnNAme,1,2)+'59');


      Panel7.CAption:='Запись на '+DateToStr(dsCAlendar1.DAte)+' до '+Global_ColumnNAme;
      SelectedTimeLAbel.CAption:=Global_ColumnNAme;

      PAnel6.Left:=(AppointmentForm.ClientWidth div 2) - (PAnel6.Width div 2);
      PAnel6.Top:=(AppointmentForm.ClientHeight div 2) - (PAnel6.Height div 2) -  BitBtn7.Height;
      PAnel6.Visible:=True;


      CreateOrderBitBtn.Enabled:=FAlse;

      if Form1.SpeedButton3.Visible=True then
         CreateOrderBitBtn.Enabled:=True;

      DBGrid3.SetFocus;

      if AppTAbleTime.RecordCount>0 then exit;

      end;
   except
   end;
//19112021




//есть более ранняя запись перекрывающая это время
   try
   if (Trim(DAyTable.FieldByNAme('H'+Copy(Global_ColumnNAme,1,2)).AsString)<>'') then
      begin


      if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


      if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
         begin
         AppTAbleTime.Close;
         AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
         AppTAbleTime.IndexFiles.Clear;
         AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
         end;

      AppTAbleTime.Open;
      AppTAbleTime.IndexName:='DATE_APP';
      AppTAbleTime.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+DAyTAble.FieldByName('POSTCODE').AsString+Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2));

      Panel7.CAption:='Запись на '+DateToStr(dsCAlendar1.DAte)+' до '+Global_ColumnNAme;
      SelectedTimeLAbel.CAption:=Global_ColumnNAme;

      PAnel6.Left:=(AppointmentForm.ClientWidth div 2) - (PAnel6.Width div 2);
      PAnel6.Top:=(AppointmentForm.ClientHeight div 2) - (PAnel6.Height div 2) -  BitBtn7.Height;
      PAnel6.Visible:=True;


      CreateOrderBitBtn.Enabled:=FAlse;

      if Form1.SpeedButton3.Visible=True then
         CreateOrderBitBtn.Enabled:=True;

      DBGrid3.SetFocus;

      exit;

      end;
   except
   end;



   if (dsCAlendar1.DAte)<Trunc(date) then
       begin
       my_messageTime('Внимание!','Дата уже прошла. Запись невозможна.',clYellow,5000);
       exit;
       end;


   erace_fields;


   if RezhimLAbel.CAption<>'' then
      begin
      OrderNumberMaskEdit.Text:=RezhimLAbel.CAption;

      if ZajavkiForm.Zaj_oQuery.FieldByNAme('N_DOK').AsString=OrderNumberMaskEdit.Text then
         begin
         CarNumberMaskEdit.Text:=ZajavkiForm.MaskEdit5.Text;
         CommentMemo.Text:='Заявка на ремонт уже создана.';
         FindByCarNumberBitBtn.Click;
         if ClientCarsListBox.Items.Count=1 then
            begin
            ClientCarsListBox.ItemIndex:=0;
            ClientCarsListBoxClick(ClientCarsListBox);
            end;

         end;
      end;


   try
   StartTimeMaskEdit.Text:=Global_ColumnNAme;
   StartTimeMaskEdit.ReadOnly:=True;
   except
   end;

   DetailPanel.Left:=(AppointmentForm.ClientWidth div 2) - (DetailPanel.Width div 2);
   DetailPanel.Top:=(AppointmentForm.ClientHeight div 2) - (DetailPanel.Height div 2);



   Label13.Caption:=LAbel2.CAption;
   LAbel18.Caption:=DAyTable.FieldByNAme('NAIM').ASString;

   PageControl1.ActivePage:=Tab10;
   DetailPanel.Visible:=True;

   ClientNameMaskEdit.SetFocus;



end;

procedure TAppointmentForm.N1Click(Sender: TObject);
var
   _year: string;
begin

      if StrToInt(FormatDateTime('YYYY',dsCalendar1.Date))>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=FormatDateTime('YYYY',dsCalendar1.Date);

      if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
         begin
         d_offTable.Close;
         d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
         d_offTable.IndexFiles.Clear;
         d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');

         if FileExists(d_offTable.TableName)=False then
            begin
            Create_d_offTable(_year);
            Create_d_offTable('NY');
            end;
         end;


   if Pos('Открыть',N1.CAption)=1 then
      begin
      d_offTable.Open;
      d_offTable.IndexName:='KOD';
      d_offTable.SetKey;
      d_offTableKOD.AsString:=DayTAble.FieldByNAme('POSTCODE').ASString;
      d_offTableD1.Value:=dsCalendar1.Date;
      if d_offTable.GotoKey then
         d_offTable.Delete;

      BitBtn1.Click;
      exit;
      end;


   //закрыть



      if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
         begin
         AppTAbleTime.Close;
         AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
         AppTAbleTime.IndexFiles.Clear;
         AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
         end;

      AppTAbleTime.Open;
      AppTAbleTime.IndexName:='DATE_APP';

      AppTAbleTime.SetRange(FormatDateTime('YYYYNNDD',dsCalendar1.Date)+DayTAble.FieldByNAme('POSTCODE').ASString+'',FormatDateTime('YYYYNNDD',dsCalendar1.Date)+DayTAble.FieldByNAme('POSTCODE').ASString+'яя');

      if AppTAbleTime.RecordCount>0 then
         if my_dlg('Внимание!','Уже есть запись клиентов на выбранную дату и пост.'+#13+#13+'Продолжить?',clYellow)=False then exit;

      //добавляем запись
    d_offTable.Open;
    d_offTable.AppendRecord([ DayTAble.FieldByNAme('POSTCODE').ASString,
                              dsCalendar1.Date,
                              NULL,
                              kod_operatora,
                              date]);



   BitBtn1.Click;



end;

procedure TAppointmentForm.N2Click(Sender: TObject);
var
   _year: string;
begin //нет записи время

   if Pos('Открыть',N2.CAption)=1 then
      begin

      AppTAble.Open;
      AppTAble.IndexName:='DATE_APP1';
      AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+''+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+'я'+'яя');
      AppTAble.First;

      if AppTAble.Locate('POSTCODE;TIME1;NAVL_PR',VarArrayOf([DayTAble.FieldByNAme('POSTCODE').ASString,Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),'1']),[])=True then
         AppTAble.Delete;

      BitBtn1.Click;

      exit;
      end;


   //закрыть

   if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
      _year:='NY'
   else
      _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


   if AppTAbleForAdd.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
      begin
      AppTAbleForAdd.Close;
      AppTAbleForAdd.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
      AppTAbleForAdd.IndexFiles.Clear;
      AppTAbleForAdd.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
      end;

   AppTAbleForAdd.Open;
   AppTAbleForAdd.IndexName:='DATE_APP';
   AppTAbleForAdd.AppendRecord([dsCAlendar1.Date,
                          Copy(Global_ColumnNAme,1,2)+Copy(Global_ColumnNAme,4,2),
                          NULL,
                          DayTAble.FieldByNAme('POSTCODE').ASString,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          NULL,
                          kod_operatora,//NAVL_IO
                          '1',          //navl_pr
                          NULL,
                          NULL,
                          NULL,
                          kod_operatora,
                          date,
                          FormatDAteTime('nn:ss',now)]);

  AppTAbleForAdd.Close;

  BitBtn1.Click;
end;

procedure TAppointmentForm.N3Click(Sender: TObject);
begin
   if N3.CAption='Восстановить запись' then
      begin
      AppTAble.Edit;
      AppTAbleIO_UDL.ASString:='';
      AppTAbleD_UDL.AsString:='';
      AppTAble.Edit;
      AppTAble.Post;

      exit;
      end;

   DeleteBitBtn.Click;
end;

procedure TAppointmentForm.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumnEh;
  State: TGridDrawState);
var
   String_array: TString_array;
   tel: string;
   Bitmap: TBitmap;
begin
   with DBGrid2.Canvas do
      begin
      if AppTAble.IsEmpty=True then exit;


      if (not (gdSelected in State)) and (AppTableNAVL_PR.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (AppTableIO_UDL.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (not (gdSelected in State)) and (AppTableBRIGHT.AsString<>'') then
          Brush.Color:=Install.PAnel2.Color;



      DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);




      if (Trim(Column.FieldName)='TEL') then
         begin
         fillRect(Rect);
         tel:=UnCryptString(AppTableTEL.AsString,key1c,key2c);
         TextOut(Rect.Left+2,Rect.Top,Copy(tel,1,2)+'('+Copy(tel,3,3)+')'+Copy(tel,6,10));
         end;

      if (Trim(Column.FieldName)='GOS_N') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top+2,UnCryptString(AppTableGOS_N.AsString,key1c,key2c));

         if (AppTableN_ZAJ.AsString<>'') then
             begin
             Bitmap:=BitBtn1.Glyph;
             Bitmap.Transparent:=True;

             DBGrid2.Canvas.Draw(Rect.Right-Bitmap.Width, Rect.Top+2,Bitmap);
             end;


         end;

      if (Trim(Column.FieldName)='FIO') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(AppTableFIO.AsString,DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;


      if (Trim(Column.FieldName)='OBR_KL') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(UnCryptString(AppTableOBR_KL.AsString,key1c,key2c),DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;

      if (Trim(Column.FieldName)='COMM') then
         begin
         fillRect(Rect);
         String_array:=delenie_of_string_for_print(UnCryptString(AppTableCOMM.AsString,key1c,key2c),DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);
         end;





      if (Trim(Column.FieldName)='POSTCODE') then
         begin
         fillRect(Rect);

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='30';
         Form1.Spr00KOD.AsString:=AppTablePOSTCODE.AsString;
         if Form1.Spr00.GotoKey then
            begin
            String_array:=delenie_of_string_for_print(Form1.Spr00NAIM.AsString,DBGrid2.Canvas,Rect.Right-Rect.Left);

            TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

            if String_array.Str2<>'' then
               TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

            end;


         end;

      if (Trim(Column.FieldName)='TIME1') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTableTIME1.AsString,1,2)+':'+Copy(AppTableTIME1.AsString,3,2));
         end;

      if (Trim(Column.FieldName)='TIME2') and (trim(AppTableTIME2.AsString)<>'') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTableTIME2.AsString,1,2)+':'+Copy(AppTableTIME2.AsString,3,2));
         end;


      if (Trim(Column.FieldName)='IO_ADD') then
         begin
         fillRect(Rect);

         font.Size:=7;

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='02';
         Form1.Spr00KOD.AsString:=AppTableIO_ADD.AsString;
         if Form1.Spr00.GotoKey then
            TextOut(Rect.Left+2,Rect.Top,Form1.Spr00NAIM.AsString);

         end;



      end;


end;

procedure TAppointmentForm.SpeedButton1Click(Sender: TObject);
begin
  DetailPanel.Visible:=FAlse;

end;

procedure TAppointmentForm.SaveBitBtnClick(Sender: TObject);
var
  rec_id, postcode, _year: string;
  data_changed: Boolean;
begin

  if (Trim(ClientNameMaskEdit.Text)='') and (Trim(CarNumberMaskEdit.Text)='') then
     begin
     my_messageTime('Внимание!','Не заполнены обязательные поля!',clYellow,5000);
     exit;
     end;


  if (StartTimeMaskEdit.Text='') or (StartTimeMaskEdit.Text='__:__') or (StartTimeMaskEdit.Text=':') or (StartTimeMaskEdit.Text='  :  ') then
     begin
     my_messageTime('Внимание!','Не заполнены обязательные поля!',clYellow,5000);
     exit;
     end;


  RezhimLAbel.CAption:='';


  postcode:='';

  if StoPostComboBox.Visible=False then
     begin
     Form1.Spr00.Open;
     Form1.Spr00.IndexName:='NAIM';
     Form1.Spr00.SetRange('','');
     Form1.Spr00.SetKey;
     Form1.Spr00GR.AsString:='30';
     Form1.Spr00NAIM.AsString:=LAbel18.CAption;
     if Form1.Spr00.GoToKey then
        postcode:=Form1.Spr00KOD.AsString;

     end;




  if (addRecordCheckBox.Checked=True) and (addRecordCheckBox.Visible=True) then
      begin
      Rec_IDLabel.CAption:='';

      if (StoPostComboBox.Visible=True) and (Trim(StoPostComboBox.Text)<>'') then
          begin
          Form1.Spr00.Open;
          Form1.Spr00.IndexName:='NAIM';
          Form1.Spr00.SetRange('','');
          Form1.Spr00.SetKey;
          Form1.Spr00GR.AsString:='30';
          Form1.Spr00NAIM.AsString:=StoPostComboBox.Text;
          if Form1.Spr00.GoToKey then
              postcode:=Form1.Spr00KOD.AsString;

          end;
      end;




  if (Rec_IDLabel.CAption<>'') then
     begin

     postcode:='';

     Form1.Spr00.Open;
     Form1.Spr00.IndexName:='NAIM';
     Form1.Spr00.SetRange('','');
     Form1.Spr00.SetKey;
     Form1.Spr00GR.AsString:='30';
     Form1.Spr00NAIM.AsString:=StoPostComboBox.Text;
     if Form1.Spr00.GoToKey then
        postcode:=Form1.Spr00KOD.AsString;



     AppTAble.IndexName:='DATE_APP';
     AppTAble.SetRange(FormatDateTime('YYYYMMDD',DAteTimePicker1.Date)+postcode+'',FormatDateTime('YYYYMMDD',DAteTimePicker1.Date)+postcode+'яя');
     if AppTAble.Locate('POSTCODE;TIME1;NAVL_PR',VarArrayOf([postcode,Copy(StartTimeMaskEdit.Text,1,2)+Copy(StartTimeMaskEdit.Text,4,2),'1']),[])=True then
        begin
        my_messageTime('Внимание!','Для выбранного поста, даты и времени запись закрыта!',clYellow,5000);
        exit;
        end
     else //ищем закрытие на весь день
     if AppTAble.Locate('POSTCODE;NAVL_PR',VarArrayOf([postcode,'0']),[])=True then
        begin
        my_messageTime('Внимание!','Для выбранного поста, даты и времени запись закрыта!',clYellow,5000);
        exit;
        end;


     AppTAble.IndexName:='REC_ID';
     if AppTAble.FindKey([Rec_IDLabel.CAption])=True then
        begin
        data_changed:=False;

        if trunc(AppTAbleDATE_APP.VAlue)<>trunc(DAteTimePicker1.Date) then
           data_changed:=True;

        if AppTAbleTIME1.AsString<>Copy(StartTimeMaskEdit.Text,1,2)+Copy(StartTimeMaskEdit.Text,4,2) then
           data_changed:=True;


        AppTAble.Edit;

        AppTAbleDATE_APP.VAlue:=DAteTimePicker1.Date;

        AppTAbleTIME1.AsString:=Copy(StartTimeMaskEdit.Text,1,2)+Copy(StartTimeMaskEdit.Text,4,2);

        if (Length(Trim(EndTimeMaskEdit.Text))>4) and (Pos('_',EndTimeMaskEdit.Text)=0) then
           AppTAbleTIME2.AsString:=Copy(EndTimeMaskEdit.Text,1,2)+Copy(EndTimeMaskEdit.Text,4,2);


        if (StoPostComboBox.Visible=True) and (Trim(StoPostComboBox.Text)<>'') then
           begin
           Form1.Spr00.Open;
           Form1.Spr00.IndexName:='NAIM';
           Form1.Spr00.SetKey;
           Form1.Spr00GR.AsString:='30';
           Form1.Spr00NAIM.AsString:=StoPostComboBox.Text;
           if Form1.Spr00.GoToKey then
              AppTAblePOSTCODE.AsString:=Form1.Spr00KOD.AsString;

           end;


        AppTAbleFIO.AsString:=ClientNameMaskEdit.Text;


        AppTAbleGOS_N.AsString:=CryptString(CarNumberMaskEdit.Text,key1c,key2c); //gos_n

        AppTAbleMARKA.AsString:=MarkaMaskEdit.Text; //marka
        AppTAbleMODEL.AsString:=ModelMaskEdit.Text; //

        AppTAbleN_ZAJ.AsString:=OrderNumberMaskEdit.Text; //

        AppTAbleK_KL.AsString:=k_klLAbel.Caption;

        AppTAbleTEL.AsString:=CryptString(telCountryCodeMaskEdit.Text+telCodeMaskEdit.Text+telNumberMaskEdit.Text,key1c,key2c); //tel

        AppTAbleCAR_N_ZAP.AsString:=car_n_zapLabel.Caption;

        AppTAbleOBR_KL.AsString:=CryptString(ReasonMemo.Text,key1c,key2c);
        AppTAbleCOMM.AsString:=CryptString(CommentMemo.Text,key1c,key2c);

        AppTAbleIO_ADD.AsString:=kod_operatora;  //io_add
        AppTAbleD_ADD.VAlue:=date;           //d_add
        AppTAbleT_ADD.AsString:=FormatDAteTime('nn:ss',now);  //t_add

        AppTAble.Edit;
        AppTAble.Post;

        rec_id:=AppTableREC_ID.AsString;

        if Data_changed=True then
           begin
           try
           if ConfirmSMSCheckBox.Checked=True then
              begin
              Navigator_can_show:=False;
              FormForSendSMS_editRec:=True;
              FormForSendSMS_addRec:=False;

              SmsForm.ShowModal;

              FormForSendSMS_editRec:=False;
              FormForSendSMS_addRec:=False;
              Navigator_can_show:=True;
              end;

           except
           end;

          AppTAbleForAdd.Close;
          end;

        end
      else
        my_messageTime('Внимание!','Запись не найдена!',clYellow,3000);
     end
  else
     begin //ADD
     rec_id:=FormatDateTime('DD',date)+FormatDateTime('hhss',now);

     Form1.LastNum.Open;
     Form1.LastNum.IndexName:='MX';
     Form1.LastNum.SetKey;
     Form1.LastNumMX.AsString:='#';
     Form1.LastNumKOD.AsString:='ID';
     if Form1.LastNum.GoToKey=False then
        Form1.LastNum.AppendRecord(['#','ID','0000000','ID записи']);

     if (Form1.LastnumMX.AsString='#') and (Form1.LastnumKOD.AsString='ID') then
         begin
         Form1.Lastnum.edit;
         Form1.LastNumPEREM.AsString:=New_n_zap(Form1.LastNumPEREM.AsString);
         Form1.LastNum.Edit;
         Form1.LastNum.Post;
         end;

     rec_id:=Form1.LastNumPEREM.AsString;


     if StrToInt(FormatDAteTime('YYYY',dsCAlendar1.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',dsCAlendar1.DAte);


     if AppTAbleForAdd.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
        begin
        AppTAbleForAdd.Close;
        AppTAbleForAdd.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
        AppTAbleForAdd.IndexFiles.Clear;
        AppTAbleForAdd.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
        end;



     AppTAbleForAdd.Open;
     AppTAbleForAdd.IndexName:='DATE_APP';
     AppTAbleForAdd.AppendRecord([dsCAlendar1.Date,
                                  Copy(StartTimeMaskEdit.Text,1,2)+Copy(StartTimeMaskEdit.Text,4,2),
                                  NULL,
                                  postcode,
                                  CryptString(CarNumberMaskEdit.Text,key1c,key2c), //gos_n
                                  NULL,                                    //VIN
                                  MarkaMaskEdit.Text,
                                  ModelMaskEdit.Text,
                                  k_klLAbel.Caption,           //k_kl
                                  ClientNameMaskEdit.Text,
                                  CryptString(telCountryCodeMaskEdit.Text+telCodeMaskEdit.Text+telNumberMaskEdit.Text,key1c,key2c), //tel
                                  car_n_zapLabel.Caption,           //car_n_zap
                                  CryptString(ReasonMemo.Text,key1c,key2c),
                                  CryptString(CommentMemo.Text,key1c,key2c),
                                  NULL,           //tab_n
                                  OrderNumberMaskEdit.Text, //n_zaj
                                  NULL,          //NAVL_IO
                                  NULL,          //navl_PR
                                  NULL,          //dop_pr
                                  NULL,          //d_udl
                                  NULL,          //io_udl
                                  kod_operatora,  //io_add
                                  date,           //d_add
                                  FormatDAteTime('nn:ss',now),  //t_add
                                  NULL,
                                  rec_id]);   //rec_id


     if (Length(Trim(EndTimeMaskEdit.Text))>4) and (Pos('_',EndTimeMaskEdit.Text)=0) then
        begin
        AppTAbleForAdd.Edit;
        AppTAbleForAddTIME2.AsString:=Copy(EndTimeMaskEdit.Text,1,2)+Copy(EndTimeMaskEdit.Text,4,2);
        AppTAbleForAdd.Edit;
        AppTAbleForAdd.Post;
        end;


     try
     if ConfirmSMSCheckBox.Checked=True then
        begin

        Navigator_can_show:=False;


        FormForSendSMS_editRec:=False;
        FormForSendSMS_addRec:=True;

        SmsForm.ShowModal;

        FormForSendSMS_editRec:=False;
        FormForSendSMS_addRec:=False;

        Navigator_can_show:=True;


        end;

     except
     end;


     AppTAbleForAdd.Close;


     end;  //add





  BitBtn1.Click;

  try
  if (ZajavkiForm.Showing=True) and (ZajavkiForm.PageControl1.ActivePage=ZajavkiForm.Tab2) then
  if OrderNumberMaskEdit.Text=ZajavkiForm.Zaj_oQuery.FieldByNAme('N_DOK').AsString then
     Zajavki.Select_proc;
  except
  end;


  if PAgeControl1.ActivePAge=Tab1 then
     BitBtn7.SetFocus;


  if PAgeControl1.ActivePAge=Tab2 then
     begin
     DBGrid2.SetFocus;
     AppTAble.Locate('REC_ID',rec_id,[]);
     end;

end;

procedure TAppointmentForm.CancelBitBtnClick(Sender: TObject);
begin
  DetailPanel.Visible:=FAlse;

  if PAgeControl1.ActivePAge=Tab1 then
     BitBtn7.SetFocus;

  if PAgeControl1.ActivePAge=Tab2 then
     DBGrid2.SetFocus;

end;

procedure TAppointmentForm.FormPaint(Sender: TObject);
begin
  if Navigator.Showing=True then
     SetWindowPos(Navigator.Handle,HWND_TOPMOST,Navigator.Left,Navigator.Top,Navigator.Width,Navigator.Height,SWP_NOACTIVATE);

end;

procedure TAppointmentForm.telCountryCodeMaskEditChange(Sender: TObject);
function del_all_space_and_other(str: string): string;
var
   i: integer;
   out_str: string;
begin
   out_str:='';
   for i:=1 to Length(str) do
                   // and (Ord(str[i])<160) выкидывала русские буквы
       if (str[i]<>' ') and (str[i]<>'"') and (str[i]<>#39) and (str[i]<>' ') and (str[i]<>'(') and (str[i]<>')') and (str[i]<>'-') then
          begin   //Это не пробел                              //Это не пробел!!! а какой-то другой символ
          out_str:=out_str+String(str[i]);
          end;

   Result:=out_str;
   del_all_space_and_other:=out_str;
end;
var
  str: string;

begin
  try
  if Length(telCountryCodeMaskEdit.Text)>2 then
     begin
     str:=del_all_space_and_other(telCountryCodeMaskEdit.Text);

     if Copy(telCountryCodeMaskEdit.Text,1,1)<>'+' then
        telCountryCodeMaskEdit.Text:='+'+telCountryCodeMaskEdit.Text;

     telCountryCodeMaskEdit.Text:=Copy(str,1,2);
     telCodeMaskEdit.Text:=Copy(str,3,3);
     telNumberMaskEdit.Text:=Copy(str,6,10); // +7925010-2610
     end;

  if Length(Trim(telCountryCodeMaskEdit.Text))=2 then telCodeMaskEdit.SetFocus;

  except
  end;


end;

procedure TAppointmentForm.telCodeMaskEditChange(Sender: TObject);
begin
  try
  if Length(Trim(telCodeMaskEdit.Text))=3 then telNumberMaskEdit.SetFocus;
  except
  end;


end;

procedure TAppointmentForm.telNumberMaskEditChange(Sender: TObject);
begin
  try
  if Length(Trim(telNumberMaskEdit.Text))=7 then
     begin
     CarNumberMaskEdit.SetFocus;

     if Rec_IDLabel.CAption='' then
     if ccfForm.CheckBox53.Checked=True then ConfirmSMSCheckBox.Checked:=True;

     end;
  except
  end;

end;

procedure TAppointmentForm.telCountryCodeMaskEditClick(Sender: TObject);
begin
  telCountryCodeMaskEdit.SelectAll;

end;

procedure TAppointmentForm.telCodeMaskEditClick(Sender: TObject);
begin
  telCodeMaskEdit.SelectAll;

end;

procedure TAppointmentForm.telNumberMaskEditClick(Sender: TObject);
begin
  telNumberMaskEdit.SelectAll;

end;

procedure TAppointmentForm.NewRecBitBtnClick(Sender: TObject);
begin

  if (dsCAlendar1.DAte)<Trunc(date) then
      begin
      my_messageTime('Внимание!','Дата уже прошла. Запись невозможна.',clYellow,5000);
      exit;
      end;

  Rec_IDLAbel.CAption:='';

  PAnel6.Visible:=False;

  erace_fields;


  if RezhimLAbel.CAption<>'' then
     begin
     OrderNumberMaskEdit.Text:=RezhimLAbel.CAption;

     if ZajavkiForm.Zaj_oQuery.FieldByNAme('N_DOK').AsString=OrderNumberMaskEdit.Text then
        begin
        CarNumberMaskEdit.Text:=ZajavkiForm.MaskEdit5.Text;
        CommentMemo.Text:='Заявка на ремонт уже создана.';
        FindByCarNumberBitBtn.Click;
        if ClientCarsListBox.Items.Count=1 then
           begin
           ClientCarsListBox.ItemIndex:=0;
           ClientCarsListBoxClick(ClientCarsListBox);
           end;

        end;

     end;


  try
  StartTimeMaskEdit.Text:=SelectedTimeLAbel.CAption;
  except
  end;



  DetailPanel.Left:=(AppointmentForm.ClientWidth div 2) - (DetailPanel.Width div 2);
  DetailPanel.Top:=(AppointmentForm.ClientHeight div 2) - (DetailPanel.Height div 2);


  Label13.Caption:=LAbel2.CAption;
  LAbel18.Caption:=DAyTable.FieldByNAme('NAIM').ASString;

  PageControl1.ActivePage:=Tab10;
  DetailPanel.Visible:=True;


end;

procedure TAppointmentForm.DBGrid3DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumnEh;
  State: TGridDrawState);
var
  String_array: TString_array;
  tel: string;
  Bitmap: TBitmap;
begin
 with DBGrid3.Canvas do
      begin
      if AppTableTime.IsEmpty=True then exit;



      if (not (gdSelected in State)) and (AppTableTimeNAVL_PR.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (not (gdSelected in State)) and (AppTableTimeIO_UDL.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (not (gdSelected in State)) and (AppTableTimeBRIGHT.AsString<>'') then
          Brush.Color:=Install.PAnel2.Color;


      DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);


      if (Trim(Column.FieldName)='TEL') then
         begin
         fillRect(Rect);
         tel:=UnCryptString(AppTableTimeTEL.AsString,key1c,key2c);
         TextOut(Rect.Left+2,Rect.Top,Copy(tel,1,2)+'('+Copy(tel,3,3)+')'+Copy(tel,6,10));
         end;

      if (Trim(Column.FieldName)='FIO') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(AppTableTimeFIO.AsString,DBGrid3.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;



      if (Trim(Column.FieldName)='GOS_N') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top+2,UnCryptString(AppTableTimeGOS_N.AsString,key1c,key2c));


         if (AppTableTimeN_ZAJ.AsString<>'') then
             begin
             Bitmap:=BitBtn1.Glyph;
             Bitmap.Transparent:=True;

             DBGrid3.Canvas.Draw(Rect.Right-Bitmap.Width, Rect.Top+2,Bitmap);
             end;


         end;

      if (Trim(Column.FieldName)='OBR_KL') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(UnCryptString(AppTableTimeOBR_KL.AsString,key1c,key2c),DBGrid3.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;

      if (Trim(Column.FieldName)='COMM') then
         begin
         fillRect(Rect);
         String_array:=delenie_of_string_for_print(UnCryptString(AppTableTimeCOMM.AsString,key1c,key2c),DBGrid3.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);
         end;


      if (Trim(Column.FieldName)='POSTCODE') then
         begin
         fillRect(Rect);

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='30';
         Form1.Spr00KOD.AsString:=AppTableTimePOSTCODE.AsString;
         if Form1.Spr00.GotoKey then
            begin
            String_array:=delenie_of_string_for_print(Form1.Spr00NAIM.AsString,DBGrid3.Canvas,Rect.Right-Rect.Left);

            TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

            if String_array.Str2<>'' then
               TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

            end;

         end;

      if (Trim(Column.FieldName)='TIME1') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTableTimeTIME1.AsString,1,2)+':'+Copy(AppTableTimeTIME1.AsString,3,2));
         end;


      if (Trim(Column.FieldName)='TIME2') and (trim(AppTableTimeTIME2.AsString)<>'') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTableTimeTIME2.AsString,1,2)+':'+Copy(AppTableTimeTIME2.AsString,3,2));
         end;

      if (Trim(Column.FieldName)='IO_ADD') then
         begin
         fillRect(Rect);

         font.Size:=7;

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='02';
         Form1.Spr00KOD.AsString:=AppTableTimeIO_ADD.AsString;
         if Form1.Spr00.GotoKey then
            TextOut(Rect.Left+2,Rect.Top,Form1.Spr00NAIM.AsString);

         end;



      end;


end;

procedure TAppointmentForm.SpeedButton2Click(Sender: TObject);
begin
  PAnel6.Visible:=False;
  BitBtn1.Click;

end;

procedure TAppointmentForm.DeleteBitBtnClick(Sender: TObject);
begin
  if AppTAbleDATE_APP.AsString='' then exit;

  if my_dlg('Внимание!','Вы действительно хотите удалить запись?',clYellow)=FAlse then exit;

  if AppTAbleNAVL_PR.AsString<>'' then
     AppTAble.Delete
  else
     begin
     AppTAble.Edit;
     AppTAbleIO_UDL.ASString:=kod_operatora;
     AppTAbleD_UDL.VAlue:=date;
     AppTAble.Edit;
     AppTAble.Post;
     end;

end;

procedure TAppointmentForm.PopupMenu2Popup(Sender: TObject);
begin
   if AppTAbleDATE_APP.AsString='' then
      exit;

   if AppTAbleIO_UDL.AsString='' then
      N3.CAption:='Удалить запись'
   else
      N3.CAption:='Восстановить запись';


   N22.Visible:=FAlse;
   if (Length(AppTAbleN_ZAJ.ASString)=9) and (Copy(AppTAbleN_ZAJ.ASString,4,1)='J') then
       begin
       N22.Caption:='-> Заявка на ремонт '+AppTAbleN_ZAJ.ASString;
       N22.Visible:=True;
       end;

end;

procedure TAppointmentForm.N5Click(Sender: TObject);
begin
  EditOrderBitBtn2.Click;
end;

procedure TAppointmentForm.EditOrderBitBtn2Click(Sender: TObject);
begin
  if AppTAbleIO_UDL.ASString<>'' then
     begin
     my_messageTime('Внимание!','Запись удалена!'+#13+'Сначала восстановите запись по правой кнопке мыши.',clYellow,5000);
     exit;
     end;

  if AppTAblePOSTCODE.ASString='' then exit;

  fill_form(AppTAble);

  DetailPanel.Visible:=True;



end;

procedure TAppointmentForm.N6Click(Sender: TObject);
begin
  if CreateOrderBitBtn.Enabled=True then
     CreateOrderBitBtn.Click;
end;

procedure TAppointmentForm.N8Click(Sender: TObject);
begin
  NewRecBitBtn.Click;
end;

procedure TAppointmentForm.N9Click(Sender: TObject);
begin
  EditRecBitBtn.Click;
end;

procedure TAppointmentForm.N10Click(Sender: TObject);
begin
  CreateOrderBitBtn.Click;
end;

procedure TAppointmentForm.N11Click(Sender: TObject);
begin
  DelRecBitBtn.Click;
end;

procedure TAppointmentForm.TelMaskEditCountryCodeChange(Sender: TObject);
function del_all_space_and_other(str: string): string;
var
   i: integer;
   out_str: string;
begin
   out_str:='';
   for i:=1 to Length(str) do
       if (str[i]<>' ') and (str[i]<>'"') and (str[i]<>#39) and (str[i]<>' ') and (str[i]<>'(') and (str[i]<>')') and (str[i]<>'-') then
          begin   //Это не пробел                              //Это не пробел!!!
          out_str:=out_str+String(str[i]);
          end;

   Result:=out_str;
   del_all_space_and_other:=out_str;
end;
var
  str: string;

begin
  try
  if Length(TelMaskEditCountryCode.Text)>2 then
     begin
     str:=del_all_space_and_other(TelMaskEditCountryCode.Text);

     if Copy(TelMaskEditCountryCode.Text,1,1)<>'+' then
        TelMaskEditCountryCode.Text:='+'+TelMaskEditCountryCode.Text;

     TelMaskEditCountryCode.Text:=Copy(str,1,2);
     TelMaskEditCode.Text:=Copy(str,3,3);
     TelMaskEditNumber.Text:=Copy(str,6,10);
     end;

  if Length(Trim(TelMaskEditCountryCode.Text))=2 then TelMaskEditCode.SetFocus;

  except
  end;



end;

procedure TAppointmentForm.TelMaskEditCountryCodeClick(Sender: TObject);
begin
  TelMaskEditCountryCode.SelectAll;

end;

procedure TAppointmentForm.TelMaskEditCodeChange(Sender: TObject);
begin
  try
  if Length(Trim(TelMaskEditCode.Text))=3 then TelMaskEditNumber.SetFocus;
  except
  end;

end;

procedure TAppointmentForm.TelMaskEditCodeClick(Sender: TObject);
begin
  TelMaskEditCode.SelectAll;

end;

procedure TAppointmentForm.TelMaskEditNumberClick(Sender: TObject);
begin
  TelMaskEditNumber.SelectAll;

end;

procedure TAppointmentForm.TelMaskEditNumberChange(Sender: TObject);
begin
  try
  if Length(Trim(TelMaskEditNumber.Text))=7 then find_by_tel_proc;
  except
  end;

end;

procedure TAppointmentForm.CarNumberFindEditClick(Sender: TObject);
begin
if install.CheckBox9.Checked=True then
  LoadKeyboardLayout('00000419',KLF_ACTIVATE);

  CarNumberFindEdit.SelectAll;
end;

procedure TAppointmentForm.CarNumberFindEditKeyPress(Sender: TObject;
  var Key: Char);
begin
   if Key=#13 then find_by_gosn_proc;
end;

procedure TAppointmentForm.EditRecBitBtnClick(Sender: TObject);
begin
  if AppTAbleTimeIO_UDL.ASString<>'' then
     begin
     my_messageTime('Внимание!','Запись удалена!'+#13+'Сначала восстановите запись по правой кнопке мыши.',clYellow,5000);
     exit;
     end;

  if AppTAbleTimePOSTCODE.ASString='' then exit;


  fill_form(AppTAbleTime);

  DetailPanel.Visible:=True;

end;

procedure TAppointmentForm.DelRecBitBtnClick(Sender: TObject);
begin
  if AppTAbleTimeDATE_APP.AsString='' then exit;

  if my_dlg('Внимание!','Вы действительно хотите удалить запись?',clYellow)=FAlse then exit;

  if AppTAbleTimeNAVL_PR.AsString<>'' then
     AppTAbleTime.Delete
  else
     begin
     AppTAbleTime.Edit;
     AppTAbleTimeIO_UDL.ASString:=kod_operatora;
     AppTAbleTimeD_UDL.VAlue:=date;
     AppTAbleTime.Edit;
     AppTAbleTime.Post;
     end;

  Panel6.Visible:=False;
  BitBtn1.Click;



end;

procedure TAppointmentForm.DetailPanelExit(Sender: TObject);
begin
  PAnel9.Visible:=FAlse;
end;

procedure TAppointmentForm.DBGrid3DblClick(Sender: TObject);
begin
  EditRecBitBtn.Click;
end;

procedure TAppointmentForm.DBGrid2DblClick(Sender: TObject);
begin
   EditOrderBitBtn2.Click;
end;

procedure TAppointmentForm.ClientSearchBitBtnClick(Sender: TObject);
var
  SelectKLient: function(_manager,_k_kl,_kl_naim,_email,_kod_operatora,_Path_base: PChar; Form_Color,Panel_Color,Font_Color: TColor; _Only_svoi,_KL_SpravDostup, SHowNewKLButton, group_select: Boolean): PChar; stdcall;

  LibHandle: THandle;
  handle: integer;
  fun_result, tel,str: string;

begin

     @SelectKLient:=nil;
     LibHandle:=LoadLibrary('sprav.dll');
     if LibHandle>=32 then
        begin
        @SelectKLient:=GetProcAddress(LibHandle,'SelectKLient');
        if @SelectKLient<>nil then
           begin
            if Users.CheckBox7.Checked=True then
               fun_result:=StrPas(SelectKLient('','','','',PChar(kod_operatora),PChar(Path_base),AppointmentForm.Color,AppointmentForm.DBGrid1.Color,Label2.Font.Color,True,False,false,false))
            else
               fun_result:=StrPas(SelectKLient('','','','',PChar(kod_operatora),PChar(Path_base),AppointmentForm.Color,AppointmentForm.DBGrid1.Color,Label2.Font.Color,False,False,false,false));
           end
        else ShowMessage('Библиотека sprav.dll не найдена...');
        end
        else my_message('Внимание!','Файл sprav.dll не найден.',clYellow);


    if (Trim(fun_result)<>'') and (fun_result<>'ERASE') then
       begin
       k_klLAbel.CAption:=Copy(fun_result,1,6);

       if trim(k_klLabel.Caption)<>'' then
          ClientSpravBitBtn.Visible:=True
       else
          ClientSpravBitBtn.Visible:=False;


       ClientNameMaskEdit.Text:=Copy(fun_result,7,100);
       end;


       if Trim(fun_result)='ERASE' then
          begin
          k_klLAbel.CAption:='';
          ClientNameMaskEdit.Text:='';
          ClientSpravBitBtn.Visible:=False;
          end
       else
       if (Copy(k_klLAbel.CAption,2,1)='F') or (Copy(k_klLAbel.CAption,2,1)='U') then
          zapolnenie_kl_proc;



     FreeLibrary(LibHandle);

     car_from_sprav_proc;
end;

procedure TAppointmentForm.ClientSpravBitBtnClick(Sender: TObject);
begin
  password_bar:=False;


  if Copy(k_klLabel.CAption,2,1)='F' then
     begin
     if Copy(UncryptString(Form1.Label5.Caption,key1v,key2v),17,1)<>'T' then //нет оптовой торговли-2
        Clients_proc('xx',
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,
                     DBGrid1.Color,
                     Label2.Font.Color,
                     'Izm_fiz',
                     PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible)
     else
     if Users.CheckBox7.Checked=False then
        Clients_proc('*',
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,
                     DBGrid1.Color,
                     Label2.Font.Color,
                     'Izm_fiz',
                     PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible)
     else
        Clients_proc(PChar(kod_operatora),
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,
                     DBGrid1.Color,
                     Label2.Font.Color,
                     'Izm_fiz',
                     PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible);
     end
  else
  if Copy(k_klLabel.CAption,2,1)='U' then
     begin
     if Copy(UncryptString(Form1.Label5.Caption,key1v,key2v),17,1)<>'T' then //нет оптовой торговли-2
        Clients_proc('xx',
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,DBGrid1.Color,
                     Label2.Font.Color,
                     'Izm_ur',PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible)
     else
     if Users.CheckBox7.Checked=False then
        Clients_proc('*',
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,
                     DBGrid1.Color,
                     Label2.Font.Color,
                     'Izm_ur',
                     PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible)
     else
        Clients_proc(PChar(kod_operatora),
                     PChar(Mx),
                     PChar(kod_operatora+users.MaskEdit100.Text),
                     PChar(Path_base),
                     Form1.Color,
                     DBGrid1.Color,
                     Label8.Font.Color,
                     'Izm_ur',
                     PChar(name),
                     PChar(k_klLabel.CAption),
                     Form1.N35.Visible,
                     Form1.N19.Visible,
                     Users.CheckBox14.Checked,
                     Form1.N41.Visible);
     end;

    Form1.Timer3.Enabled:=False;
    Form1.Timer3.Enabled:=True;
    password_bar:=True;

end;

procedure TAppointmentForm.SpeedButton4Click(Sender: TObject);
begin
  Panel9.Visible:=FAlse;
end;

procedure TAppointmentForm.ClientCarsListBoxMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ClientCarsListBox.ItemIndex := ClientCarsListBox.ItemAtPos(Point(X, Y), True);

end;

procedure TAppointmentForm.ClientCarsListBoxClick(Sender: TObject);
var
  i: integer;
begin


try
if Trim(Copy(ClientCarsListBox.Items[ClientCarsListBox.ItemIndex],1,9))='' then
   begin
   PAnel9.Visible:=False;
   exit;
   end;
except
      PAnel9.Visible:=False;
      exit;
end;

   if Pos('автомобиль',PAnel10.CAption)>0 then
      begin

      carsQuery.SQL.Clear;
      carsQuery.SQL.Add('SELECT * FROM "'+path_base+'sprav\cars.dbf" WHERE (N_ZAP LIKE :N_ZAP)');
      carsQuery.ParamByName('N_ZAP').AsString:=Trim(Copy(ClientCarsListBox.Items[ClientCarsListBox.ItemIndex],1,9));
      carsQuery.Open;

      CarNumberMaskEdit.Text:=UnCryptString(CarsQUERY.FieldByName('GOS_N').AsString,key1a,key2a);

      CAr_n_zapLAbel.Caption:=CarsQUERY.FieldByName('N_ZAP').AsString;

      if trim(car_n_zapLAbel.CAption)<>'' then
         begin
         BitBtn15.Visible:=True;
         BitBtn15.Enabled:=True;
         end;

      Form1.Spr01.Open;
      Form1.Spr01.IndexName:='KOD';
      Form1.Spr01.SetRange('','');
      Form1.Spr01.SetKey;
      Form1.Spr01GR.AsString:='11';
      Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,2);
      if Form1.Spr01.GoToKey then
         MarkaMaskEdit.Text:=Trim(Form1.Spr01NAIM.AsString);

      Form1.Spr01.SetKey;
      Form1.Spr01GR.AsString:='12';
      Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,4);
      if Form1.Spr01.GoToKey then
         ModelMaskEdit.Text:=Trim(Form1.Spr01NAIM.AsString);


      Form1.Spr01.SetKey;
      Form1.Spr01GR.AsString:='13';
      Form1.Spr01KOD.AsString:=Copy(CarsQUERY.FieldByName('AUTO').AsString,1,6);
      if Form1.Spr01.GoToKey then
         ModelMaskEdit.Text:=TrimRight(ModelMaskEdit.Text)+', '+Trim(Form1.Spr01NAIM.AsString);


      i:=Pos('  ',ModelMaskEdit.Text);
      while i>0 do
            begin
            ModelMaskEdit.Text:=Copy(ModelMaskEdit.Text,1,i-1)+' '+Copy(ModelMaskEdit.Text,i+2,100);
            i:=Pos('  ',ModelMaskEdit.Text);
            end;


      PAnel9.Visible:=False;
      end
   else
   if Pos('клиент',PAnel10.CAption)>0 then
      begin
      k_klLAbel.CAption:=Copy(ClientCarsListBox.Items[ClientCarsListBox.ItemIndex],107,6);


      zapolnenie_kl_proc;
      PAnel9.Visible:=False;

      car_from_sprav_proc;

      if trim(k_klLabel.Caption)<>'' then
         ClientSpravBitBtn.Visible:=True
      else
         ClientSpravBitBtn.Visible:=False;

      end;

end;

procedure TAppointmentForm.SearchByCardBitBtnClick(Sender: TObject);
begin
   if Length(trim(CardNumberMaskEdit.Text))<3 then exit;

   ClientCarsListBox.Items.Clear;

   SqlQuery1.SQL.Clear;
   SqlQuery1.SQL.Add('SELECT * FROM "'+path_base+'sprav\cl_f.dbf" WHERE (DISCONT LIKE :DISCONT)');
   SqlQuery1.ParamByName('DISCONT').AsString:='%'+CryptString(CardNumberMaskEdit.Text,key1c,key2c)+'%';
   SqlQuery1.Open;
   SqlQuery1.First;
   while SqlQuery1.Eof=FAlse do
         begin
         ClientCarsListBox.Items.Add(Copy(UnCryptString(SqlQuery1.FieldByName('NAME').ASString,key1c,key2c)+
         '                                                                                                              ',1,106)+SqlQuery1.FieldByName('K_KL').AsString);

         SqlQuery1.Next;
         end;


   SqlQuery1.SQL.Clear;
   SqlQuery1.SQL.Add('SELECT * FROM "'+path_base+'sprav\cl_u.dbf" WHERE (DISCONT LIKE :DISCONT)');
   SqlQuery1.ParamByName('DISCONT').AsString:='%'+CryptString(CardNumberMaskEdit.Text,key1c,key2c)+'%';
   SqlQuery1.Open;
   SqlQuery1.First;
   while SqlQuery1.Eof=FAlse do
         begin
         ClientCarsListBox.Items.Add(Copy(UnCryptString(SqlQuery1.FieldByName('ORG').ASString,key1c,key2c)+
         '                                                                                                              ',1,106)+SqlQuery1.FieldByName('K_KL').AsString);

         SqlQuery1.Next;
         end;




   if ClientCarsListBox.Items.Count=1 then
      begin
      k_klLAbel.CAption:=Copy(ClientCarsListBox.Items[0],107,6);
      zapolnenie_kl_proc;
      car_from_sprav_proc;

      if trim(k_klLabel.Caption)<>'' then
         ClientSpravBitBtn.Visible:=True
      else
         ClientSpravBitBtn.Visible:=False;
      end
   else
   if ClientCarsListBox.Items.Count>1 then
      begin
      PAnel10.Caption:='Уточните клиента';

      PAnel9.Visible:=True;
      end
   else
      my_messageTime('Внимание!','Клиент не найден.',clYellow,3000);


end;

procedure TAppointmentForm.CardNumberMaskEditClick(Sender: TObject);
begin
   CardNumberMaskEdit.SelectAll;
end;

procedure TAppointmentForm.CardNumberMaskEditKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key=#13 then SearchByCardBitBtn.Click;
end;

procedure TAppointmentForm.DateTimePicker2Change(Sender: TObject);
begin
  if Trunc(DAteTimePicker3.Date)<Trunc(DAteTimePicker2.Date) then
     DAteTimePicker3.Date:=DAteTimePicker2.Date;
end;

procedure TAppointmentForm.StartTimeMaskEditClick(Sender: TObject);
begin
  StartTimeMaskEdit.SelectAll;
end;

procedure TAppointmentForm.EndTimeMaskEditClick(Sender: TObject);
begin
  EndTimeMaskEdit.SelectAll;
end;

procedure TAppointmentForm.N14Click(Sender: TObject);
begin
  d_offTableSTO.Delete;

  fill_sto_days_off(dsCalendar1,'STO');

end;

procedure TAppointmentForm.SpeedButton5Click(Sender: TObject);
var
  i: integer;
  _year: string;
begin

  if Trunc(DAteTimePicker3.Date)<Trunc(DateTimePicker2.Date) then
     begin
     my_messageTime('Внимание!','Неверный диапазон дат.',clYellow,3000);
     exit;
     end;

  if StrToInt(FormatDAteTime('YYYY',DateTimePicker2.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
     _year:='NY'
  else
     _year:=FormatDAteTime('YYYY',DateTimePicker2.DAte);


  if d_offTableSTO.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
     begin
     Create_d_offTable(_year);

     d_offTableSTO.Close;
     d_offTableSTO.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
     d_offTableSTO.IndexFiles.Clear;
     d_offTableSTO.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');
     end;

  d_offTableSTO.Open;
  d_offTableSTO.IndexName:='KOD';
  d_offTableSTO.SetRange('STO   '+'','STO   '+'яя');
  if d_offTableSTO.Locate('D1',FormatDAteTime('YYYYMMDD',DateTimePicker2.DAte),[])=True then
     begin
     my_messageTime('Внимание!','Запись уже есть.'+#13+'Удалите имеющуюся.',clYellow,3000);
     exit;
     end;

  d_offTableSTO.AppendRecord(['STO',
                             DateTimePicker2.Date,
                             DateTimePicker3.Date,
                             NULL,
                             NULL,
                             NULL,
                             kod_operatora,
                             date]);


  if Trunc(DAteTimePicker3.Date)>Trunc(DateTimePicker2.Date) then
     begin
     if StrToInt(FormatDAteTime('YYYY',DateTimePicker3.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
        _year:='NY'
     else
        _year:=FormatDAteTime('YYYY',DateTimePicker3.DAte);



    if d_offTableSTO.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
       begin
       Create_d_offTable(_year);

       d_offTableSTO.Close;
       d_offTableSTO.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
       d_offTableSTO.IndexFiles.Clear;
       d_offTableSTO.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');
       end;

    d_offTableSTO.Open;
    d_offTableSTO.IndexName:='KOD';
    d_offTableSTO.SetRange('STO   '+'','STO   '+'яя');
    if d_offTableSTO.Locate('D1',FormatDAteTime('YYYYMMDD',DateTimePicker3.DAte),[])=True then
       begin
       my_messageTime('Внимание!',PChar('Запись уже есть в '+FormatDAteTime('YYYY',DateTimePicker3.DAte)+' году.'+#13+'Удалите имеющуюся.'),clYellow,3000);
       exit;
       end;

    d_offTableSTO.AppendRecord(['STO',
                                StrToDAte('01.01.'+FormatDAteTime('YYYY',DateTimePicker3.DAte)),
                                DateTimePicker3.Date,
                                NULL,
                                NULL,
                                NULL,
                                kod_operatora,
                                date]);



    end;


   dsCAlendar1.Date:=DateTimePicker2.Date;
   fill_sto_days_off(dsCalendar1,'STO');

end;

procedure TAppointmentForm.CarNumberMaskEditClick(Sender: TObject);
begin
  if install.CheckBox9.Checked=True then
     LoadKeyboardLayout('00000419',KLF_ACTIVATE);

   CarNumberMaskEdit.SelectAll;
end;

procedure TAppointmentForm.ClientNameMaskEditClick(Sender: TObject);
begin
  if install.CheckBox9.Checked=True then
     LoadKeyboardLayout('00000419',KLF_ACTIVATE);
end;

procedure TAppointmentForm.FindByCarNumberBitBtnClick(Sender: TObject);
var
  klient_name, auto: string;
  i: integer;
begin
        SqlQuery1.SQL.Clear;
        SqlQuery1.SQL.Add('SELECT * FROM "'+path_base+'sprav\cars.dbf" WHERE (GOS_N LIKE :GOS_N1) OR (GOS_N LIKE :GOS_N2)');
        SqlQuery1.ParamByName('GOS_N1').AsString:='%'+CryptString(CarNumberMaskEdit.Text,key1a,key2a)+'%';
        SqlQuery1.ParamByName('GOS_N2').AsString:='%'+CryptString(LowerCaseRus(LowerCase(CarNumberMaskEdit.Text)),key1a,key2a)+'%';
        SqlQuery1.Open;

        if SqlQuery1.RecordCount=0 then
           begin
           SqlQuery1.SQL.Clear;
           SqlQuery1.SQL.Add('SELECT * FROM "'+path_base+'sprav\cars.dbf" WHERE (GOS_N LIKE :GOS_N)');
           SqlQuery1.ParamByName('GOS_N').AsString:='%'+CryptString(remove_all_except_digits(CarNumberMaskEdit.Text),key1a,key2a)+'%';
           SqlQuery1.Open;
           end;


        SqlQuery1.First;

        if SqlQuery1.RecordCount>=1 then
            begin


            ClientCarsListBox.Items.Clear;
            SqlQuery1.First;
             while SqlQuery1.Eof=FAlse do
                   begin
                   auto:='';

                   Form1.Spr01.Open;
                   Form1.Spr01.IndexName:='KOD';
                   Form1.Spr01.SetKey;

                   if Length(trim(SqlQuery1.FieldByName('AUTO').AsString))>=4 then
                      begin
                      Form1.Spr01GR.AsString:='12';
                      Form1.Spr01KOD.AsString:=Copy(SqlQuery1.FieldByName('AUTO').AsString,1,4);
                      end
                   else
                      begin
                      Form1.Spr01GR.AsString:='11';
                      Form1.Spr01KOD.AsString:=Copy(SqlQuery1.FieldByName('AUTO').AsString,1,2);
                      end;


                   if Form1.Spr01.GoToKey then
                      begin
                      auto:=Form1.Spr01NAIM.AsString;

                      i:=Pos('  ',auto);
                      while i>0 do
                            begin
                            auto:=Copy(auto,1,i-1)+' '+Copy(auto,i+2,100);
                            i:=Pos('  ',auto);
                            end;

                      end;


                   klient_name:='';

                   if Copy(SqlQuery1.FieldByName('K_KL').AsString,2,1)='F' then
                      begin
                      Form1.cl_f.Open;
                      Form1.cl_f.IndexName:='K_KL';
                      if Form1.cl_f.FindKey([SqlQuery1.FieldByName('K_KL').AsString]) then
                         klient_name:=UnCryptString(Form1.cl_fNAME.AsString,key1c,key2c);

                      end;

                   if Copy(SqlQuery1.FieldByName('K_KL').AsString,2,1)='U' then
                      begin
                      Form1.cl_u.Open;
                      Form1.cl_u.IndexName:='K_KL';
                      if Form1.cl_u.FindKey([SqlQuery1.FieldByName('K_KL').AsString]) then
                         klient_name:=UnCryptString(Form1.cl_uORG.AsString,key1c,key2c);

                      end;


                   ClientCarsListBox.Items.Add('Гос.№: '+Copy(UnCryptString(SqlQuery1.FieldByName('GOS_N').AsString,key1a,key2a)+'          ',1,10)+' - '+Copy(auto+'                    ',1,20)+' - '+Copy(klient_name+'                                                            ',1,60)+' - '+SqlQuery1.FieldByName('K_KL').AsString);

                   SqlQuery1.Next;
                   end;


            PAnel10.Caption:='Уточните клиента';
            PAnel9.Visible:=True;
            end
        else
            my_messageTime('Внимание!','Клиент не найден.',clYellow,3000);

end;

procedure TAppointmentForm.CarNumberMaskEditKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (key=#13) and (ClientNameMaskEdit.Text='') and (k_klLAbel.CAption='') then FindByCarNumberBitBtn.Click;

end;

procedure TAppointmentForm.CarNumberMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DOWN then MarkaMaskEdit.SetFocus;
end;

procedure TAppointmentForm.MarkaMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_UP then CarNumberMaskEdit.SetFocus;
  if Key=VK_DOWN then ModelMaskEdit.SetFocus;
end;

procedure TAppointmentForm.ModelMaskEditKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_UP then MarkaMaskEdit.SetFocus;
  if Key=VK_DOWN then ReasonMemo.SetFocus;

end;

procedure TAppointmentForm.BitBtn15Click(Sender: TObject);
var
  Rezhim: string;

begin
  if Copy(k_klLAbel.CAption,2,1)='U' then Rezhim:='Client_ur'
      else Rezhim:='Client_fiz';


  Cars_proc(PChar(kod_operatora),
            PChar(Path_base),
            Form1.Color,
            DBGrid1.Color,
            Label2.Font.Color,
            PChar(Rezhim),
            PChar(k_klLAbel.CAption),
            Form1.N35.Visible,
            PChar(Car_n_zapLAbel.CAption),
            '');

end;

procedure TAppointmentForm.dsCalendar2DateSelect(Sender: TObject);
var
   _year: string;
begin

EXIT; //not use

   if StrToInt(FormatDAteTime('YYYY',dsCalendar2.DAte))>StrToInt(FormatDAteTime('YYYY',date)) then
      _year:='NY'
   else
      _year:=FormatDAteTime('YYYY',dsCalendar2.DAte);


   if d_offTableSTO.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
      begin
      Create_d_offTable(_year);

      d_offTableSTO.Close;
      d_offTableSTO.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
      d_offTableSTO.IndexFiles.Clear;
      d_offTableSTO.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');
      end;

    d_offTableSTO.Open;
    d_offTableSTO.IndexName:='KOD';
    d_offTableSTO.SetRange('STO   '+'','STO   '+'яя');
    if d_offTableSTO.Locate('D1',FormatDAteTime('YYYYMMDD',dsCalendar2.Date),[])=True then
       d_offTableSTO.Delete
    else
       d_offTableSTO.AppendRecord(['STO',
                             dsCalendar2.Date,
                             NULL,
                             NULL,
                             NULL,
                             NULL,
                             kod_operatora,
                             date]);


    dsCalendar2.DAte:=dsCalendar2.DAte+1;
    fill_sto_days_off(dsCalendar2,'STO');


end;

procedure TAppointmentForm.ResourceItemsComboBoxChange(Sender: TObject);
begin

PostCodeLAbel.CAption:='';

if ResourceItemsComboBox.ItemIndex=1 then
   begin    //Персонал СТО
   PostCodeLAbel.CAption:='STAFF';
   DBGrid7.Visible:=FAlse;
   DBGrid9.Visible:=True;
   DBGrid9.DataSource:=DAtaSource8;
   MonthsComboBoxChange(MonthsComboBox);
   end
else
if ResourceItemsComboBox.ItemIndex=2 then
   begin    //Нерабочие дни СТО
   PostCodeLAbel.CAption:='STO';
   DBGrid7.Visible:=True;
   DBGrid9.Visible:=False;
   MonthsComboBoxChange(MonthsComboBox);
   end
else
if ResourceItemsComboBox.ItemIndex=0 then
   begin    //все ресурсы СТО
   PostCodeLAbel.CAption:='STO';
   DBGrid7.Visible:=False;
   DBGrid9.Visible:=True;
   DBGrid9.DataSource:=DAtaSource6;
   MonthsComboBoxChange(MonthsComboBox);
   end
else
   begin
   Form1.Spr00.Open;
   Form1.Spr00.SetRange('','');
   Form1.Spr00.IndexName:='NAIM';
   Form1.Spr00.SetKey;
   Form1.Spr00GR.AsString:='30';
   Form1.Spr00NAIM.AsString:=ResourceItemsComboBox.Text;
   if Form1.Spr00.GotoKey then
      PostCodeLAbel.CAption:=Form1.Spr00KOD.AsString;

   DBGrid7.Visible:=True;
   DBGrid9.Visible:=False;
   MonthsComboBoxChange(MonthsComboBox);
   end;

 dbGrid7.Repaint;
end;

procedure TAppointmentForm.MonthsComboBoxChange(Sender: TObject);
begin
 CreateMonthTableProc(MonthsComboBox.ItemIndex+1,StrToInt(YearComboBox.Text));

 CreateMonthTableProcPosts(MonthsComboBox.ItemIndex+1,StrToInt(YearComboBox.Text));

 CreateMonthTableProcStaff(MonthsComboBox.ItemIndex+1,StrToInt(YearComboBox.Text));

 try
 dbGrid7.SetFocus;
 except

 try
 dbGrid9.SetFocus;
 except
 end;


 end;
end;

procedure TAppointmentForm.YearComboBoxChange(Sender: TObject);
begin
   MonthsComboBoxChange(MonthsComboBox);
end;

procedure TAppointmentForm.DBGrid7DrawDataCell(Sender: TObject;
  const Rect: TRect; Field: TField; State: TGridDrawState);
var
 _year, month, day: string;
begin
 with DBGrid7.Canvas do
      begin


      if (Field.AsString<>'') then
          begin
          brush.Color:=$0097FF97;


          try
          if PostCodeLAbel.Caption<>'' then
             begin

             if StrToInt(YearComboBox.Text)>StrToInt(FormatDAteTime('YYYY',date)) then
                _year:='NY'
             else
                _year:=YearComboBox.Text;

             if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
                begin
                d_offTable.Close;
                d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
                d_offTable.IndexFiles.Clear;
                d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');

                if FileExists(d_offTable.TableName)=False then
                   begin
                   Create_d_offTable(_year);
                   Create_d_offTable('NY');
                   end;
                end;

             if (MonthsComboBox.ItemIndex+1)>9 then
                 month:=IntToStr(MonthsComboBox.ItemIndex+1)
             else
                 month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

             day:=Field.Value;
             if Length(day)=1 then
                day:='0'+day;


             d_offTable.Open;
             d_offTable.IndexName:='KOD';
             d_offTable.SetKey;
             d_offTableKOD.AsString:=PostCodeLAbel.Caption;
             d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);

             if d_offTable.GotoKey then
                brush.Color:=$00AAAAFF;
             end;

           except
           end;


           fillRect(Rect);
           Font.Size:=10;
           TextOut(Rect.Left+((Rect.Right - Rect.Left) div 2)-(TextWidth(Field.Value) div 2),Rect.Top+((Rect.Bottom-Rect.Top) div 2)-(TextHeight(Field.Value) div 2), Field.Value);
           end;

      end;
end;

procedure TAppointmentForm.DBGrid7CellClick(Column: TColumnEh);
var
   _year, month, day: string;
begin


   try
   if MonthTable.FieldByNAme(Column.FieldName).AsString='' then
      exit;
   except
   end;


  if (MonthsComboBox.ItemIndex+1)>9 then
      month:=IntToStr(MonthsComboBox.ItemIndex+1)
   else
      month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

   day:=MonthTable.FieldByNAme(Column.FieldName).AsString;
   if Length(day)=1 then
      day:='0'+day;

   d_offTable.Open;
   d_offTable.Setrange('','');
   d_offTable.IndexName:='KOD';
   d_offTable.SetKey;
   d_offTableKOD.AsString:=PostCodeLAbel.Caption;
   d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
   if d_offTable.GotoKey then
      d_offTable.Delete
   else
      begin

      if StrToInt(YearComboBox.Text)>StrToInt(FormatDAteTime('YYYY',date)) then
         _year:='NY'
      else
         _year:=YearComboBox.Text;


      if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
         begin
         AppTAbleTime.Close;
         AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
         AppTAbleTime.IndexFiles.Clear;
         AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
         end;

      AppTAbleTime.Open;
      AppTAbleTime.IndexName:='DATE_APP';

      if postCodelabel.CAption<>'STO' then
         AppTAbleTime.SetRange(YearComboBox.Text+month+day+postCodelabel.CAption+'',YearComboBox.Text+month+day+postCodelabel.CAption+'яя')
      else
         AppTAbleTime.SetRange(YearComboBox.Text+month+day+'',YearComboBox.Text+month+day+'яя');


      if AppTAbleTime.RecordCount>0 then
         begin
         if postCodelabel.CAption<>'STO' then
         if my_dlg('Внимание!','Уже есть запись клиентов на выбранную дату и пост.'+#13+#13+'Продолжить?',clYellow)=False then exit;

         if postCodelabel.CAption='STO' then
         if my_dlg('Внимание!','Уже есть запись клиентов на выбранную дату.'+#13+#13+'Продолжить?',clYellow)=False then exit;
         end;

      //добавляем запись
      d_offTable.Open;
      d_offTable.AppendRecord([ PostCodeLAbel.CAption,
                              StrToDate(day+'.'+month+'.'+YearComboBox.Text),
                              NULL,
                              kod_operatora,
                              date]);

      end;



   dbGrid7.Repaint;
   dbGrid9.Repaint;

   if postCodelabel.CAption='STO' then
      fill_sto_days_off(dsCalendar1,'STO')


end;

procedure TAppointmentForm.DBGrid7MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  Global_X:=X;
  Global_Y:=Y;


end;

procedure TAppointmentForm.DBGrid9DrawDataCell(Sender: TObject;
  const Rect: TRect; Field: TField; State: TGridDrawState);
var
  String_array: TString_array;
  _year, month, day, str: string;

begin
  with DBGrid9.Canvas do
       begin


      //*************************************************************************************************************
      if DBGrid9.DataSource=DAtaSource6 then
      BEGIN

      if (Trim(Field.FieldName)='POSTCODE') then
         begin
         fillRect(Rect);

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='30';
         Form1.Spr00KOD.AsString:=MonthTAblePostsPOSTCODE.AsString;
         if Form1.Spr00.GotoKey then
            begin
            String_array:=delenie_of_string_for_print(Form1.Spr00NAIM.AsString,DBGrid9.Canvas,Rect.Right-Rect.Left);

            TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

            if String_array.Str2<>'' then
               TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

            end;


         end
       else
         begin
         brush.Color:=$0097FF97;


         try
         if MonthTablePostsPOSTCODE.AsString<>'' then
              begin

              if StrToInt(YearComboBox.Text)>StrToInt(FormatDAteTime('YYYY',date)) then
                 _year:='NY'
              else
                 _year:=YearComboBox.Text;

              if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
                 begin
                 d_offTable.Close;
                 d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
                 d_offTable.IndexFiles.Clear;
                 d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');

                 if FileExists(d_offTable.TableName)=False then
                    begin
                    Create_d_offTable(_year);
                    Create_d_offTable('NY');
                    end;
                 end;

              if (MonthsComboBox.ItemIndex+1)>9 then
                 month:=IntToStr(MonthsComboBox.ItemIndex+1)
              else
                 month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

              day:=Copy(Field.FieldName,2,2);
              if Length(day)=1 then
                 day:='0'+day;


              d_offTable.Open;
              d_offTable.IndexName:='KOD';
              d_offTable.SetKey;
              d_offTableKOD.AsString:=MonthTablePostsPOSTCODE.AsString;
              d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
              if d_offTable.GotoKey then
                 brush.Color:=$00AAAAFF;


              d_offTable.Open;
              d_offTable.IndexName:='KOD';
              d_offTable.SetKey;
              d_offTableKOD.AsString:='STO';
              d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
              if d_offTable.GotoKey then
                 brush.Color:=$00AAAAFF;


              end;

           except
           end;


           fillRect(Rect);

           Font.Size:=8;
           Font.Color:=clGray;

           str:=Copy(Field.DisplayNAme,1,2);
           TextOut(Rect.Left+((Rect.Right-Rect.Left) div 2) - (TextWidth(str) div 2),Rect.Top,Str);
           str:=Copy(Field.DisplayNAme,4,2);
           TextOut(Rect.Left+((Rect.Right-Rect.Left) div 2) - (TextWidth(str) div 2),Rect.Top+TextWidth('Aa'),Str);

           end;

      END;



      //*************************************************************************************************************
      if DBGrid9.DataSource=DAtaSource8 then
      BEGIN

      if (Trim(Field.FieldName)='STAFFCODE') then
         begin

         fillRect(Rect);

         if Copy(MonthTableStaffSTAFFCODE.AsString,1,2)='MM' then
            begin
            Form1.Spr00.Open;
            Form1.Spr00.IndexName:='KOD';
            Form1.Spr00.SetKey;
            Form1.Spr00GR.AsString:='32';
            Form1.Spr00KOD.AsString:=Copy(MonthTableStaffSTAFFCODE.AsString,3,10);
            if Form1.Spr00.GotoKey then
               begin
               String_array:=delenie_of_string_for_print('Мастер - '+Form1.Spr00NAIM.AsString,DBGrid9.Canvas,Rect.Right-Rect.Left);

               TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

               if String_array.Str2<>'' then
                  TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

               end;
            end;

         if Copy(MonthTableStaffSTAFFCODE.AsString,1,2)='ST' then
            begin
            Tab_nTable.Open;
            Tab_nTable.IndexName:='TAB_N';
            if Tab_nTable.FindKey([Copy(MonthTableStaffSTAFFCODE.AsString,3,10)]) then
               begin
               String_array:=delenie_of_string_for_print('Таб.№ '+Tab_nTableTAB_N.AsString+' - '+Tab_nTableFIO.AsString,DBGrid9.Canvas,Rect.Right-Rect.Left);

               TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

               if String_array.Str2<>'' then
                  TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

               end;
            end;

         end
       else
         begin
         brush.Color:=$0097FF97;


         try
         if MonthTableStaffSTAFFCODE.AsString<>'' then
              begin

              if StrToInt(YearComboBox.Text)>StrToInt(FormatDAteTime('YYYY',date)) then
                 _year:='NY'
              else
                 _year:=YearComboBox.Text;

              if d_offTable.TableName<>path_base+'HIST\'+_year+'\D_offTable.dbf' then
                 begin
                 d_offTable.Close;
                 d_offTable.TableName:=path_base+'HIST\'+_year+'\D_offTable.dbf';
                 d_offTable.IndexFiles.Clear;
                 d_offTable.IndexFiles.ADd(path_base+'HIST\'+_year+'\D_offTable.cdx');

                 if FileExists(d_offTable.TableName)=False then
                    begin
                    Create_d_offTable(_year);
                    Create_d_offTable('NY');
                    end;
                 end;

              if (MonthsComboBox.ItemIndex+1)>9 then
                 month:=IntToStr(MonthsComboBox.ItemIndex+1)
              else
                 month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

              day:=Copy(Field.FieldName,2,2);
              if Length(day)=1 then
                 day:='0'+day;


              d_offTable.Open;
              d_offTable.IndexName:='KOD';
              d_offTable.SetKey;
              d_offTableKOD.AsString:=MonthTableStaffSTAFFCODE.AsString;
              d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
              if d_offTable.GotoKey then
                 brush.Color:=$00AAAAFF;


              d_offTable.Open;
              d_offTable.IndexName:='KOD';
              d_offTable.SetKey;
              d_offTableKOD.AsString:='STO';
              d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
              if d_offTable.GotoKey then
                 brush.Color:=$00AAAAFF;


              end;

           except
           end;


           fillRect(Rect);

           Font.Size:=8;
           Font.Color:=clGray;

           str:=Copy(Field.DisplayNAme,1,2);
           TextOut(Rect.Left+((Rect.Right-Rect.Left) div 2) - (TextWidth(str) div 2),Rect.Top,Str);
           str:=Copy(Field.DisplayNAme,4,2);
           TextOut(Rect.Left+((Rect.Right-Rect.Left) div 2) - (TextWidth(str) div 2),Rect.Top+TextWidth('Aa'),Str);

           end;

      END;


    end;
end;

procedure TAppointmentForm.DBGrid9MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  Global_X:=X;
  Global_Y:=Y;

end;

procedure TAppointmentForm.DBGrid9CellClick(Column: TColumnEh);
var
   _year, month, day: string;

begin

 //************************************************************************************************************************
 if DBGrid9.DataSource=DataSource6 then
    BEGIN

    try
    if Column.FieldName='POSTCODE' then exit;
       except
       exit; //кликнули в POSTCODE
       end;

    try
    if (MonthTablePostsPOSTCODE.ASString='') then
       exit;
    except
    end;

    if (MonthsComboBox.ItemIndex+1)>9 then
       month:=IntToStr(MonthsComboBox.ItemIndex+1)
    else
       month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

    day:=Copy(MonthTablePosts.FieldByNAme(Column.FieldName).DisplayLAbel,4,2);
    if Length(day)=1 then
       day:='0'+day;


    d_offTable.Open;
    d_offTable.Setrange('','');
    d_offTable.IndexName:='KOD';
    d_offTable.SetKey;
    d_offTableKOD.AsString:=MonthTablePostsPOSTCODE.ASString;
    d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
    if d_offTable.GotoKey then
       d_offTable.Delete
    else
       begin

       if StrToInt(YearComboBox.Text)>StrToInt(FormatDAteTime('YYYY',date)) then
          _year:='NY'
       else
          _year:=YearComboBox.Text;


       if AppTAbleTime.TableName<>path_base+'HIST\'+_year+'\apptable.dbf' then
          begin
          AppTAbleTime.Close;
          AppTAbleTime.TableName:=path_base+'HIST\'+_year+'\apptable.dbf';
          AppTAbleTime.IndexFiles.Clear;
          AppTAbleTime.IndexFiles.Add(path_base+'HIST\'+_year+'\apptable.cdx');
          end;

       AppTAbleTime.Open;
       AppTAbleTime.IndexName:='DATE_APP';

       AppTAbleTime.SetRange(YearComboBox.Text+month+day+MonthTablePostsPOSTCODE.ASString+'',YearComboBox.Text+month+day+MonthTablePostsPOSTCODE.ASString+'яя');

       if AppTAbleTime.RecordCount>0 then
          if my_dlg('Внимание!','Уже есть запись клиентов на выбранную дату и пост.'+#13+#13+'Продолжить?',clYellow)=False then exit;

      //добавляем запись
       d_offTable.Open;
       d_offTable.AppendRecord([ MonthTablePostsPOSTCODE.ASString,
                                 StrToDate(day+'.'+month+'.'+YearComboBox.Text),
                                 NULL,
                                 kod_operatora,
                                 date]);

      end;
   END;



 //************************************************************************************************************************
 if DBGrid9.DataSource=DataSource8 then
    BEGIN

    try
    if Column.FieldName='STAFFCODE' then exit;
    except
     exit; //ткнули в STAFFCODE
    end;


    try
    if (MonthTableStaffSTAFFCODE.ASString='') then
        exit;
    except
    end;

    if (MonthsComboBox.ItemIndex+1)>9 then
        month:=IntToStr(MonthsComboBox.ItemIndex+1)
    else
        month:='0'+IntToStr(MonthsComboBox.ItemIndex+1);

    day:=Copy(MonthTableStaff.FieldByNAme(Column.FieldName).DisplayLAbel,4,2);
    if Length(day)=1 then
       day:='0'+day;


    d_offTable.Open;
    d_offTable.Setrange('','');
    d_offTable.IndexName:='KOD';
    d_offTable.SetKey;
    d_offTableKOD.AsString:=MonthTableStaffSTAFFCODE.ASString;
    d_offTableD1.Value:=StrToDAte(day+'.'+month+'.'+YearComboBox.Text);
    if d_offTable.GotoKey then
       d_offTable.Delete
    else
       begin

      //добавляем запись
       d_offTable.Open;
       d_offTable.AppendRecord([ MonthTableStaffSTAFFCODE.ASString,
                              StrToDate(day+'.'+month+'.'+YearComboBox.Text),
                              NULL,
                              kod_operatora,
                              date]);

      end;
   END;




   dbGrid7.Repaint;
   dbGrid9.Repaint;
end;


procedure TAppointmentForm.CreateOrderBitBtn2Click(Sender: TObject);
var
  rezult: string;
begin


   if (k_klLAbel.CAption<>'') and (Car_N_ZAPLAbel.CAption<>'') then
       begin
       rezult:=my_vybor('Внимание!','Создать заявку на ремонт.','Перейти в макет документов.','NULL','NULL',clYellow);


       if rezult='Создать заявку на ремонт.' then
          begin

          if PageControl1.ActivePage=Tab1 then
             createNewZajProc(AppTableTime,k_klLAbel.CAption, Car_N_ZAPLAbel.CAption)
          else
          if PageControl1.ActivePage=Tab2 then
             createNewZajProc(AppTable,k_klLAbel.CAption, Car_N_ZAPLAbel.CAption);
          end
       else
          begin


          if PageControl1.ActivePage=Tab1 then
             begin
             if AppTableTimeN_ZAJ.ASString='' then
                begin
                AppTableTime.Edit;
                AppTableTimeN_ZAJ.ASString:='XXX';
                AppTableTime.Edit;
                AppTableTime.Post;
                end;
             end
          else
          if PageControl1.ActivePage=Tab2 then
             begin
             if AppTableN_ZAJ.ASString='' then
                begin
                AppTable.Edit;
                AppTableN_ZAJ.ASString:='XXX';
                AppTable.Edit;
                AppTable.Post;
                end;
             end;


          Form1.SpeedButton3.Click;

          InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
          InsertForm.PageControl1Change(InsertForm.PageControl1);

//          передать код клиента в поиск по клиенту
          if (Length(K_klLAbel.CAption)=6) and ((Copy(K_klLAbel.CAption,2,1)='F') or (Copy(K_klLAbel.CAption,2,1)='U')) then
              begin
              FindUnit.client_karman:='';
              FindUnit.Zapret_na_automat_chenge_skidki:=False;

              InsertForm.BitBtn19.Click;
              InsertForm.ComboBox5.ItemIndex:=0;
              InsertForm.MAskEdit10.Text:=K_klLAbel.CAption;
              if Copy(InsertForm.MAskEdit10.Text,2,1)='F' then
              if InsertForm.clientsQUERY.RecordCount=1 then
                 begin
                 InsertForm.RadioButton3.Checked:=True;
                 InsertForm.DBGrid3DBLClick(Sender as TObject);
                 end;

              if Copy(InsertForm.MAskEdit10.Text,2,1)='U' then
              if InsertForm.clientsQUERY_UR.RecordCount=1 then
                 begin
                 InsertForm.RadioButton4.Checked:=True;
                 InsertForm.DBGrid11DBLClick(Sender as TObject);
                 end;

              end;



          end;

       end
   else
       begin
       Form1.SpeedButton3.Click;
       InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
       InsertForm.PageControl1Change(InsertForm.PageControl1);


       end;

   DetailPanel.Visible:=False;
end;


procedure TAppointmentForm.CreateOrderBitBtn3Click(Sender: TObject);
var
  rezult: string;
begin
   if AppTAbleD_UDL.AsString<>'' then
      begin
      my_messageTime('Внимание!','Запись удалена.',clYellow,3000);
      exit;
      end;

   if (AppTAblek_kl.AsString<>'') and (AppTAbleCAR_N_ZAP.AsString<>'') then
       begin
       rezult:=my_vybor('Внимание!','Создать заявку на ремонт.','Перейти в макет документов.','NULL','NULL',clYellow);


       if rezult='Создать заявку на ремонт.' then
          begin
          createNewZajProc(AppTable,AppTAblek_kl.AsString, AppTAbleCAR_N_ZAP.AsString);
          end
       else
          begin

          if AppTableN_ZAJ.ASString='' then
             begin
             AppTable.Edit;
             AppTableN_ZAJ.ASString:='XXX';
             AppTable.Edit;
             AppTable.Post;
             end;

          Form1.SpeedButton3.Click;

          InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
          InsertForm.PageControl1Change(InsertForm.PageControl1);

//          передать код клиента в поиск по клиенту
          if (Length(AppTAblek_kl.AsString)=6) and ((Copy(AppTAblek_kl.AsString,2,1)='F') or (Copy(AppTAblek_kl.AsString,2,1)='U')) then
              begin
              FindUnit.client_karman:='';
              FindUnit.Zapret_na_automat_chenge_skidki:=False;

              InsertForm.BitBtn19.Click;
              InsertForm.ComboBox5.ItemIndex:=0;
              InsertForm.MAskEdit10.Text:=AppTAblek_kl.AsString;
              if Copy(InsertForm.MAskEdit10.Text,2,1)='F' then
              if InsertForm.clientsQUERY.RecordCount=1 then
                 begin
                 InsertForm.RadioButton3.Checked:=True;
                 InsertForm.DBGrid3DBLClick(Sender as TObject);
                 end;

              if Copy(InsertForm.MAskEdit10.Text,2,1)='U' then
              if InsertForm.clientsQUERY_UR.RecordCount=1 then
                 begin
                 InsertForm.RadioButton4.Checked:=True;
                 InsertForm.DBGrid11DBLClick(Sender as TObject);
                 end;

              end;



          end;

       end
   else
       begin
       Form1.SpeedButton3.Click;
       InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
       InsertForm.PageControl1Change(InsertForm.PageControl1);
       end;
end;

procedure TAppointmentForm.CreateOrderBitBtnClick(Sender: TObject);
var
  rezult: string;
begin

   if AppTAbleTimeD_UDL.AsString<>'' then
      begin
      my_messageTime('Внимание!','Запись удалена.',clYellow,3000);
      exit;
      end;


   if (AppTAbleTimek_kl.AsString<>'') and (AppTAbleTimeCAR_N_ZAP.AsString<>'') then
       begin

       rezult:=my_vybor('Внимание!','Создать заявку на ремонт.','Перейти в макет документов.','NULL','NULL',clYellow);


       if rezult='Создать заявку на ремонт.' then
          begin
          createNewZajProc(AppTableTime,AppTAbleTimek_kl.AsString, AppTAbleTimeCAR_N_ZAP.AsString)
          end
       else
          begin

          if AppTableTimeN_ZAJ.ASString='' then
          begin
             AppTableTime.Edit;
             AppTableTimeN_ZAJ.ASString:='XXX';
             AppTableTime.Edit;
             AppTableTime.Post;

          end;

          Form1.SpeedButton3.Click;

          InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
          InsertForm.PageControl1Change(InsertForm.PageControl1);


//          передать код клиента в поиск по клиенту
          if (Length(AppTAbleTimek_kl.AsString)=6) and ((Copy(AppTAbleTimek_kl.AsString,2,1)='F') or (Copy(AppTAbleTimek_kl.AsString,2,1)='U')) then
              begin
              FindUnit.client_karman:='';
              FindUnit.Zapret_na_automat_chenge_skidki:=False;

              InsertForm.BitBtn19.Click;
              InsertForm.ComboBox5.ItemIndex:=0;
              InsertForm.MAskEdit10.Text:=AppTAbleTimek_kl.AsString;
              if Copy(InsertForm.MAskEdit10.Text,2,1)='F' then
              if InsertForm.clientsQUERY.RecordCount=1 then
                 begin
                 InsertForm.RadioButton3.Checked:=True;
                 InsertForm.DBGrid3DBLClick(Sender as TObject);
                 end;

              if Copy(InsertForm.MAskEdit10.Text,2,1)='U' then
              if InsertForm.clientsQUERY_UR.RecordCount=1 then
                 begin
                 InsertForm.RadioButton4.Checked:=True;
                 InsertForm.DBGrid11DBLClick(Sender as TObject);
                 end;

              end;



          end;


       end
   else
       begin
       Form1.SpeedButton3.Click;
       InsertForm.PageControl1.ActivePage:=InsertForm.Tab2;
       InsertForm.PageControl1Change(InsertForm.PageControl1);
       end;

   Panel6.Visible:=False;

end;

procedure TAppointmentForm.DBGrid20DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumnEh;
  State: TGridDrawState);
var
   String_array: TString_array;
   tel: string;
   Bitmap: TBitmap;
begin
   with DBGrid20.Canvas do
      begin
      if AppTAbleMonth.IsEmpty=True then exit;


      if (not (gdSelected in State)) and (AppTAbleMonthNAVL_PR.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (AppTAbleMonthIO_UDL.AsString<>'') then
          font.Color:=AppointmentForm.Color;

      if (not (gdSelected in State)) and (AppTAbleMonthBRIGHT.AsString<>'') then
          Brush.Color:=Install.PAnel2.Color;

      DBGrid20.DefaultDrawColumnCell(Rect, DataCol, Column, State);





      if (Trim(Column.FieldName)='TEL') then
         begin
         fillRect(Rect);
         tel:=UnCryptString(AppTAbleMonthTEL.AsString,key1c,key2c);
         TextOut(Rect.Left+2,Rect.Top,Copy(tel,1,2)+'('+Copy(tel,3,3)+')'+Copy(tel,6,10));
         end;

      if (Trim(Column.FieldName)='GOS_N') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top+2,UnCryptString(AppTAbleMonthGOS_N.AsString,key1c,key2c));

         if (AppTAbleMonthN_ZAJ.AsString<>'') then
             begin
             Bitmap:=BitBtn1.Glyph;
             Bitmap.Transparent:=True;

             DBGrid20.Canvas.Draw(Rect.Right-Bitmap.Width, Rect.Top+2,Bitmap);
             end;


         end;

      if (Trim(Column.FieldName)='FIO') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(AppTAbleMonthFIO.AsString,DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;


      if (Trim(Column.FieldName)='OBR_KL') then
         begin
         fillRect(Rect);

         String_array:=delenie_of_string_for_print(UnCryptString(AppTAbleMonthOBR_KL.AsString,key1c,key2c),DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);

         end;

      if (Trim(Column.FieldName)='COMM') then
         begin
         fillRect(Rect);
         String_array:=delenie_of_string_for_print(UnCryptString(AppTAbleMonthCOMM.AsString,key1c,key2c),DBGrid2.Canvas,Rect.Right-Rect.Left);
         TextOut(Rect.Left+2,Rect.Top,String_array.Str1);
         if Trim(String_array.Str1)<>'' then
             TextOut(Rect.Left+2,Rect.Top+TextHeight('A'),String_array.Str2);
         end;





      if (Trim(Column.FieldName)='POSTCODE') then
         begin
         fillRect(Rect);

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='30';
         Form1.Spr00KOD.AsString:=AppTAbleMonthPOSTCODE.AsString;
         if Form1.Spr00.GotoKey then
            begin
            String_array:=delenie_of_string_for_print(Form1.Spr00NAIM.AsString,DBGrid2.Canvas,Rect.Right-Rect.Left);

            TextOut(Rect.Left+2,Rect.Top,String_array.Str1);

            if String_array.Str2<>'' then
               TextOut(Rect.Left+2,Rect.Top+TextWidth('Aa'),String_array.Str2);

            end;


         end;

      if (Trim(Column.FieldName)='TIME1') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTAbleMonthTIME1.AsString,1,2)+':'+Copy(AppTAbleMonthTIME1.AsString,3,2));
         end;

      if (Trim(Column.FieldName)='TIME2') and (trim(AppTAbleMonthTIME2.AsString)<>'') then
         begin
         fillRect(Rect);
         TextOut(Rect.Left+2,Rect.Top,Copy(AppTAbleMonthTIME2.AsString,1,2)+':'+Copy(AppTAbleMonthTIME2.AsString,3,2));
         end;


      if (Trim(Column.FieldName)='IO_ADD') then
         begin
         fillRect(Rect);

         font.Size:=7;

         Form1.Spr00.Open;
         Form1.Spr00.IndexName:='KOD';
         Form1.Spr00.SetKey;
         Form1.Spr00GR.AsString:='02';
         Form1.Spr00KOD.AsString:=AppTAbleMonthIO_ADD.AsString;
         if Form1.Spr00.GotoKey then
            TextOut(Rect.Left+2,Rect.Top,Form1.Spr00NAIM.AsString);

         end;



      end;



end;

procedure TAppointmentForm.ClientNameMaskEditDblClick(Sender: TObject);
begin
  ClientSpravBitBtn.Click;
end;

procedure TAppointmentForm.BitBtn56Click(Sender: TObject);
begin
  MonthsComboBox.ItemIndex:=MonthsComboBox.ItemIndex-1;
  if MonthsComboBox.ItemIndex=-1 then
     begin
     MonthsComboBox.ItemIndex:=11;
     YearComboBox.ItemIndex:=YearComboBox.ItemIndex-1;
     if YearComboBox.ItemIndex=-1 then
        begin
        YearComboBox.ItemIndex:=0;
        MonthsComboBox.ItemIndex:=0;
        end;
     end;

  MonthsComboBoxChange(MonthsComboBox);

     
end;

procedure TAppointmentForm.BitBtn57Click(Sender: TObject);
begin
  try
  if MonthsComboBox.ItemIndex=11 then
     begin
     MonthsComboBox.ItemIndex:=0;
     try
     YearComboBox.ItemIndex:=YearComboBox.ItemIndex+1;
     except
        YearComboBox.ItemIndex:=YearComboBox.Items.Count-1;
        MonthsComboBox.ItemIndex:=MonthsComboBox.Items.Count-1;
     end;
     end
  else
     MonthsComboBox.ItemIndex:=MonthsComboBox.ItemIndex+1;
  except
  end;

  MonthsComboBoxChange(MonthsComboBox);
end;

procedure TAppointmentForm.SunMaskEditChange(Sender: TObject);
begin
  if SunMaskEdit.Text='с   :   до   :  ' then
     SunMaskEdit.Color:=$00AAAAFF
  else
     SunMaskEdit.Color:=Form1.Panel2.Color;

end;

procedure TAppointmentForm.FriMaskEditChange(Sender: TObject);
begin
  if FriMaskEdit.Text='с   :   до   :  ' then
     FriMaskEdit.Color:=$00AAAAFF
  else
     FriMaskEdit.Color:=Form1.Panel2.Color;


end;

procedure TAppointmentForm.ThuMaskEditChange(Sender: TObject);
begin
  if ThuMaskEdit.Text='с   :   до   :  ' then
     ThuMaskEdit.Color:=$00AAAAFF
  else
     ThuMaskEdit.Color:=Form1.Panel2.Color;


end;

procedure TAppointmentForm.WedMaskEditChange(Sender: TObject);
begin
  if WedMaskEdit.Text='с   :   до   :  ' then
     WedMaskEdit.Color:=$00AAAAFF
  else
     WedMaskEdit.Color:=Form1.Panel2.Color;


end;


procedure export_proc(postcode: string);
var
  i: integer;
  ExcelApp, WorkSheet: Variant;

begin
  if not IsOLEObjectInstalled('Excel.Application') then //нет Excel
     begin
     my_messageTime('Внимание!','Microsoft Excel не найден на Вашем компьютере.',clYellow,5000);
     exit;
     end;

with AppointmentForm do
begin

ExcelApp:=CreateOleObject('Excel.Application');
ExcelApp.Workbooks.Add;
ExcelApp.Application.EnableEvents := false;
WorkSheet:=ExcelApp.Workbooks[1].Worksheets.Add;


ExcelApp.Visible:=True;
SetForegroundWindow(ExcelApp.Hwnd);


WorkSheet.Activate;




ExcelApp.Workbooks[1].WorkSheets[1].Cells.Font.Size:=9;

  ExcelApp.Workbooks[1].WorkSheets[1].Cells[1,1].Font.Color:=clRed;
  ExcelApp.Workbooks[1].WorkSheets[1].Cells[2,1].Font.Color:=clRed;


 if postcode<>'' then
    begin
    Form1.Spr00.Open;
    Form1.Spr00.IndexName:='KOD';
    Form1.Spr00.SetKey;
    Form1.Spr00GR.AsString:='30';
    Form1.Spr00KOD.AsString:=postcode;
    if Form1.Spr00.GoToKey then
        ExcelApp.Workbooks[1].WorkSheets[1].Cells[1,1]:='Запись на ремонт: '+Form1.Spr00NAIM.AsString;
    end
  else
        ExcelApp.Workbooks[1].WorkSheets[1].Cells[1,1]:='Запись на ремонт';

  ExcelApp.Workbooks[1].WorkSheets[1].Cells[2,1]:='Дата: '+FormatDateTime('dd MMMM YYYY',dsCalendar1.Date);


  i:=4;
  ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,1]:='Время';

  ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,2]:='Пост';
  ExcelApp.Workbooks[1].WorkSheets[1].Range['B'+IntToStr(i),'B'+IntToStr(i)].ColumnWidth :=15;

  ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,3]:='Автомобиль';
  ExcelApp.Workbooks[1].WorkSheets[1].Range['C'+IntToStr(i),'C'+IntToStr(i)].ColumnWidth :=15;

  ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,4]:='Причина обращения клиента / комментарий';
  ExcelApp.Workbooks[1].WorkSheets[1].Range['D'+IntToStr(i),'D'+IntToStr(i)].ColumnWidth :=45;



          //рисуем рамку ***************
  ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'D'+IntToStr(i)].Borders.Color:=clBlack;
  ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'D'+IntToStr(i)].Borders.Weight:=3; //xlThin
  // ****************************


   inc(i);


   AppTable.First;
   while AppTable.Eof=False do
         begin
         if (postcode=AppTablePOSTCODE.AsString) or (postcode='') then
         if AppTableIO_UDL.AsString='' then
            begin
            ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,1].NumberFormat:='@';
            ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,1]:=Copy(AppTableTIME1.AsString,1,2)+':'+Copy(AppTableTIME1.AsString,3,2);

            Form1.Spr00.Open;
            Form1.Spr00.IndexName:='KOD';
            Form1.Spr00.SetKey;
            Form1.Spr00GR.AsString:='30';
            Form1.Spr00KOD.AsString:=AppTablePOSTCODE.AsString;
            if Form1.Spr00.GoToKey then
               ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,2]:=Form1.Spr00NAIM.AsString;

            ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,3]:=AppTableMARKA.AsString+' - '+AppTableMODEL.AsString;

            if trim(AppTableCOMM.AsString)='' then
                ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,4]:=UnCryptString(AppTableOBR_KL.AsString,key1c,key2c)
            else
                ExcelApp.Workbooks[1].WorkSheets[1].Cells[i,4]:=UnCryptString(AppTableOBR_KL.AsString,key1c,key2c)+#13+#10+#13+#10+'Комментарий:'+#13+#10+UnCryptString(AppTableCOMM.AsString,key1c,key2c);


            ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'E'+IntToStr(i)].WrapText:=True;

           //выравнивание текста
            ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'D'+IntToStr(i)].VerticalAlignment:=1;//top ;//2 - xlCenter; //3 - bottom


          //рисуем рамку ***************
            ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'D'+IntToStr(i)].Borders.Color:=clBlack;
            ExcelApp.Workbooks[1].WorkSheets[1].Range['A'+IntToStr(i),'D'+IntToStr(i)].Borders.Weight:=2; //xlThin
           // ****************************


            inc(i)
            end;

         AppTable.Next;
         end;


     ExcelApp.Visible:=True; ///!!! Открывает Excell
     SetForegroundWindow(ExcelApp.Hwnd);
     ExcelApp.Workbooks[1].Sheets.Item[1].Activate;
     SetWindowPos(ExcelApp.Hwnd,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE+SWP_NOSIZE);
     SetWindowPos(ExcelApp.Hwnd,HWND_NOTOPMOST,0,0,0,0,SWP_NOMOVE+SWP_NOSIZE);
end;


end;



procedure TAppointmentForm.N16Click(Sender: TObject);
var
   rezult: string;
begin
   rezult:=my_vybor('Внимание!','Экспорт по текущему посту.','Экспорт по всем постам.','NULL','NULL',clYellow);

   if rezult='Экспорт по текущему посту.' then
      export_proc(AppTablePOSTCODE.AsString)
   else
   if rezult='Экспорт по всем постам.' then
      export_proc('');




end;

procedure TAppointmentForm.BitBtn17Click(Sender: TObject);
var
   rezult: string;
begin
   rezult:=my_vybor('Внимание!','Экспорт по текущему посту.','Экспорт по всем постам.','NULL','NULL',clYellow);

   if rezult='Экспорт по текущему посту.' then
      export_proc(AppTablePOSTCODE.AsString)
   else
   if rezult='Экспорт по всем постам.' then
      export_proc('');

end;

procedure TAppointmentForm.N17Click(Sender: TObject);
begin
try
      AppTAble.Edit;

      if AppTAbleN_ZAJ.ASString='' then
         AppTAbleN_ZAJ.ASString:='XXX'
      else
      if AppTAbleN_ZAJ.ASString='XXX' then
         AppTAbleN_ZAJ.ASString:='';

      AppTAble.Edit;
      AppTAble.Post;

except
end;


end;

procedure TAppointmentForm.N19Click(Sender: TObject);
begin
try
      AppTAbleTime.Edit;

      if AppTAbleTimeN_ZAJ.ASString='' then
         AppTAbleTimeN_ZAJ.ASString:='XXX'
      else
      if AppTAbleTimeN_ZAJ.ASString='XXX' then
         AppTAbleTimeN_ZAJ.ASString:='';

      AppTAbleTime.Edit;
      AppTAbleTime.Post;



except
end;

end;

procedure TAppointmentForm.AllPostsComboBoxChange(Sender: TObject);
begin
     AppTAbleMonth.Open;

     if AllPostsComboBox.ItemIndex=0 then
        begin
        AppTAbleMonth.IndexName:='DATE_APP1';
        AppTAbleMonth.SetRange(FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'01'+''+'',FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'31'+'"'+'яя');
        end
     else
        begin
        Form1.Spr00.IndexName:='NAIM';
        Form1.Spr00.SetRAnge('','');
        Form1.Spr00.SetKey;
        Form1.Spr00GR.AsString:='30';
        Form1.Spr00NAIM.AsString:=AllPostsComboBox.Text;
        if Form1.Spr00.GotoKey then
           begin
           AppTAbleMonth.IndexName:='DATE_APP'; //DTOS(DATE_APP)+POSTCODE+TIME1
           AppTAbleMonth.SetRange(FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'01'+Form1.Spr00KOD.AsString+'',FormatDateTime('YYYYMM',dsCAlendar1.DAte)+'31'+Form1.Spr00KOD.AsString+'яя');
           end;



        end;


     AppTAbleMonth.First;

end;

procedure TAppointmentForm.WorkPostsComboBoxChange(Sender: TObject);
begin
    try
    DBGrid2.SetFocus;
    except
    end;

     if WorkPostsComboBox.ItemIndex=0 then
        begin
        AppTAble.Open;
        AppTAble.IndexName:='DATE_APP1';
        AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+''+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+'"'+'яя');
        AppTAble.First;

        While AppTAble.Eof=FAlse do
              begin
              try
              AppTAble.Edit;
              AppTAbleBRIGHT.AsString:='';
              AppTAble.Edit;
              AppTAble.Post;
              except
              end;

              AppTAble.Next;
              end;

        end
     else
        begin
        Form1.Spr00.IndexName:='NAIM';
        Form1.Spr00.SetRAnge('','');
        Form1.Spr00.SetKey;
        Form1.Spr00GR.AsString:='30';
        Form1.Spr00NAIM.AsString:=WorkPostsComboBox.Text;
        if Form1.Spr00.GotoKey then
           begin
           AppTAble.Open;
           AppTAble.IndexName:='DATE_APP';
           AppTAble.SetRange(FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+Form1.Spr00KOD.AsString+'',FormatDateTime('YYYYMMDD',dsCAlendar1.DAte)+Form1.Spr00KOD.AsString+'яя');
           AppTAble.First;

           While AppTAble.Eof=FAlse do
                 begin
                 try
                 AppTAble.Edit;
                 AppTAbleBRIGHT.AsString:='';
                 AppTAble.Edit;
                 AppTAble.Post;
                 except
                 end;

                 AppTAble.Next;
                 end;

           end;


        end;

end;

procedure TAppointmentForm.OrderNumberMaskEditClick(Sender: TObject);
begin
   OrderNumberMaskEdit.SelectAll;
end;

procedure TAppointmentForm.OrderNumberMaskEditChange(Sender: TObject);
begin
   if (Length(Trim(OrderNumberMaskEdit.Text))=9) and (Copy(OrderNumberMaskEdit.Text,4,1)='J') then OpenOrderBitBtn.Enabled:=True
   else
       OpenOrderBitBtn.Enabled:=FAlse;

end;

procedure TAppointmentForm.N22Click(Sender: TObject);
begin
            ZajavkiForm.Show;
            if ZajavkiForm.Zaj_oQuery.Locate('N_DOK',AppTAbleN_ZAJ.ASString,[])=True then
               begin
               ZajavkiForm.DBGrid2DBLClick(ZajavkiForm.DBGrid2);
               AppointmentForm.Close;
               end
            else
               my_messageTime('Внимание!','Документ не найден.',clYellow,3000);

end;

procedure TAppointmentForm.OpenOrderBitBtnClick(Sender: TObject);
begin
            ZajavkiForm.Show;
            if ZajavkiForm.Zaj_oQuery.Locate('N_DOK',OrderNumberMaskEdit.Text,[])=True then
               begin
               ZajavkiForm.DBGrid2DBLClick(ZajavkiForm.DBGrid2);
               CreateOrderBitBtn2.Click;
               AppointmentForm.Close;
               end
            else
               my_messageTime('Внимание!','Документ не найден.',clYellow,3000);

end;

procedure TAppointmentForm.N24Click(Sender: TObject);
begin
            ZajavkiForm.Show;
            if ZajavkiForm.Zaj_oQuery.Locate('N_DOK',AppTAbleTimeN_ZAJ.ASString,[])=True then
               begin
               ZajavkiForm.DBGrid2DBLClick(ZajavkiForm.DBGrid2);
               AppointmentForm.Close;
               end
            else
               my_messageTime('Внимание!','Документ не найден.',clYellow,3000);

end;

procedure TAppointmentForm.PopupMenu3Popup(Sender: TObject);
begin
  N24.Visible:=False;

  if (AppTAbleTimeN_ZAJ.ASString<>'XXX') and (AppTAbleTimeN_ZAJ.ASString<>'') then
      begin
      N24.Caption:='Переход в созданную заявку на ремонт -> '+AppTAbleTimeN_ZAJ.ASString;
      N24.Visible:=True;
      end;
end;

end.
