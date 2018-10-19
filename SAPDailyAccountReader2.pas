unit SAPDailyAccountReader2;
{
闻泰问题：
  1、来料出库
  2、出货扣帐
  3、调整
}
interface

uses
  Classes, SysUtils, ComObj, CommUtils, ADODB, StockMZ2FacReader;

type
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TSAPDailyAccountReader2_icmo_mz2fac = class;

  TDailyAccount_winB = packed record
    sfacname: string;
    sbillno: string;
    sdoc: string;
    snumber: string;
    smodel: string;
    snumber_wt: string;
    sname: string;
    sstock: string;  
    sstock_ml: string;
    dQty: Double;
    sunit: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;
    sicmo: string;
    sstock_desc_wt: string;
    smvr_desc: string;

    dt: TDateTime;
    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;     
    dtCheck: TDateTime;    
    suse: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    sbiller: string;
    sclose: string;
    sdept: string;
    schecktype: string;
    sedi: string;
    ssourcebillno: string;

    scheckflag: string;
    sstock_yd: string;
    bcalc: Boolean;
  end;
  PDailyAccount_winB = ^TDailyAccount_winB;

  TSAPDailyAccountReader2_winB = class  
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_winB;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_winB read GetItems;
  end;
  
  TSAPDailyAccountReader2_winB_ML = class(TSAPDailyAccountReader2_winB)
  protected
    procedure Open; override; 
  end;                               
  
  TSAPDailyAccountReader2_winB_wt = class(TSAPDailyAccountReader2_winB)
  protected
    procedure Open; override; 
  end;                                
  
  TSAPDailyAccountReader2_winB_yd = class(TSAPDailyAccountReader2_winB)
  protected
    procedure Open; override; 
  end;

  TDailyAccount_RTV = packed record
    sbillno: string;
    snumber: string;
    sname: string;
    sstock: string;  
    sstock_ml: string;
    dQty: Double;
    sunit: string;

    dt: TDateTime;     
    dtCheck: TDateTime;    
    suse: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    sbiller: string;
    sclose: string;
    sdept: string;
    scheckflag: string;
    sedi: string;
    ssourcebillno: string;    
  end;
  PDailyAccount_RTV = ^TDailyAccount_RTV;

  TSAPDailyAccountReader2_RTV = class  
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_RTV;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_RTV read GetItems;
  end;
  
  TSAPDailyAccountReader2_RTV_ML = class(TSAPDailyAccountReader2_RTV)
  protected
    procedure Open; override; 
  end;
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_cpin = packed record
    sicmo: string;
    sstock_desc_wt: string;
    dt: TDateTime;
    dtcheck: TDateTime;
    snumber: string;
    sname: string;
    dqty: Double;
    sbatchno: string; 
    sstock_ml: string;
    sstock: string;
    sbillno: string;
    snote: string;
    sdept: string;
    sbiller: string;
    schecker: string;
    scheckflag: string;
    sbackflush: string;
    sedi: string;

    sfacname: string;
    sdoc: string;
    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;
    snumber_wt: string;
    smodel: string;
    sunit: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    smvr_desc: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;

    sstock_yd: string;
  end;
  PDailyAccount_cpin = ^TDailyAccount_cpin;

  TSAPDailyAccountReader2_cpin = class 
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_cpin;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_cpin read GetItems;
  end;
  
  TSAPDailyAccountReader2_cpin_ML = class(TSAPDailyAccountReader2_cpin)
  protected
    procedure Open; override; 
  end;  
  
  TSAPDailyAccountReader2_cpin_wt = class(TSAPDailyAccountReader2_cpin)
  protected
    procedure Open; override; 
  end;          
  
  TSAPDailyAccountReader2_cpin_yd = class(TSAPDailyAccountReader2_cpin)
  protected
    procedure Open; override; 
  end;   
                
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_qin = packed record
    sbillno: string;
    snumber: string;
    sname: string;
    dqty: Double;


    dt: TDateTime;
    dtcheck: TDateTime;
    suse: string;
    ssupplier: string;
    snote: string;
    sstock_ml: string;
    sstock: string;
    ssummary: string;
    sbiller: string;
    scloseflag: string;
    sdept: string;
    schecktype: string;
    sedit: string;
    ssourcebillno: string;

    sfacname: string;
    sdoc: string;
    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;
    snumber_wt: string;
    smodel: string;
    sunit: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    smvr_desc: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;
    sicmo: string;
    sstock_desc_wt: string;

    sstock_yd: string;


    sstock_in_ml: string;
    sstock_in: string;   

    sedi: string;



    ///////

    sclose: string;

    scheckflag: string;
  end;
  PDailyAccount_qin = ^TDailyAccount_qin;

  TSAPDailyAccountReader2_qin = class 
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_qin;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_qin read GetItems;
  end;
  
  TSAPDailyAccountReader2_qin_ML = class(TSAPDailyAccountReader2_qin)
  protected
    procedure Open; override; 
  end;  
  
  TSAPDailyAccountReader2_qin_wt = class(TSAPDailyAccountReader2_qin)
  protected
    procedure Open; override; 
  end;
  
  TSAPDailyAccountReader2_qin_yd = class(TSAPDailyAccountReader2_qin)
  protected
    procedure Open; override; 
  end;
                
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_qout = packed record
    snumber: string;
    sname: string;
    dqty: Double;


    dt: TDateTime;
    dtcheck: TDateTime;
    sstock_ml: string;
    sstock: string;
    sdetp: string;
    sbillno: string;
    suse1: string;
    snote: string;
    sbiller: string;
    sunit: string;
    scheckflag: string;
    souttype: string;
    suse2: string;
    sedi: string;

    sfacname: string;
    sdoc: string;
    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;
    snumber_wt: string;
    smodel: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    smvr_desc: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;
    sicmo: string;
    sstock_desc_wt: string;

    sstock_yd: string;
  end;
  PDailyAccount_qout = ^TDailyAccount_qout;

  TSAPDailyAccountReader2_qout = class 
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_qout;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_qout read GetItems;
  end;
  
  TSAPDailyAccountReader2_qout_ML = class(TSAPDailyAccountReader2_qout)
  protected
    procedure Open; override; 
  end;  
  
  TSAPDailyAccountReader2_qout_wt = class(TSAPDailyAccountReader2_qout)
  protected
    procedure Open; override; 
  end;                               
  
  TSAPDailyAccountReader2_qout_yd = class(TSAPDailyAccountReader2_qout)
  protected
    procedure Open; override; 
  end;

  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_DB = packed record
    sbillno: string;
    snumber: string;
    sname: string;      
    sstock_in_ml: string;
    sstock_out_ml: string;
    sstock_in: string;  
    sstock_out: string;
    dQty: Double;

    dt: TDateTime;     
    dtCheck: TDateTime;    
    suse: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    sbiller: string;
    scheckflag: string;
    sedi: string;

    sfacname: string;
    sdoc: string;
    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;
    snumber_wt: string;
    smodel: string;
    sunit: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;  
    sstock: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    smvr_desc: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;
    sicmo: string;
    sstock_desc_wt: string;
    sstock_desc: string;

    bCalc: Boolean;

    sstock_out_yd: string;
    sstock_in_yd: string;
  end;
  PDailyAccount_DB = ^TDailyAccount_DB;

  TSAPDailyAccountReader2_DB = class // 外购入库 蓝字
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_DB;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_DB read GetItems;
  end;
  
  TSAPDailyAccountReader2_DB_ML = class(TSAPDailyAccountReader2_DB)
  protected
    procedure Open; override; 
  end;  
  
  TSAPDailyAccountReader2_DB_wt = class(TSAPDailyAccountReader2_DB)
  protected
    procedure Open; override;
  public
    function GetItem2(aDailyAccount_DBPtr: PDailyAccount_DB): PDailyAccount_DB;
  end;
  
  TSAPDailyAccountReader2_DB_yd = class(TSAPDailyAccountReader2_DB)
  protected
    procedure Open; override; 
  end;
       
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_DB_in = packed record // 调入
    sbillno: string;
    snumber: string;
    sname: string;      
    sstock_in_ml: string; 
    sstock_in: string;   
    dQty: Double;

    dt: TDateTime;     
    dtCheck: TDateTime;    
    suse: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    sbiller: string;
    scloseflag: string;
    sdept: string;
    schecktype: string;
    sedi: string;
    ssourcebillno: string;



    ///////
    sfacname: string;
    sdoc: string;
    smodel: string;
    snumber_wt: string;
    sstock: string;
    sstock_ml: string;
    sunit: string;
    stext: string;
    swc: string;
    sitemtext: string;
    sitemno: string;
    sstock_wt: string;
    sfacno: string;
    sitemgroupname: string;
    sitemgroup: string;
    sordertype: string;
    dicmoqty: Double;
    sdoc_item: string;
    smvt_desc: string;
    sstatus: string;
    dtbill: TDateTime;
    dbillqty: Double;
    sfac: string;
    sicmo: string;
    sstock_desc_wt: string;
    smvr_desc: string;

    smpn: string;
    smpn_name: string;
    smvt: string;
    smvr: string;     
    sclose: string;

    scheckflag: string;
    sstock_yd: string;

    sstockno_out_yd: string;
    sstockno_in_yd: string;      
    sstockno_out: string;
    sstockno_in: string;
  end;
  PDailyAccount_DB_in = ^TDailyAccount_DB_in;

  TSAPDailyAccountReader2_DB_in = class // 外购入库 蓝字
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_DB_in;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_DB_in read GetItems;
  end;
  
  TSAPDailyAccountReader2_DB_in_ML = class(TSAPDailyAccountReader2_DB_in)
  protected
    procedure Open; override; 
  end;
  
  TSAPDailyAccountReader2_DB_in_wt = class(TSAPDailyAccountReader2_DB_in)
  protected
    procedure Open; override; 
  end;
  
  TSAPDailyAccountReader2_DB_in_yd = class(TSAPDailyAccountReader2_DB_in)
  protected
    procedure Open; override; 
  end;
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////





  
  TDailyAccount_coois = packed record // 投料单
    sbillno_fac: string; //代工厂工单
    sbillno: string; //	订单
    scategory: string; //	类型
    dtfac: TDateTime; //	代工厂单据日期
    sbiller: string; //	制单人
    snumber: string; //	物料
    dtFinish: TDateTime; //	计划完工
    sbillno_plan: string; //	计划订单
    dqtyorder: Double; //	订单数量
    sBUn: string; //	BUn
    sstockname: string; //库位
    snumber_item: string; //	物料
    dtneed: TDateTime; //	需求日期
    dqtyneed: Double; //	需求量
    sunit: string; //计
    sfac: string; //工厂
    sFix: string; //	Fix
    dtChangeDate: TDateTime; //	变更日期
    dtChangeTime: TDateTime; //	变更时间
    dQtyIn: Double; //	收货数量

    bCalc: Boolean;
    sMatchType: string;
  end;
  PDailyAccount_coois = ^TDailyAccount_coois;

  TSAPDailyAccountReader2_coois = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_coois;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader;
      aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_coois read GetItems;
    function Fac2MZBillno(const sicmobillno: string): string;
  end;


  TDailyAccount_FacICMO2MZICMO = packed record // 投料单
    sicmo_fac: string;
    sicmo: string;
    stype: string;
    sfacno: string;
    snumber: string;
    sname: string;
    slang: string;
    swwpo: string;
    ssourceorder: string;
    dtFac: TDateTime;
    dtend: TDateTime;
    dtbegin: TDateTime;
 

    bCalc: Boolean;
    sMatchType: string;
  end;
  PDailyAccount_FacICMO2MZICMO = ^TDailyAccount_FacICMO2MZICMO;

  TSAPDailyAccountReader2_FacICMO2MZICMO = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_FacICMO2MZICMO;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader;
      aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_FacICMO2MZICMO read GetItems;
  end;


  TDailyAccount_icmo_mz2fac = packed record // 投料单
    sicmolbillno_fac: string; //代工厂工单
    sicmobillno: string; //	订单
    stype: string; //	类型
    sfacno: string; //	代工厂代号
    dtdate_fac: TDateTime; //	代工厂单据日期
    sbiller: string; //	制单人
    ssourcebillno: string; //	来源订单号
    swwcontract1: string; //	委外合同1
    dqty_contract_alloc1: Double; // 合同分配数量1
//    sEUn: string; //EUn
    swwcontract2: string; //	委外合同2
    dqty_contract_alloc2: Double; // 合同分配数量2
    //EUn
    //EUn
    snote: string; //备注
    sall_transfer_flag: string; //	完全转换标志
    dtChangeDate: TDateTime; //	变更日期
    dtChangeTime: TDateTime; //	变更时间
  end;
  PDailyAccount_icmo_mz2fac = ^TDailyAccount_icmo_mz2fac;

  TSAPDailyAccountReader2_icmo_mz2fac = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_icmo_mz2fac;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function billno_mz2fac(const sbillno_mz: string): string;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_icmo_mz2fac read GetItems;
  end;

  TDailyAccount_PPBom = packed record // 投料单
   dtdate: TDateTime; //制单日期
   dtCheck: TDateTime; //	审核日期
   sicmobillno: string; //	生产/委外订单号
   snumber: string; //	产品代码
   sname: string; //	产品名称
   dqty: Double; //	生产数量
   snote: string; //	备注
   sppbombillno: string; //	生产投料单号
   snumber_item: string; //	子项物料长代码
   sname_item: string; // 子项物料名称
   dqtyplan: Double; //	计划投料数量
   dqtyshould: Double; //	应发数量
   sstockname: string; //	仓库
   sstockname_ml: string; //	仓库
   dusage: Double; //	单位用量
   scheckflag: string; //	审核标志
   sworkshopname: string; //	生产车间
   sedi: string; //	EDI提交

   sfacname: string;
   sfac: string;
   sicmotye: string;
   dtRelease: TDateTime;
   dtClose: TDateTime;
   dtBegin: TDateTime;
   dtEnd: TDateTime;
   splanbillno: string;
   splanbillno_mz: string;
   snumber_wt: string;
   svItemFlag: string;
   sItemCode: string;
   dICMOQty: Double;
   snote1: string;
   iChangeCount: string;
   irowitem: string;
   snumber_item_wt: string;
   dqtyout: Double;
   sstockname_wt: string;
   dqty0: Double;
   sgroup: string;
   sprioriry: string;
   dper: Double;
   sunit: string;
   snote2: string;
   schangelog: string;

   sstockname_yd: string;
  end;
  PDailyAccount_PPBom = ^TDailyAccount_PPBom;

  TSAPDailyAccountReader2_PPBOM = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_PPBom;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_PPBom read GetItems;
  end;
       
  TSAPDailyAccountReader2_PPBOM_ml = class(TSAPDailyAccountReader2_PPBOM)
  protected
    procedure Open; override; 
  end;  
       
  TSAPDailyAccountReader2_PPBOM_wt = class(TSAPDailyAccountReader2_PPBOM)
  protected
    procedure Open; override;
  end;
       
  TSAPDailyAccountReader2_PPBOM_yd = class(TSAPDailyAccountReader2_PPBOM)
  protected
    procedure Open; override;
  end;
         
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_PPBomChange_mz = packed record
    schangebillno: string; //变更单号
    sarea: string; //MRP 范围
    sbillno: string; //单据编号
    sorderbillno: string; //订单
    snumber: string; //物料
    sunit: string; //计
    splanroder: string; //计划订单
    sqtychangeflag: string; //数量变更标志
    sreason: string; //变更原因
    schangetime: string; //变更时间
    snumber_item: string; //组件
    sZTIPP007B_ITEM: string; //ZTIPP007B-ITEM
    sZTIPP007B_LGORT: string; //ZTIPP007B-LGORT
    sZTIPP007B_ALPGR: string; //ZTIPP007B-ALPGR
    sZTIPP007B_ALPRF: string; //ZTIPP007B-ALPRF
    sZTIPP007B_EWAHR: string; //ZTIPP007B-EWAHR
    sZTIPP007B_ITEM_FLAG: string; //ZTIPP007B-ITEM_FLAG
    sZTIPP007B_REMARK: string; //ZTIPP007B-REMARK
    sZTIPP007B_UPDKZ: string; //ZTIPP007B-UPDKZ
    sicmo_fac: string; //代工厂工单
    dqty: Double;//数量
    //计
    dqtyBefore: Double; //修改前数量
    //计
    dtChange: TDateTime;//变更日期
    sZTIPP007B_MENGE: string; //ZTIPP007B-MENGE
    dtNeed: TDateTime; //需求日期
    sZTIPP007B_MENGE_B: string; //ZTIPP007B-MENGE_B
    sZTIPP007B_MENGE_T: string; //ZTIPP007B-MENGE_T
    //变更日期
    bCalc:Boolean;
    sMatchType: string;
  end;
  PDailyAccount_PPBomChange_mz = ^TDailyAccount_PPBomChange_mz;
       
  TDailyAccount_PPBomChange_yd = packed record
    sChangeFlag: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('变更标志').AsString;
    snumber: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('产品代码').AsString;
    sname: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('产品名称').AsString;
    sppbombillno: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('生产投料单号').AsString;
    snumber_item: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('物料代码').AsString;
    sname_item: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('物料名称').AsString;
    susage: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('标准用量').AsString;
    sstock_fac: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('仓库').AsString;
    sChangeReason: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('变更原因').AsString;
    sdt: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('制单日期').AsString;
    sdtCheck: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('审核日期').AsString;
    sChangeVer: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('变更版次').AsString;
    dQty: Double; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('计划投料数量').AsString;
    sstock: string;
  end;
  PDailyAccount_PPBomChange_yd = ^TDailyAccount_PPBomChange_yd;
  
  TSAPDailyAccountReader2_PPBOMChange = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
  end;
 
  TSAPDailyAccountReader2_PPBOMChange_mz = class(TSAPDailyAccountReader2_PPBOMChange)
  protected
    function GetItems(i: Integer): PDailyAccount_PPBomChange_mz;
    procedure Open; override;
  public
    property Items[i: Integer]: PDailyAccount_PPBomChange_mz read GetItems;
  end;    
 
  TSAPDailyAccountReader2_PPBOMChange_yd = class(TSAPDailyAccountReader2_PPBOMChange)
  protected                       
    function GetItems(i: Integer): PDailyAccount_PPBomChange_yd;
    procedure Open; override;   
  public
    property Items[i: Integer]: PDailyAccount_PPBomChange_yd read GetItems;
  end;

  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_DB_out = packed record // 调出
    sbillno: string;
    snumber_wt: string;
    snumber: string;
    sname: string;      
    sstock_out_ml: string;
    dQty: Double;

    dt: TDateTime;     
    dtCheck: TDateTime;
    suse1: string;
    suse2: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    souttype: string;
    sbiller: string;
    scloseflag: string;
    sdept: string;
    scheckflag: string;
    sunit: string;
    sedi: string;

    sfacname: string;
    sdoc: string;
    smvt: string;
    smvtdesc: string;
    sfac: string;
    sstock_out_wt: string;
    sstock_out: string;
    sstock_out_desc: string;
    sstock_out_desc_wt: string;
  end;
  PDailyAccount_DB_out = ^TDailyAccount_DB_out;

  TSAPDailyAccountReader2_DB_out = class 
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_DB_out;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_DB_out read GetItems;
  end;
  
  TSAPDailyAccountReader2_DB_out_ML = class(TSAPDailyAccountReader2_DB_out)
  protected
    procedure Open; override; 
  end;

  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_OutAInBC = packed record // 调出
    sbillno: string;
    snumber: string;
    sname: string;      
    sstock_out_ml: string;
    sstock_out: string;
    dQty: Double;

    dt: TDateTime;     
    dtCheck: TDateTime;
    suse1: string;
    suse2: string;
    ssupplier: string;
    snote: string;
    ssummary: string;
    souttype: string;
    sbiller: string;
    scloseflag: string;
    sdept: string;
    scheckflag: string;
    sunit: string;
    sedi: string;    
  end;
  PDailyAccount_OutAInBC = ^TDailyAccount_OutAInBC;

  TSAPDailyAccountReader2_03to01 = class 
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_OutAInBC;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_OutAInBC read GetItems;
  end;

  TSAPDailyAccountReader2_03to01_ml = class(TSAPDailyAccountReader2_03to01)
  protected
    procedure Open; override; 
  end;          
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_Xout = packed record 
    sxoutbillno: string; //发货单号
    sxoutdept: string; //发货单位
    snumber: string; //料号
    sname: string; //产品名称
    dqty: Double; //数量
    sorder: string; //订单单号
    sproxy: string; //代理商简称
    sexp: string; //快递公司
    sebillno: string; //电子单号
    smnote: string; //主单备注
    sddate: string; //发货时间
    sstock_fac: string; //仓位
    sdate: string; //过账
    snote: string; //备注

    sstock_mz: string; //仓位
  end;
  pDailyAccount_Xout = ^TDailyAccount_Xout;

  TSAPDailyAccountReader2_xout = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): pDailyAccount_Xout;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: pDailyAccount_Xout read GetItems;
  end;

  TSAPDailyAccountReader2_xout_ml = class(TSAPDailyAccountReader2_xout)
  protected
    procedure Open; override; 
  end;   

            
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TDailyAccount_sout = packed record 
    sicmo: string;
    dt: TDateTime;
    dtCheck: TDateTime;
    scostnumber: string;
    scostname: string;
    snote: string;
    snumber: string;
    sname: string;
    dqty: Double;
    sstock_ml: string;
    sstock: string;
    sbillno: string;
    sdept: string;
    suse: string;
    sbatchno: string;
    schecker: string;
    scheckflag: string;
    sbiller: string;
    sedi: string;

    sfac: string;
    snumber_wt: string;
    dicmoqty: Double;
    snote1: string;
    dtout: TDateTime;
    snumber_child: string;
    sname_child: string;
    dqtyout: Double;
    sstock_wt: string;
    sbomusage: string;
    snote2: string;
    sicmotype: string;


    sstock_yd: string;
    dusage: string;
  end;
  PDailyAccount_sout = ^TDailyAccount_sout;

  TSAPDailyAccountReader2_sout = class
  private         
    FList: TList;
    fsheet: string;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PDailyAccount_sout;
  protected
    FStockMZ2FacReader: TStockMZ2FacReader;
    procedure Open; virtual; abstract;
  public
    constructor Create(const sfile: string; const ssheet: string;
      aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PDailyAccount_sout read GetItems;
  end;

  TSAPDailyAccountReader2_sout_ml = class(TSAPDailyAccountReader2_sout)
  protected
    procedure Open; override; 
  end;   

  TSAPDailyAccountReader2_sout_wt = class(TSAPDailyAccountReader2_sout)
  protected
    procedure Open; override; 
  end;

  TSAPDailyAccountReader2_sout_yd = class(TSAPDailyAccountReader2_sout)
  protected
    procedure Open; override;
  end;
                
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TZPP_PRDORD_004Record = packed record
    sicmobillno_fac: string; // 代工厂工单
    sicmobillno: string; //订单
  end;
  PZPP_PRDORD_004Record = ^TZPP_PRDORD_004Record;
         
  TZPP_PRDORD_004Reader = class
  private         
    FList: TList; 
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PZPP_PRDORD_004Record;
  protected 
    procedure Open;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PZPP_PRDORD_004Record read GetItems;
    function ICMOBillno2fac(const smz: string): string;
  end;
                 
               
  //////////////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////////////

  TCPINmz2facRecord = packed record
    scpinbillno_fac: string; // 代工厂工单
    scpinbillno: string; //订单
    sMvT: string;
  end;
  PCPINmz2facRecord = ^TCPINmz2facRecord;

  TCPINmz2facReader = class
  private
    FList: TList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PCPINmz2facRecord;
  protected
    procedure Open;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PCPINmz2facRecord read GetItems;
    function cpin_mz2fac(const smz: string): string;
  end;

      
implementation
 
         
{ TSAPDailyAccountReader2_winB }

constructor TSAPDailyAccountReader2_winB.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_winB.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_winB.Clear;
var
  i: Integer;
  p: PDailyAccount_winB;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_winB(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_winB.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_winB.GetItems(i: Integer): PDailyAccount_winB;
begin
  Result := PDailyAccount_winB(FList[i]);
end;

procedure TSAPDailyAccountReader2_winB.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_winB_ml }

procedure TSAPDailyAccountReader2_winB_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_winB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
           
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if ADOTabXLS.FieldByName('物料长代码').AsString = '' then
      begin
        ADOTabXLS.Next;
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;

      if ADOTabXLS.FindField('实收数量') <> nil then
      begin
        aSAPOPOAllocPtr^.dQty  := ADOTabXLS.FieldByName('实收数量').AsFloat;
      end
      else
      begin
        aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('数量').AsFloat;
      end;

      sdt := ADOTabXLS.FieldByName('日期').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('用途').AsString;
      if ADOTabXLS.FindField('供货单位') <> nil then
      begin
        aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供货单位').AsString;
      end
      else if ADOTabXLS.FindField('供应商') <> nil then
      begin
        aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供应商').AsString;
      end;

      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;

      if ADOTabXLS.FindField('退料仓库') <> nil then
      begin
        aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('退料仓库').AsString;
      end
      else if ADOTabXLS.FindField('收料仓库') <> nil then
      begin
        aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('收料仓库').AsString;
      end;

      if ADOTabXLS.FindField('摘要') <> nil then
      begin
        aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('摘要').AsString;
      end;
      
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('关闭标志') >= 0 then
        aSAPOPOAllocPtr^.sclose := ADOTabXLS.FieldByName('关闭标志').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('部门') >= 0 then
        aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('部门').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('检验方式') >= 0 then
        aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('检验方式').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('源单单号') >= 0 then
        aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('源单单号').AsString;

      if aSAPOPOAllocPtr^.sstock_ml = '原材料待检仓' then
      begin
        if aSAPOPOAllocPtr^.suse = '量产' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力原材料仓';
        end
        else if aSAPOPOAllocPtr^.suse = '推广' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力推广仓';
        end
        else if aSAPOPOAllocPtr^.suse = '试产' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力RM试产原料仓';
        end
        else
        begin
          aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
        end;
      end
      else
      begin
        aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
      end;
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

{ TSAPDailyAccountReader2_winB_wt }

procedure TSAPDailyAccountReader2_winB_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_winB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
           
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('移动原因描述').AsString <> '客供不良退货') and
        (ADOTabXLS.FieldByName('移动原因描述').AsString <> '客供入库') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;


      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_desc_wt);

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

 
          

{ TSAPDailyAccountReader2_winB_yd }

procedure TSAPDailyAccountReader2_winB_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_winB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
           
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;


      if ADOTabXLS.FieldByName('单据编号').AsString = '' then
      begin      
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('实收数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      if sdt <> '' then
        aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt)
      else aSAPOPOAllocPtr^.dtCheck := 0;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供应商').AsString;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('收料仓库').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('摘要').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;

 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_yd);

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

 



 
         
{ TSAPDailyAccountReader2_RTV }

constructor TSAPDailyAccountReader2_RTV.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_RTV.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_RTV.Clear;
var
  i: Integer;
  p: PDailyAccount_RTV;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_RTV(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_RTV.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_RTV.GetItems(i: Integer): PDailyAccount_RTV;
begin
  Result := PDailyAccount_RTV(FList[i]);
end;

procedure TSAPDailyAccountReader2_RTV.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_RTV_ml }

procedure TSAPDailyAccountReader2_RTV_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_RTV;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
           
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_RTV);
      FList.Add(aSAPOPOAllocPtr);
                                                                                
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供货单位').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('部门').AsString;
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('用途').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('单位').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('退料仓库').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('源单单号').AsString;   
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;                                                                                 
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;


      if aSAPOPOAllocPtr^.sstock_ml = '原材料待检仓' then
      begin
        if aSAPOPOAllocPtr^.suse = '量产' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力原材料仓';
        end
        else if aSAPOPOAllocPtr^.suse = '推广' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力推广仓';
        end
        else if aSAPOPOAllocPtr^.suse = '试产' then
        begin
          aSAPOPOAllocPtr^.sstock := '魅力RM试产原料仓';
        end
        else
        begin
          aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
        end;
      end
      else
      begin
        aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
      end;
        
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;



 

         
{ TSAPDailyAccountReader2_cpin }

constructor TSAPDailyAccountReader2_cpin.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_cpin.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_cpin.Clear;
var
  i: Integer;
  p: PDailyAccount_cpin;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_cpin(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_cpin.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_cpin.GetItems(i: Integer): PDailyAccount_cpin;
begin
  Result := PDailyAccount_cpin(FList[i]);
end;

procedure TSAPDailyAccountReader2_cpin.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_cpin_ml }

procedure TSAPDailyAccountReader2_cpin_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_cpin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产任务单号').AsString;

      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('实收数量').AsFloat;
      aSAPOPOAllocPtr^.sbatchno := ADOTabXLS.FieldByName('批号').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('收货仓库').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('交货单位').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.schecker := ADOTabXLS.FieldByName('审核人').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sbackflush := ADOTabXLS.FieldByName('倒冲标志').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;      

{ TSAPDailyAccountReader2_cpin_wt }

procedure TSAPDailyAccountReader2_cpin_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_cpin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      if sdt = '' then
      begin
        aSAPOPOAllocPtr^.dtbill := 0
      end
      else
      begin
        aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      end;
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;
 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_desc_wt);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
           

{ TSAPDailyAccountReader2_cpin_yd }

procedure TSAPDailyAccountReader2_cpin_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_cpin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('工单号').AsString;
      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('代工厂').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('成品料号').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('成品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('入库数量').AsFloat;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('收货仓库').AsString;


      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_yd);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
      
     

         
{ TSAPDailyAccountReader2_qin }

constructor TSAPDailyAccountReader2_qin.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_qin.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_qin.Clear;
var
  i: Integer;
  p: PDailyAccount_qin;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_qin(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPDailyAccountReader2_qin.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_qin.GetItems(i: Integer): PDailyAccount_qin;
begin
  Result := PDailyAccount_qin(FList[i]);
end;

procedure TSAPDailyAccountReader2_qin.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_qin_ml }

procedure TSAPDailyAccountReader2_qin_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_qin);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').asstring;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;
      //SAP数量
      //差异
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('用途').AsString;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供应商').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('收料仓库').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('摘要').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('关闭标志').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('部门').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('检验方式').AsString;
      aSAPOPOAllocPtr^.sedit := ADOTabXLS.FieldByName('EDI提交').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('源单单号').AsString; 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
          

{ TSAPDailyAccountReader2_qin_wt }

procedure TSAPDailyAccountReader2_qin_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      //ADOTabXLS.TableName:='['+fsheet+'$]';
      ADOTabXLS.TableName:='[来料入库$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('移动原因描述').AsString <> '客供赠品') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;


      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_desc_wt);

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

          

{ TSAPDailyAccountReader2_qin_yd }

procedure TSAPDailyAccountReader2_qin_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qin;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      if ADOTabXLS.FieldByName('单据编号').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供应商').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('收料仓库').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('摘要').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('关闭标志').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('部门').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('检验方式').AsString;

 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_yd);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
     

         
{ TSAPDailyAccountReader2_qout }

constructor TSAPDailyAccountReader2_qout.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_qout.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_qout.Clear;
var
  i: Integer;
  p: PDailyAccount_qout;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_qout(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPDailyAccountReader2_qout.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_qout.GetItems(i: Integer): PDailyAccount_qout;
begin
  Result := PDailyAccount_qout(FList[i]);
end;

procedure TSAPDailyAccountReader2_qout.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_qout_ml }

procedure TSAPDailyAccountReader2_qout_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_qout);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品长代码').asstring;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.dqty := - ADOTabXLS.FieldByName('数量').AsFloat;


      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('发货仓库').AsString;
      aSAPOPOAllocPtr^.sdetp := ADOTabXLS.FieldByName('领料部门').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('用途1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('单位').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('出库类别').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('用途2').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;           

{ TSAPDailyAccountReader2_qout_wt }

procedure TSAPDailyAccountReader2_qout_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_qout);
      FList.Add(aSAPOPOAllocPtr);



      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;
 
 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_desc_wt);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;                   

{ TSAPDailyAccountReader2_qout_yd }

procedure TSAPDailyAccountReader2_qout_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_qout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_qin_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      if ADOTabXLS.FieldByName('产品长代码').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('发货仓库').AsString;
      aSAPOPOAllocPtr^.sdetp := ADOTabXLS.FieldByName('领料部门').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('用途1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('单位').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('出库类别').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('用途2').AsString;
 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_yd);
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

         
{ TSAPDailyAccountReader2_DB }

constructor TSAPDailyAccountReader2_DB.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_DB.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_DB.Clear;
var
  i: Integer;
  p: PDailyAccount_DB;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_DB(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_DB.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_DB.GetItems(i: Integer): PDailyAccount_DB;
begin
  Result := PDailyAccount_DB(FList[i]);
end;

procedure TSAPDailyAccountReader2_DB.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_DB_ML }

procedure TSAPDailyAccountReader2_DB_ML.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_DB);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('调出仓库').AsString;
      aSAPOPOAllocPtr^.sstock_in_ml := ADOTabXLS.FieldByName('调入仓库').AsString;       
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('调拨数量').AsFloat;

      sdt := ADOTabXLS.FieldByName('日期').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);

//      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;


      aSAPOPOAllocPtr^.sstock_in := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_in_ml);
      aSAPOPOAllocPtr^.sstock_out := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_out_ml);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
        

{ TSAPDailyAccountReader2_DB_wt }

procedure TSAPDailyAccountReader2_DB_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_DB);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.bCalc := False;

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      if UpperCase(Copy(aSAPOPOAllocPtr^.snumber, 1, 3)) = 'KMZ' then
      begin
        aSAPOPOAllocPtr^.snumber := Copy(aSAPOPOAllocPtr^.snumber, 4, Length(aSAPOPOAllocPtr^.snumber) - 3);
      end;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;
                                                                                                
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_wt);
      aSAPOPOAllocPtr^.sstock_desc := FStockMZ2FacReader.ToName(aSAPOPOAllocPtr^.sstock);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

function TSAPDailyAccountReader2_DB_wt.GetItem2(aDailyAccount_DBPtr: PDailyAccount_DB): PDailyAccount_DB;
var
  i: Integer;
  p: PDailyAccount_DB;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_DB(FList[i]);
    if p^.bCalc then Continue;
    if (p^.sbillno = aDailyAccount_DBPtr^.sbillno)
      and (p^.snumber = aDailyAccount_DBPtr^.snumber)
      and (p^.sstock_desc_wt <> aDailyAccount_DBPtr^.sstock_desc_wt)
      and (p^.dQty + aDailyAccount_DBPtr^.dQty = 0) then
    begin
      Result := p;
      Break;
    end;  
  end;
end;
     

        

{ TSAPDailyAccountReader2_DB_yd }

procedure TSAPDailyAccountReader2_DB_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_DB);
      FList.Add(aSAPOPOAllocPtr);


      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sbillno := StringReplace(aSAPOPOAllocPtr^.sbillno, '?', '', [rfReplaceAll]);
      aSAPOPOAllocPtr^.sstock_out_yd := ADOTabXLS.FieldByName('调出仓库').AsString;
      aSAPOPOAllocPtr^.sstock_in_yd := ADOTabXLS.FieldByName('调入仓库').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('调拨数量').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;

      aSAPOPOAllocPtr^.sstock_in := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_in_yd);
      aSAPOPOAllocPtr^.sstock_out := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_out_yd);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
         

         
{ TSAPDailyAccountReader2_DB_in }

constructor TSAPDailyAccountReader2_DB_in.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_DB_in.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_DB_in.Clear;
var
  i: Integer;
  p: PDailyAccount_DB_in;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_DB_in(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_DB_in.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_DB_in.GetItems(i: Integer): PDailyAccount_DB_in;
begin
  Result := PDailyAccount_DB_in(FList[i]);
end;

procedure TSAPDailyAccountReader2_DB_in.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_DB_in_ML }

procedure TSAPDailyAccountReader2_DB_in_ML.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB_in;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('用途').AsString;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('供应商').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sstock_in_ml := ADOTabXLS.FieldByName('收料仓库').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('摘要').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('关闭标志').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('部门').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('检验方式').AsString;  
      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('源单单号').AsString;
       
      aSAPOPOAllocPtr^.sstock_in := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_in_ml);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;                        


{ TSAPDailyAccountReader2_DB_in_wt }

procedure TSAPDailyAccountReader2_DB_in_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB_in;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      //ADOTabXLS.TableName:='['+fsheet+'$]';    
      ADOTabXLS.TableName:='[来料入库$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
    
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('移动原因描述').AsString <> '客供工厂间调拨') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('工厂名称').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('物料凭证').AsString;
      sdt := ADOTabXLS.FieldByName('过帐日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('制造商代码').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('制造商描述').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('移动类型').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('移动原因').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('规格型号').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('过账数量').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('基本计量单位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('凭证抬头文本').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('工作中心名称').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('项目文本').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('单据项目号').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('库存地点').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('工厂编号').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('移动原因描述').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('物料组').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('订单类型').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('生产订单数量').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('异动状况').AsString;
      sdt := ADOTabXLS.FieldByName('单据日期').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('单据数量').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产订单号').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('仓储地点的描述').AsString;


      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_desc_wt);

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;                        

        
                 

{ TSAPDailyAccountReader2_DB_in_yd }

procedure TSAPDailyAccountReader2_DB_in_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB_in;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;

  icolDate: Integer;//  sdt := ADOTabXLS.FieldByName('日期').AsString; 
  icolBillno: Integer; //  aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
  icolStockOut: Integer; //  aSAPOPOAllocPtr^.sstockno_out_yd := ADOTabXLS.FieldByName('调出仓库').AsString;
  icolStockIn: Integer; //  aSAPOPOAllocPtr^.sstockno_in_yd := ADOTabXLS.FieldByName('调入仓库').AsString;
  icolNumber: Integer; //  aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
  icolName: Integer; //  aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
  icolQty: Integer;  //  aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('调拨数量').AsFloat;
  sSheet: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;




  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;

        if sSheet <> fsheet then Continue;

        irow := 1;
        icolDate := IndexOfCol(ExcelApp, irow, '日期');
        icolBillno := IndexOfCol(ExcelApp, irow, '单据编号');
        icolStockOut := IndexOfCol(ExcelApp, irow, '调出仓库');
        icolStockIn := IndexOfCol(ExcelApp, irow, '调入仓库');
        icolNumber := IndexOfCol(ExcelApp, irow, '物料长代码');
        icolName := IndexOfCol(ExcelApp, irow, '物料名称');
        icolQty := IndexOfCol(ExcelApp, irow, '调拨数量');        

        irow := irow + 1;
        snumber := ExcelApp.Cells[irow, icolNumber].Value;
        while snumber <> '' do
        begin

          aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
          FList.Add(aSAPOPOAllocPtr);
                                                                         
          sdt := ExcelApp.Cells[irow, icolDate].Value;
          aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
          aSAPOPOAllocPtr^.sbillno := ExcelApp.Cells[irow, icolBillno].Value;
          aSAPOPOAllocPtr^.sstockno_out_yd := ExcelApp.Cells[irow, icolStockOut].Value;
          aSAPOPOAllocPtr^.sstockno_in_yd := ExcelApp.Cells[irow, icolStockIn].Value;
          aSAPOPOAllocPtr^.snumber := snumber;
          aSAPOPOAllocPtr^.sname := ExcelApp.Cells[irow, icolName].Value;
          aSAPOPOAllocPtr^.dQty := ExcelApp.Cells[irow, icolQty].Value;

          aSAPOPOAllocPtr^.sstockno_out := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstockno_out_yd);
          aSAPOPOAllocPtr^.sstockno_in := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstockno_in_yd);

        
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, icolNumber].Value;
        end;
        
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;



  (*

  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';


      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
    
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
      FList.Add(aSAPOPOAllocPtr);
                                                                         
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sstockno_out_yd := ADOTabXLS.FieldByName('调出仓库').AsString;
      aSAPOPOAllocPtr^.sstockno_in_yd := ADOTabXLS.FieldByName('调入仓库').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('调拨数量').AsFloat;

      aSAPOPOAllocPtr^.sstockno_out := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstockno_out_yd);
      aSAPOPOAllocPtr^.sstockno_in := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstockno_in_yd);

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
  *)
end;                        

        

         
{ TSAPDailyAccountReader2_coois }

constructor TSAPDailyAccountReader2_coois.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_coois.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_coois.Clear;
var
  i: Integer;
  p: PDailyAccount_coois;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_coois(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPDailyAccountReader2_coois.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_coois.GetItems(i: Integer): PDailyAccount_coois;
begin
  Result := PDailyAccount_coois(FList[i]);
end;

procedure TSAPDailyAccountReader2_coois.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TSAPDailyAccountReader2_coois.Fac2MZBillno(const sicmobillno: string): string;
var
  i: Integer;
  ptrDailyAccount_coois: PDailyAccount_coois;
  sbillno: string;
  idx: Integer;
begin
  Result := '';
  for i := 0 to self.Count - 1 do
  begin
    ptrDailyAccount_coois := Items[i];


    sbillno := ptrDailyAccount_coois^.sbillno_fac;
    idx := Pos('-', sbillno);
    if idx > 0 then
    begin
      sbillno := Copy(sbillno, 1, idx - 1);
    end;
                 
    idx := Pos('/', sbillno);
    if idx > 0 then
    begin
      sbillno := Copy(sbillno, 1, idx - 1);
    end;
                   
    if Copy(sbillno, 1, 3) = 'NWT' then
    begin
      sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
    end; 
          
    if Copy(sbillno, 1, 2) = 'WT' then
    begin
      sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
    end; 

    if sbillno = sicmobillno then
    begin
      Result := ptrDailyAccount_coois^.sbillno;
      Break;
    end;
  end;
end;

procedure TSAPDailyAccountReader2_coois.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_coois;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_coois);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sbillno_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('订单').AsString;
      aSAPOPOAllocPtr^.scategory := ADOTabXLS.FieldByName('类型').AsString;
      aSAPOPOAllocPtr^.dtfac := ADOTabXLS.FieldByName('代工厂单据日期').AsDateTime;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单人').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.dtFinish := ADOTabXLS.FieldByName('计划完工').AsDateTime;
      aSAPOPOAllocPtr^.sbillno_plan := ''; //ADOTabXLS.FieldByName('计划订单').AsString;
      aSAPOPOAllocPtr^.dqtyorder := ADOTabXLS.FieldByName('订单数量').AsFloat;
      aSAPOPOAllocPtr^.sBUn := ADOTabXLS.FieldByName('BUn').AsString;
      aSAPOPOAllocPtr^.sstockname := ADOTabXLS.FieldByName('库位').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('物料1').AsString;
      aSAPOPOAllocPtr^.dtneed := ADOTabXLS.FieldByName('需求日期').AsDateTime;
      aSAPOPOAllocPtr^.dqtyneed := ADOTabXLS.FieldByName('需求量').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('计').AsString;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('工厂').AsString;
      aSAPOPOAllocPtr^.sFix := ADOTabXLS.FieldByName('Fix').AsString;
      aSAPOPOAllocPtr^.dtChangeDate := ADOTabXLS.FieldByName('变更日期').AsDateTime;
      aSAPOPOAllocPtr^.dtChangeTime := ADOTabXLS.FieldByName('变更时间').AsDateTime;
      aSAPOPOAllocPtr^.dQtyIn := ADOTabXLS.FieldByName('收货数量').AsFloat;
           
      //aSAPOPOAllocPtr^.sbillno_ml := FSAPDailyAccountReader2_icmo_mz2fac.billno_mz2fac(aSAPOPOAllocPtr^.sbillno);

      aSAPOPOAllocPtr^.bCalc := False;
      aSAPOPOAllocPtr^.sMatchType := '';
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;                        
     

         
{ TSAPDailyAccountReader2_FacICMO2MZICMO }

constructor TSAPDailyAccountReader2_FacICMO2MZICMO.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_FacICMO2MZICMO.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_FacICMO2MZICMO.Clear;
var
  i: Integer;
  p: PDailyAccount_coois;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_coois(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_FacICMO2MZICMO.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_FacICMO2MZICMO.GetItems(i: Integer): PDailyAccount_FacICMO2MZICMO;
begin
  Result := PDailyAccount_FacICMO2MZICMO(FList[i]);
end;

procedure TSAPDailyAccountReader2_FacICMO2MZICMO.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TSAPDailyAccountReader2_FacICMO2MZICMO.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_FacICMO2MZICMO;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_FacICMO2MZICMO);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmo_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('订单').AsString;
      aSAPOPOAllocPtr^.stype := ADOTabXLS.FieldByName('类型').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('代工厂代号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料描述').AsString;
      aSAPOPOAllocPtr^.slang := ADOTabXLS.FieldByName('语言').AsString;
      aSAPOPOAllocPtr^.swwpo := ADOTabXLS.FieldByName('委外采购订单').AsString;
      aSAPOPOAllocPtr^.ssourceorder := ADOTabXLS.FieldByName('来源订单号').AsString;
      sdt := ADOTabXLS.FieldByName('代工厂单据日期').AsString;
      aSAPOPOAllocPtr^.dtFac := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('基本完成日期').AsString;;
      aSAPOPOAllocPtr^.dtend := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('基本开始日期').AsString;;
      aSAPOPOAllocPtr^.dtbegin := myStrToDateTime(sdt);
 
           
      //aSAPOPOAllocPtr^.sbillno_ml := FSAPDailyAccountReader2_icmo_mz2fac.billno_mz2fac(aSAPOPOAllocPtr^.sbillno);

      aSAPOPOAllocPtr^.bCalc := False;
      aSAPOPOAllocPtr^.sMatchType := '';
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;                        
     


         
{ TSAPDailyAccountReader2_icmo_mz2fac }

constructor TSAPDailyAccountReader2_icmo_mz2fac.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_icmo_mz2fac.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_icmo_mz2fac.Clear;
var
  i: Integer;
  p: PDailyAccount_icmo_mz2fac;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_icmo_mz2fac(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPDailyAccountReader2_icmo_mz2fac.billno_mz2fac(const sbillno_mz: string): string;
var
  i: Integer;
  p: PDailyAccount_icmo_mz2fac;
begin
  Result := '';
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_icmo_mz2fac(FList[i]);
    if p^.sicmobillno = sbillno_mz then
    begin
      Result := p^.sicmolbillno_fac;
      Break;
    end;
  end;
end;  

function TSAPDailyAccountReader2_icmo_mz2fac.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_icmo_mz2fac.GetItems(i: Integer): PDailyAccount_icmo_mz2fac;
begin
  Result := PDailyAccount_icmo_mz2fac(FList[i]);
end;

procedure TSAPDailyAccountReader2_icmo_mz2fac.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;
  
procedure TSAPDailyAccountReader2_icmo_mz2fac.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_icmo_mz2fac;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PDailyAccount_icmo_mz2fac);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmolbillno_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('订单').AsString;
      aSAPOPOAllocPtr^.stype := ADOTabXLS.FieldByName('类型').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('代工厂代号').AsString;
      aSAPOPOAllocPtr^.dtdate_fac := ADOTabXLS.FieldByName('代工厂单据日期').AsDateTime;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单人').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('来源订单号').AsString;
      aSAPOPOAllocPtr^.swwcontract1 := ADOTabXLS.FieldByName('委外合同1').AsString;
      aSAPOPOAllocPtr^.dqty_contract_alloc1 := ADOTabXLS.FieldByName('合同分配数量1').AsFloat;
      aSAPOPOAllocPtr^.swwcontract2 := ADOTabXLS.FieldByName('委外合同2').AsString;
      aSAPOPOAllocPtr^.dqty_contract_alloc1 := ADOTabXLS.FieldByName('合同分配数量2').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sall_transfer_flag := ADOTabXLS.FieldByName('完全转换标志').AsString;
      aSAPOPOAllocPtr^.dtChangeDate := ADOTabXLS.FieldByName('变更日期').AsDateTime;
      aSAPOPOAllocPtr^.dtChangeTime := ADOTabXLS.FieldByName('变更时间').AsDateTime; 
            
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
     
        
         

         
{ TSAPDailyAccountReader2_PPBOM }

constructor TSAPDailyAccountReader2_PPBOM.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_PPBOM.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_PPBOM.Clear;
var
  i: Integer;
  p: PDailyAccount_PPBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_PPBom(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_PPBOM.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_PPBOM.GetItems(i: Integer): PDailyAccount_PPBOM;
begin
  Result := PDailyAccount_PPBom(FList[i]);
end;

procedure TSAPDailyAccountReader2_PPBOM.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;        

{ TSAPDailyAccountReader2_PPBOM_ml }  

procedure TSAPDailyAccountReader2_PPBOM_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_PPBom;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_PPBom);
      FList.Add(aSAPOPOAllocPtr);

      sdt := ADOTabXLS.FieldByName('制单日期').AsString;

      aSAPOPOAllocPtr^.dtdate := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('生产/委外订单号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('生产数量').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('生产投料单号').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('子项物料长代码').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('子项物料名称').AsString;
      aSAPOPOAllocPtr^.dqtyplan := ADOTabXLS.FieldByName('计划投料数量').AsFloat;
      aSAPOPOAllocPtr^.dqtyshould := ADOTabXLS.FieldByName('应发数量').AsFloat;
      aSAPOPOAllocPtr^.sstockname := ADOTabXLS.FieldByName('仓库').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('单位用量').AsFloat;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sworkshopname := ADOTabXLS.FieldByName('生产车间').AsString;
      aSAPOPOAllocPtr^.sedi := ''; //ADOTabXLS.FieldByName('EDI提交').AsString;
 
//      aSAPOPOAllocPtr^.sstock_in := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_in_ml);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;         

{ TSAPDailyAccountReader2_PPBOM_wt }

procedure TSAPDailyAccountReader2_PPBOM_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;                      
  aSAPOPOAlloc: TDailyAccount_PPBom;
  aSAPOPOAllocPtr: PDailyAccount_PPBom;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;

  i: Integer;
  bFound: Boolean;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if ADOTabXLS.FieldByName('生产订单').AsString = '' then
      begin            
        ADOTabXLS.Next;
        Continue;
      end;
      
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;


      aSAPOPOAlloc.sfacname := ADOTabXLS.FieldByName('加工厂描述').AsString;
      aSAPOPOAlloc.sfac := ADOTabXLS.FieldByName('工厂代码').AsString;
      aSAPOPOAlloc.sicmobillno := ADOTabXLS.FieldByName('生产订单').AsString;
      aSAPOPOAlloc.sicmotye := ADOTabXLS.FieldByName('订单类型').AsString;
      sdt := ADOTabXLS.FieldByName('下达日期').AsString;
      aSAPOPOAlloc.dtRelease := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('结案日期').AsString;
      if sdt = '' then
      begin
        aSAPOPOAlloc.dtClose := 0;
      end
      else
      begin
        aSAPOPOAlloc.dtClose := myStrToDateTime(sdt);
      end;
      sdt := ADOTabXLS.FieldByName('订单开始日期').AsString;
      aSAPOPOAlloc.dtBegin := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('订单完成日期').AsString;
      aSAPOPOAlloc.dtEnd := myStrToDateTime(sdt);
      aSAPOPOAlloc.splanbillno := ''; //ADOTabXLS.FieldByName('计划订单').AsString;
      aSAPOPOAlloc.splanbillno_mz := ADOTabXLS.FieldByName('魅族计划订单').AsString;
      aSAPOPOAlloc.snumber_wt := ADOTabXLS.FieldByName('闻泰父料号').AsString;
      aSAPOPOAlloc.snumber := ADOTabXLS.FieldByName('客户父料号').AsString;
      aSAPOPOAlloc.svItemFlag := ADOTabXLS.FieldByName('虚拟项目标识').AsString;
      aSAPOPOAlloc.sname := ADOTabXLS.FieldByName('物料描述').AsString;
      aSAPOPOAlloc.sItemCode := ADOTabXLS.FieldByName('项目代码').AsString;
      aSAPOPOAlloc.dICMOQty := ADOTabXLS.FieldByName('工单数量').AsFloat;
      aSAPOPOAlloc.snote1 := ADOTabXLS.FieldByName('备注1').AsString;
      aSAPOPOAlloc.iChangeCount := ADOTabXLS.FieldByName('变更次数').AsString;
      aSAPOPOAlloc.irowitem := ADOTabXLS.FieldByName('行项目').AsString;
      aSAPOPOAlloc.snumber_item_wt := ADOTabXLS.FieldByName('闻泰子物料编码').AsString;
      aSAPOPOAlloc.snumber_item := ADOTabXLS.FieldByName('客户子物料编码').AsString;
      aSAPOPOAlloc.sname_item := ADOTabXLS.FieldByName('物料描述1').AsString;
      aSAPOPOAlloc.dqtyplan := ADOTabXLS.FieldByName('需求量').AsFloat;
      aSAPOPOAlloc.dqtyout := ADOTabXLS.FieldByName('已投料数量').AsFloat;
      aSAPOPOAlloc.sstockname_wt := ADOTabXLS.FieldByName('库位').AsString;
      aSAPOPOAlloc.dqty0 := ADOTabXLS.FieldByName('变更前数量').AsFloat;
      aSAPOPOAlloc.sgroup := ADOTabXLS.FieldByName('替代组').AsString;
      aSAPOPOAlloc.sprioriry := ADOTabXLS.FieldByName('优先级').AsString;
      aSAPOPOAlloc.dper := ADOTabXLS.FieldByName('替代比例').AsFloat;
      aSAPOPOAlloc.dqtyshould := ADOTabXLS.FieldByName('总需求量').AsFloat;
      aSAPOPOAlloc.sunit := ADOTabXLS.FieldByName('基本单位').AsString;
      aSAPOPOAlloc.snote2 := ADOTabXLS.FieldByName('备注2').AsString;
      aSAPOPOAlloc.schangelog := ADOTabXLS.FieldByName('变更情况').AsString;

 
      aSAPOPOAlloc.sstockname := FStockMZ2FacReader.Fac2MZ(aSAPOPOAlloc.sstockname_wt);

      //if aSAPOPOAlloc.dqtyplan > 0 then
      begin
        bFound := False;
        for i := 0 to FList.Count - 1 do
        begin
          aSAPOPOAllocPtr := PDailyAccount_PPBom(FList[i]);
          if (aSAPOPOAllocPtr^.sicmobillno = aSAPOPOAlloc.sicmobillno)
            and (aSAPOPOAllocPtr^.snumber_item_wt = aSAPOPOAlloc.snumber_item_wt) then
          begin
            //aSAPOPOAllocPtr^.dqtyplan := aSAPOPOAllocPtr^.dqtyplan + aSAPOPOAlloc.dqtyplan;
            bFound := True;
            Break;
          end;
        end;

        if not bFound then
        begin
          aSAPOPOAllocPtr := New(PDailyAccount_PPBom);
          aSAPOPOAllocPtr^ := aSAPOPOAlloc;
          FList.Add(aSAPOPOAllocPtr);
        end;
      end;
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

{ TSAPDailyAccountReader2_PPBOM_yd }

procedure TSAPDailyAccountReader2_PPBOM_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_PPBom;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;                   
      sdt := ADOTabXLS.FieldByName('制单日期').AsString;
      if sdt = '' then Break;

      aSAPOPOAllocPtr := New(PDailyAccount_PPBom);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.dtdate := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('生产/委外订单号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('生产数量').AsFloat;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('生产投料单号').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('子项物料长代码').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('子项物料名称').AsString;
      aSAPOPOAllocPtr^.dqtyplan := ADOTabXLS.FieldByName('计划投料数量').AsFloat;
      aSAPOPOAllocPtr^.dqtyshould := ADOTabXLS.FieldByName('应发数量').AsFloat;
      aSAPOPOAllocPtr^.sstockname_yd := ADOTabXLS.FieldByName('仓库').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('单位用量').AsFloat;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sworkshopname := ADOTabXLS.FieldByName('生产车间').AsString;      


      aSAPOPOAllocPtr^.sstockname := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstockname_yd);


      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;


{ TSAPDailyAccountReader2_PPBOMChange }

constructor TSAPDailyAccountReader2_PPBOMChange.Create(const sfile: string; const ssheet: string;
  aStockMZ2FacReader: TStockMZ2FacReader; aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_PPBOMChange.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_PPBOMChange.Clear;
var
  i: Integer;
  p: Pointer;
begin
  for i:= 0 to FList.Count - 1 do
  begin
    p := FList[i];
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TSAPDailyAccountReader2_PPBOMChange.Log(const str: string);
begin

end;

function TSAPDailyAccountReader2_PPBOMChange.GetCount: Integer;
begin
  Result := FList.Count;
end;
 
{ TSAPDailyAccountReader2_PPBOMChange_mz }

function TSAPDailyAccountReader2_PPBOMChange_mz.GetItems(i: Integer): PDailyAccount_PPBomChange_mz;
begin
  Result := PDailyAccount_PPBomChange_mz(FList[i]);
end;

procedure TSAPDailyAccountReader2_PPBOMChange_mz.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_PPBomChange_mz;
                   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='[Sheet1$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if Trim(ADOTabXLS.FieldByName('变更单号').AsString) = '' then
      begin            
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_PPBomChange_mz);
      aSAPOPOAllocPtr^.bCalc := False;
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('变更单号').AsString;//  : string; //
      aSAPOPOAllocPtr^.sarea := ADOTabXLS.FieldByName('MRP 范围').AsString;//  : string; //
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;//  : string; //
      aSAPOPOAllocPtr^.sorderbillno := ADOTabXLS.FieldByName('订单').AsString;//  : string; //
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料').AsString;//  : string; //
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('计').AsString;//  : string; //
      aSAPOPOAllocPtr^.splanroder := ''; //ADOTabXLS.FieldByName('计划订单').AsString;//  : string; //
      aSAPOPOAllocPtr^.sqtychangeflag := ADOTabXLS.FieldByName('数量变更标志').AsString;//  : string; //
      aSAPOPOAllocPtr^.sreason := ADOTabXLS.FieldByName('变更原因').AsString;//  : string; //
      aSAPOPOAllocPtr^.schangetime := ADOTabXLS.FieldByName('变更时间').AsString;//  : string; //
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('组件').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ITEM := ADOTabXLS.FieldByName('ZTIPP007B-ITEM').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_LGORT := ADOTabXLS.FieldByName('ZTIPP007B-LGORT').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ALPGR := ADOTabXLS.FieldByName('ZTIPP007B-ALPGR').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ALPRF := ADOTabXLS.FieldByName('ZTIPP007B-ALPRF').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_EWAHR := ADOTabXLS.FieldByName('ZTIPP007B-EWAHR').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ITEM_FLAG := ADOTabXLS.FieldByName('ZTIPP007B-ITEM_FLAG').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_REMARK := ADOTabXLS.FieldByName('ZTIPP007B-REMARK').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_UPDKZ := ADOTabXLS.FieldByName('ZTIPP007B-UPDKZ').AsString;//  : string; //
      aSAPOPOAllocPtr^.sicmo_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;//  : string; //
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;//  : Double;//
      aSAPOPOAllocPtr^.dqtyBefore := ADOTabXLS.FieldByName('修改前数量').AsFloat;//  : string; //
      sdt := ADOTabXLS.FieldByName('变更日期').AsString;
      aSAPOPOAllocPtr^.dtChange := myStrToDateTime(sdt);//  : TDateTime;//
      aSAPOPOAllocPtr^.sZTIPP007B_MENGE := ADOTabXLS.FieldByName('ZTIPP007B-MENGE').AsString;//  : string; //
      sdt := ADOTabXLS.FieldByName('需求日期').AsString;
      aSAPOPOAllocPtr^.dtNeed := myStrToDatetime(sdt);//  : TDateTime; //
      aSAPOPOAllocPtr^.sZTIPP007B_MENGE_B := ADOTabXLS.FieldByName('ZTIPP007B-MENGE_B').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_MENGE_T := ADOTabXLS.FieldByName('ZTIPP007B-MENGE_T').AsString;//  : string; //
 
 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;  
     
 
{ TSAPDailyAccountReader2_PPBOMChange_yd }
       
function TSAPDailyAccountReader2_PPBOMChange_yd.GetItems(i: Integer): PDailyAccount_PPBomChange_yd;
begin
  Result := PDailyAccount_PPBomChange_yd(FList[i]);
end;

procedure TSAPDailyAccountReader2_PPBOMChange_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_PPBomChange_yd;
                   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='[投料变更单$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      if ADOTabXLS.FieldByName('产品代码').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_PPBomChange_yd);
      FList.Add(aSAPOPOAllocPtr);
 
      aSAPOPOAllocPtr^.sChangeFlag := ADOTabXLS.FieldByName('变更标志').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('生产投料单号').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('物料代码').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.susage := ADOTabXLS.FieldByName('标准用量').AsString;
      aSAPOPOAllocPtr^.sstock_fac := ADOTabXLS.FieldByName('仓库').AsString;
      aSAPOPOAllocPtr^.sChangeReason := ADOTabXLS.FieldByName('变更原因').AsString; 
      aSAPOPOAllocPtr^.sdt := ADOTabXLS.FieldByName('制单日期').AsString;
      aSAPOPOAllocPtr^.sdtCheck := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.sChangeVer := ADOTabXLS.FieldByName('变更版次').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('计划投料数量').AsFloat;
                                     
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_fac);
 
      ADOTabXLS.Next;
    end; 
    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;  
 
         
{ TSAPDailyAccountReader2_DB_out }

constructor TSAPDailyAccountReader2_DB_out.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_DB_out.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_DB_out.Clear;
var
  i: Integer;
  p: PDailyAccount_DB_out;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_DB_out(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_DB_out.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_DB_out.GetItems(i: Integer): PDailyAccount_DB_out;
begin
  Result := PDailyAccount_DB_out(FList[i]);
end;

procedure TSAPDailyAccountReader2_DB_out.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_DB_out_ML }

procedure TSAPDailyAccountReader2_DB_out_ML.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_DB_out;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if ADOTabXLS.FieldByName('单据编号').AsString = '' then
      begin
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_out);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;        
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);                      
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('发货仓库').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('领料部门').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('用途1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;       
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('单位').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('出库类别').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('用途2').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
 
      aSAPOPOAllocPtr^.sstock_out := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_out_ml);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);                                                                                                      
    FreeAndNil(ADOTabXLS);
  end;
end;

    
         

         
{ TSAPDailyAccountReader2_03to01 }

constructor TSAPDailyAccountReader2_03to01.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_03to01.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_03to01.Clear;
var
  i: Integer;
  p: PDailyAccount_OutAInBC;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_OutAInBC(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_03to01.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_03to01.GetItems(i: Integer): PDailyAccount_OutAInBC;
begin
  Result := PDailyAccount_OutAInBC(FList[i]);
end;

procedure TSAPDailyAccountReader2_03to01.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_03to01_ML }

procedure TSAPDailyAccountReader2_03to01_ML.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_OutAInBC;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      if ADOTabXLS.FieldByName('单据编号').AsString = '' then
      begin    
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_OutAInBC);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('产品长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;        
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('日期').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);                      
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('发货仓库').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('领料部门').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('用途1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;       
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('单位').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('出库类别').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('用途2').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
 
      aSAPOPOAllocPtr^.sstock_out := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_out_ml);

              
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
       

         
{ TSAPDailyAccountReader2_xout }

constructor TSAPDailyAccountReader2_xout.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_xout.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_xout.Clear;
var
  i: Integer;
  p: pDailyAccount_Xout;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := pDailyAccount_Xout(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_xout.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_xout.GetItems(i: Integer): pDailyAccount_Xout;
begin
  Result := pDailyAccount_Xout(FList[i]);
end;

procedure TSAPDailyAccountReader2_xout.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_xout_ml }

procedure TSAPDailyAccountReader2_xout_ml.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_xout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_xout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sxoutbillno := ADOTabXLS.FieldByName('发货单号').AsString;
      aSAPOPOAllocPtr^.sxoutdept := ADOTabXLS.FieldByName('发货单位').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('料号').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('产品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;
      aSAPOPOAllocPtr^.sorder := ADOTabXLS.FieldByName('订单单号').AsString;
      aSAPOPOAllocPtr^.sproxy := ADOTabXLS.FieldByName('代理商简称').AsString;
      aSAPOPOAllocPtr^.sexp := ADOTabXLS.FieldByName('快递公司').AsString;
      aSAPOPOAllocPtr^.sebillno := ADOTabXLS.FieldByName('电子单号').AsString;
      aSAPOPOAllocPtr^.smnote := ADOTabXLS.FieldByName('主单备注').AsString;
      aSAPOPOAllocPtr^.sddate := ADOTabXLS.FieldByName('发货时间').AsString;
      aSAPOPOAllocPtr^.sstock_fac := ADOTabXLS.FieldByName('仓位').AsString;
      aSAPOPOAllocPtr^.sdate := ADOTabXLS.FieldByName('过账').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;

      aSAPOPOAllocPtr^.sstock_mz := aSAPOPOAllocPtr^.sstock_fac; 
          
       
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;        

  
         
{ TSAPDailyAccountReader2_sout }

constructor TSAPDailyAccountReader2_sout.Create(const sfile: string;
  const ssheet: string; aStockMZ2FacReader: TStockMZ2FacReader;
  aLogEvent: TLogEvent = nil);
begin
  fsheet := ssheet;
  FFile := sfile;
  FStockMZ2FacReader := aStockMZ2FacReader;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPDailyAccountReader2_sout.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPDailyAccountReader2_sout.Clear;
var
  i: Integer;
  p: PDailyAccount_sout;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDailyAccount_sout(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPDailyAccountReader2_sout.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPDailyAccountReader2_sout.GetItems(i: Integer): PDailyAccount_sout;
begin
  Result := PDailyAccount_sout(FList[i]);
end;

procedure TSAPDailyAccountReader2_sout.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

{ TSAPDailyAccountReader2_sout_ML }

procedure TSAPDailyAccountReader2_sout_ML.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_sout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_winB_ml.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('生产任务单号').AsString;

      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('审核日期').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.scostnumber := ADOTabXLS.FieldByName('成本对象代码').AsString;
      aSAPOPOAllocPtr^.scostname := ADOTabXLS.FieldByName('成本对象').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('物料长代码').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('实发数量').AsFloat;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('发料仓库').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('领料部门').AsString;
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('领料用途').AsString;
      aSAPOPOAllocPtr^.sbatchno := ADOTabXLS.FieldByName('批号').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('审核人').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('审核标志').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('制单').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;


      if ADOTabXLS.FieldDefList.IndexOf('EDI提交') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI提交').AsString;
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_ml);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;        

{ TSAPDailyAccountReader2_sout_wt }

procedure TSAPDailyAccountReader2_sout_wt.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_sout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_sout_wt.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('代工厂').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('工单号').AsString;
      sdt := ADOTabXLS.FieldByName('日期').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('成品料号').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('成品名称').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('工单数量').AsFloat;
      aSAPOPOAllocPtr^.snote1 := ADOTabXLS.FieldByName('备注1').AsString;
      sdt := ADOTabXLS.FieldByName('领料日期').AsString;
      aSAPOPOAllocPtr^.dtout := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_child := ADOTabXLS.FieldByName('子项料号').AsString;
      aSAPOPOAllocPtr^.sname_child := ADOTabXLS.FieldByName('子项名称').AsString;
      aSAPOPOAllocPtr^.dqtyout := ADOTabXLS.FieldByName('领料数量').AsFloat;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('发料仓库').AsString;
      aSAPOPOAllocPtr^.sbomusage := ADOTabXLS.FieldByName('BOM用量').AsString;
      aSAPOPOAllocPtr^.snote2 := ADOTabXLS.FieldByName('备注2').AsString;
      aSAPOPOAllocPtr^.sicmotype := ADOTabXLS.FieldByName('工单类型').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
 
 
      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ(aSAPOPOAllocPtr^.sstock_wt);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

{ TSAPDailyAccountReader2_sout_yd }

procedure TSAPDailyAccountReader2_sout_yd.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PDailyAccount_sout;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='['+fsheet+'$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TSAPDailyAccountReader2_sout_wt.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if ADOTabXLS.FieldByName('单据编号').AsString = '' then
      begin
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('工单号').AsString;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('代工厂').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      sdt := ADOTabXLS.FieldByName('日期').AsString; 
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('成品料号').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('成品名称').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('工单数量').AsFloat;
      sdt := ADOTabXLS.FieldByName('领料日期').AsString;
      aSAPOPOAllocPtr^.dtout := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_child := ADOTabXLS.FieldByName('子项料号').AsString;
      aSAPOPOAllocPtr^.sname_child := ADOTabXLS.FieldByName('子项名称').AsString;
      aSAPOPOAllocPtr^.dqtyout := ADOTabXLS.FieldByName('领料数量').AsFloat;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('发料仓库').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('单位用量').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('备注（替代群组）').AsString;
      aSAPOPOAllocPtr^.sicmotype := ADOTabXLS.FieldByName('工单类型').AsString;

      aSAPOPOAllocPtr^.sstock := FStockMZ2FacReader.Fac2MZ_no(aSAPOPOAllocPtr^.sstock_yd);

 
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

 
{ TZPP_PRDORD_004Reader }


constructor TZPP_PRDORD_004Reader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;
      
     
destructor TZPP_PRDORD_004Reader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TZPP_PRDORD_004Reader.Clear;
var
  i: Integer;
  p: PZPP_PRDORD_004Record;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PZPP_PRDORD_004Record(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
procedure TZPP_PRDORD_004Reader.Log(const str: string);
begin

end;

function TZPP_PRDORD_004Reader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TZPP_PRDORD_004Reader.GetItems(i: Integer): PZPP_PRDORD_004Record;
begin
  Result := PZPP_PRDORD_004Record(FList[i]);
end;

function TZPP_PRDORD_004Reader.ICMOBillno2fac(const smz: string): string;
var
  i: Integer;
  p: PZPP_PRDORD_004Record;
begin
  Result := '';
  for i := 0 to FList.Count - 1 do
  begin
    p := PZPP_PRDORD_004Record(FList[i]);
    if p^.sicmobillno = smz then
    begin
      Result := p^.sicmobillno_fac;
      Break;
    end;
  end;
end;

procedure TZPP_PRDORD_004Reader.Open;
var  
  irow: Integer;
  snumber: string;
  ptrZPP_PRDORD_004Record: PZPP_PRDORD_004Record;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable; 
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='[Sheet1$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TZPP_PRDORD_004Reader.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      ptrZPP_PRDORD_004Record := New(PZPP_PRDORD_004Record);
      FList.Add(ptrZPP_PRDORD_004Record);

      ptrZPP_PRDORD_004Record^.sicmobillno_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;
      ptrZPP_PRDORD_004Record^.sicmobillno := ADOTabXLS.FieldByName('订单').AsString;

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

 
{ TCPINmz2facReader }


constructor TCPINmz2facReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;
      
     
destructor TCPINmz2facReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TCPINmz2facReader.Clear;
var
  i: Integer;
  p: PZPP_PRDORD_004Record;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PZPP_PRDORD_004Record(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
procedure TCPINmz2facReader.Log(const str: string);
begin

end;

function TCPINmz2facReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TCPINmz2facReader.GetItems(i: Integer): PCPINmz2facRecord;
begin
  Result := PCPINmz2facRecord(FList[i]);
end;

function TCPINmz2facReader.cpin_mz2fac(const smz: string): string;
var
  i: Integer;
  p: PCPINmz2facRecord;
  s: string;
begin
  Result := '';
  for i := 0 to FList.Count - 1 do
  begin
    p := PCPINmz2facRecord(FList[i]);

    if (p^.sMvT <> '101') and (p^.sMvT <> '102') then Continue;

    if p^.scpinbillno = smz then
    begin
      s := p^.scpinbillno_fac;

      if Copy(s, 1, 2) = 'ML' then
      begin
        s := Copy(s, 3, Length(s) - 2);
      end;

      if Copy(s, 1, 3) = 'NWT' then
      begin
        s := Copy(s, 8, Length(s) - 7);
      end;   

      if Copy(s, 1, 2) = 'WT' then
      begin
        s := Copy(s, 7, Length(s) - 6);
      end;

      if Copy(s, 1, 2) = 'SY' then
      begin
        s := Copy(s, 3, Length(s) - 2);
      end;

      Result := s;
      Break;
    end;
  end;
end;

procedure TCPINmz2facReader.Open;
var  
  irow: Integer;
  snumber: string;
  ptrCPINmz2facRecord: PCPINmz2facRecord;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;

  s: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    try

      ADOTabXLS.TableName:='[Sheet1$]';

      ADOTabXLS.Active:=true;

    except
      on e: Exception do
      begin
        Log( 'TCPINmz2facReader.Open ' +e.Message);
        Exit;
      end;
    end;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin
      if Pos('合计', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;



      s := ADOTabXLS.FieldByName('代工厂入库单号').AsString;
      if s <> '' then
      begin             
        ptrCPINmz2facRecord := New(PCPINmz2facRecord);
        FList.Add(ptrCPINmz2facRecord);
        ptrCPINmz2facRecord^.scpinbillno_fac := s;
        ptrCPINmz2facRecord^.scpinbillno := ADOTabXLS.FieldByName('物料凭证').AsString;
        ptrCPINmz2facRecord^.sMvT := ADOTabXLS.FieldByName('MvT').AsString;
      end;

      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

end.
