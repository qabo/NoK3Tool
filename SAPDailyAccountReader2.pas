unit SAPDailyAccountReader2;
{
��̩���⣺
  1�����ϳ���
  2����������
  3������
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

  TSAPDailyAccountReader2_DB = class // �⹺��� ����
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

  TDailyAccount_DB_in = packed record // ����
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

  TSAPDailyAccountReader2_DB_in = class // �⹺��� ����
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





  
  TDailyAccount_coois = packed record // Ͷ�ϵ�
    sbillno_fac: string; //����������
    sbillno: string; //	����
    scategory: string; //	����
    dtfac: TDateTime; //	��������������
    sbiller: string; //	�Ƶ���
    snumber: string; //	����
    dtFinish: TDateTime; //	�ƻ��깤
    sbillno_plan: string; //	�ƻ�����
    dqtyorder: Double; //	��������
    sBUn: string; //	BUn
    sstockname: string; //��λ
    snumber_item: string; //	����
    dtneed: TDateTime; //	��������
    dqtyneed: Double; //	������
    sunit: string; //��
    sfac: string; //����
    sFix: string; //	Fix
    dtChangeDate: TDateTime; //	�������
    dtChangeTime: TDateTime; //	���ʱ��
    dQtyIn: Double; //	�ջ�����

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


  TDailyAccount_FacICMO2MZICMO = packed record // Ͷ�ϵ�
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


  TDailyAccount_icmo_mz2fac = packed record // Ͷ�ϵ�
    sicmolbillno_fac: string; //����������
    sicmobillno: string; //	����
    stype: string; //	����
    sfacno: string; //	����������
    dtdate_fac: TDateTime; //	��������������
    sbiller: string; //	�Ƶ���
    ssourcebillno: string; //	��Դ������
    swwcontract1: string; //	ί���ͬ1
    dqty_contract_alloc1: Double; // ��ͬ��������1
//    sEUn: string; //EUn
    swwcontract2: string; //	ί���ͬ2
    dqty_contract_alloc2: Double; // ��ͬ��������2
    //EUn
    //EUn
    snote: string; //��ע
    sall_transfer_flag: string; //	��ȫת����־
    dtChangeDate: TDateTime; //	�������
    dtChangeTime: TDateTime; //	���ʱ��
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

  TDailyAccount_PPBom = packed record // Ͷ�ϵ�
   dtdate: TDateTime; //�Ƶ�����
   dtCheck: TDateTime; //	�������
   sicmobillno: string; //	����/ί�ⶩ����
   snumber: string; //	��Ʒ����
   sname: string; //	��Ʒ����
   dqty: Double; //	��������
   snote: string; //	��ע
   sppbombillno: string; //	����Ͷ�ϵ���
   snumber_item: string; //	�������ϳ�����
   sname_item: string; // ������������
   dqtyplan: Double; //	�ƻ�Ͷ������
   dqtyshould: Double; //	Ӧ������
   sstockname: string; //	�ֿ�
   sstockname_ml: string; //	�ֿ�
   dusage: Double; //	��λ����
   scheckflag: string; //	��˱�־
   sworkshopname: string; //	��������
   sedi: string; //	EDI�ύ

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
    schangebillno: string; //�������
    sarea: string; //MRP ��Χ
    sbillno: string; //���ݱ��
    sorderbillno: string; //����
    snumber: string; //����
    sunit: string; //��
    splanroder: string; //�ƻ�����
    sqtychangeflag: string; //���������־
    sreason: string; //���ԭ��
    schangetime: string; //���ʱ��
    snumber_item: string; //���
    sZTIPP007B_ITEM: string; //ZTIPP007B-ITEM
    sZTIPP007B_LGORT: string; //ZTIPP007B-LGORT
    sZTIPP007B_ALPGR: string; //ZTIPP007B-ALPGR
    sZTIPP007B_ALPRF: string; //ZTIPP007B-ALPRF
    sZTIPP007B_EWAHR: string; //ZTIPP007B-EWAHR
    sZTIPP007B_ITEM_FLAG: string; //ZTIPP007B-ITEM_FLAG
    sZTIPP007B_REMARK: string; //ZTIPP007B-REMARK
    sZTIPP007B_UPDKZ: string; //ZTIPP007B-UPDKZ
    sicmo_fac: string; //����������
    dqty: Double;//����
    //��
    dqtyBefore: Double; //�޸�ǰ����
    //��
    dtChange: TDateTime;//�������
    sZTIPP007B_MENGE: string; //ZTIPP007B-MENGE
    dtNeed: TDateTime; //��������
    sZTIPP007B_MENGE_B: string; //ZTIPP007B-MENGE_B
    sZTIPP007B_MENGE_T: string; //ZTIPP007B-MENGE_T
    //�������
    bCalc:Boolean;
    sMatchType: string;
  end;
  PDailyAccount_PPBomChange_mz = ^TDailyAccount_PPBomChange_mz;
       
  TDailyAccount_PPBomChange_yd = packed record
    sChangeFlag: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�����־').AsString;
    snumber: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('��Ʒ����').AsString;
    sname: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('��Ʒ����').AsString;
    sppbombillno: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('����Ͷ�ϵ���').AsString;
    snumber_item: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('���ϴ���').AsString;
    sname_item: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('��������').AsString;
    susage: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('��׼����').AsString;
    sstock_fac: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�ֿ�').AsString;
    sChangeReason: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('���ԭ��').AsString;
    sdt: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�Ƶ�����').AsString;
    sdtCheck: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�������').AsString;
    sChangeVer: string; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('������').AsString;
    dQty: Double; //  aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�ƻ�Ͷ������').AsString;
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

  TDailyAccount_DB_out = packed record // ����
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

  TDailyAccount_OutAInBC = packed record // ����
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
    sxoutbillno: string; //��������
    sxoutdept: string; //������λ
    snumber: string; //�Ϻ�
    sname: string; //��Ʒ����
    dqty: Double; //����
    sorder: string; //��������
    sproxy: string; //�����̼��
    sexp: string; //��ݹ�˾
    sebillno: string; //���ӵ���
    smnote: string; //������ע
    sddate: string; //����ʱ��
    sstock_fac: string; //��λ
    sdate: string; //����
    snote: string; //��ע

    sstock_mz: string; //��λ
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
    sicmobillno_fac: string; // ����������
    sicmobillno: string; //����
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
    scpinbillno_fac: string; // ����������
    scpinbillno: string; //����
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
           
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if ADOTabXLS.FieldByName('���ϳ�����').AsString = '' then
      begin
        ADOTabXLS.Next;
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;

      if ADOTabXLS.FindField('ʵ������') <> nil then
      begin
        aSAPOPOAllocPtr^.dQty  := ADOTabXLS.FieldByName('ʵ������').AsFloat;
      end
      else
      begin
        aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('����').AsFloat;
      end;

      sdt := ADOTabXLS.FieldByName('����').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('��;').AsString;
      if ADOTabXLS.FindField('������λ') <> nil then
      begin
        aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('������λ').AsString;
      end
      else if ADOTabXLS.FindField('��Ӧ��') <> nil then
      begin
        aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('��Ӧ��').AsString;
      end;

      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;

      if ADOTabXLS.FindField('���ϲֿ�') <> nil then
      begin
        aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      end
      else if ADOTabXLS.FindField('���ϲֿ�') <> nil then
      begin
        aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      end;

      if ADOTabXLS.FindField('ժҪ') <> nil then
      begin
        aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('ժҪ').AsString;
      end;
      
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('�رձ�־') >= 0 then
        aSAPOPOAllocPtr^.sclose := ADOTabXLS.FieldByName('�رձ�־').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('����') >= 0 then
        aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('����').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('���鷽ʽ') >= 0 then
        aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('���鷽ʽ').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
      if ADOTabXLS.FieldDefList.IndexOf('Դ������') >= 0 then
        aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('Դ������').AsString;

      if aSAPOPOAllocPtr^.sstock_ml = 'ԭ���ϴ����' then
      begin
        if aSAPOPOAllocPtr^.suse = '����' then
        begin
          aSAPOPOAllocPtr^.sstock := '����ԭ���ϲ�';
        end
        else if aSAPOPOAllocPtr^.suse = '�ƹ�' then
        begin
          aSAPOPOAllocPtr^.sstock := '�����ƹ��';
        end
        else if aSAPOPOAllocPtr^.suse = '�Բ�' then
        begin
          aSAPOPOAllocPtr^.sstock := '����RM�Բ�ԭ�ϲ�';
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
           
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString <> '�͹������˻�') and
        (ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString <> '�͹����') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;


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
           
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;


      if ADOTabXLS.FieldByName('���ݱ��').AsString = '' then
      begin      
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_winB);
      aSAPOPOAllocPtr^.bcalc := False;
      FList.Add(aSAPOPOAllocPtr);

      
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('ʵ������').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      if sdt <> '' then
        aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt)
      else aSAPOPOAllocPtr^.dtCheck := 0;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('��Ӧ��').AsString;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('ժҪ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;

 
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
           
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_RTV);
      FList.Add(aSAPOPOAllocPtr);
                                                                                
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('������λ').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('��;').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('Դ������').AsString;   
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;                                                                                 
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;


      if aSAPOPOAllocPtr^.sstock_ml = 'ԭ���ϴ����' then
      begin
        if aSAPOPOAllocPtr^.suse = '����' then
        begin
          aSAPOPOAllocPtr^.sstock := '����ԭ���ϲ�';
        end
        else if aSAPOPOAllocPtr^.suse = '�ƹ�' then
        begin
          aSAPOPOAllocPtr^.sstock := '�����ƹ��';
        end
        else if aSAPOPOAllocPtr^.suse = '�Բ�' then
        begin
          aSAPOPOAllocPtr^.sstock := '����RM�Բ�ԭ�ϲ�';
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('�������񵥺�').AsString;

      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('ʵ������').AsFloat;
      aSAPOPOAllocPtr^.sbatchno := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('�ջ��ֿ�').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('������λ').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.schecker := ADOTabXLS.FieldByName('�����').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sbackflush := ADOTabXLS.FieldByName('�����־').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      if sdt = '' then
      begin
        aSAPOPOAllocPtr^.dtbill := 0
      end
      else
      begin
        aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      end;
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;
 
 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_cpin);
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ�Ϻ�').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('�������').AsFloat;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('�ջ��ֿ�').AsString;


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


      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').asstring;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;
      //SAP����
      //����
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('��;').AsString;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('��Ӧ��').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('ժҪ').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('�رձ�־').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('���鷽ʽ').AsString;
      aSAPOPOAllocPtr^.sedit := ADOTabXLS.FieldByName('EDI�ύ').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('Դ������').AsString; 
 
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
      ADOTabXLS.TableName:='[�������$]';

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

      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString <> '�͹���Ʒ') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;


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

      if ADOTabXLS.FieldByName('���ݱ��').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qin);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('��Ӧ��').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('ժҪ').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('�رձ�־').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('���鷽ʽ').AsString;

 
 
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


      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ������').asstring;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := - ADOTabXLS.FieldByName('����').AsFloat;


      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sdetp := ADOTabXLS.FieldByName('���ϲ���').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('��;1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('��;2').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
 
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



      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;
 
 
 
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

      if ADOTabXLS.FieldByName('��Ʒ������').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_qout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ������').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtcheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sdetp := ADOTabXLS.FieldByName('���ϲ���').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('��;1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('��;2').AsString;
 
 
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

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sstock_in_ml := ADOTabXLS.FieldByName('����ֿ�').AsString;       
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;

      sdt := ADOTabXLS.FieldByName('����').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);

//      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;


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

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      if UpperCase(Copy(aSAPOPOAllocPtr^.snumber, 1, 3)) = 'KMZ' then
      begin
        aSAPOPOAllocPtr^.snumber := Copy(aSAPOPOAllocPtr^.snumber, 4, Length(aSAPOPOAllocPtr^.snumber) - 3);
      end;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;
                                                                                                
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


      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sbillno := StringReplace(aSAPOPOAllocPtr^.sbillno, '?', '', [rfReplaceAll]);
      aSAPOPOAllocPtr^.sstock_out_yd := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sstock_in_yd := ADOTabXLS.FieldByName('����ֿ�').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;

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

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('��;').AsString;
      aSAPOPOAllocPtr^.ssupplier := ADOTabXLS.FieldByName('��Ӧ��').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sstock_in_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.ssummary := ADOTabXLS.FieldByName('ժҪ').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.scloseflag := ADOTabXLS.FieldByName('�رձ�־').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.schecktype := ADOTabXLS.FieldByName('���鷽ʽ').AsString;  
      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('Դ������').AsString;
       
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
      ADOTabXLS.TableName:='[�������$]';

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
    
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if (ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString <> '�͹����������') then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfacname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdoc := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.smpn := ADOTabXLS.FieldByName('�����̴���').AsString;
      aSAPOPOAllocPtr^.smpn_name := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvt := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      aSAPOPOAllocPtr^.smvr := ADOTabXLS.FieldByName('�ƶ�ԭ��').AsString;
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.smodel := ADOTabXLS.FieldByName('����ͺ�').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sunit := '';  //ADOTabXLS.FieldByName('����������λ').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;
      aSAPOPOAllocPtr^.swc := ''; //ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.sitemtext := ADOTabXLS.FieldByName('��Ŀ�ı�').AsString;
      aSAPOPOAllocPtr^.sitemno := ADOTabXLS.FieldByName('������Ŀ��').AsString;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ص�').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sitemgroupname := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.smvr_desc := ADOTabXLS.FieldByName('�ƶ�ԭ������').AsString;
      aSAPOPOAllocPtr^.sitemgroup := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sordertype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('������������').AsFloat;
      aSAPOPOAllocPtr^.sdoc_item := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.smvt_desc := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.sstatus := ADOTabXLS.FieldByName('�춯״��').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtbill := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.dbillqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sstock_desc_wt := ADOTabXLS.FieldByName('�ִ��ص������').AsString;


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

  icolDate: Integer;//  sdt := ADOTabXLS.FieldByName('����').AsString; 
  icolBillno: Integer; //  aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
  icolStockOut: Integer; //  aSAPOPOAllocPtr^.sstockno_out_yd := ADOTabXLS.FieldByName('�����ֿ�').AsString;
  icolStockIn: Integer; //  aSAPOPOAllocPtr^.sstockno_in_yd := ADOTabXLS.FieldByName('����ֿ�').AsString;
  icolNumber: Integer; //  aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
  icolName: Integer; //  aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
  icolQty: Integer;  //  aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;
  sSheet: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;




  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
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
        icolDate := IndexOfCol(ExcelApp, irow, '����');
        icolBillno := IndexOfCol(ExcelApp, irow, '���ݱ��');
        icolStockOut := IndexOfCol(ExcelApp, irow, '�����ֿ�');
        icolStockIn := IndexOfCol(ExcelApp, irow, '����ֿ�');
        icolNumber := IndexOfCol(ExcelApp, irow, '���ϳ�����');
        icolName := IndexOfCol(ExcelApp, irow, '��������');
        icolQty := IndexOfCol(ExcelApp, irow, '��������');        

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
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
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
    
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_in);
      FList.Add(aSAPOPOAllocPtr);
                                                                         
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sstockno_out_yd := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sstockno_in_yd := ADOTabXLS.FieldByName('����ֿ�').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('��������').AsFloat;

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


      aSAPOPOAllocPtr^.sbillno_fac := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.scategory := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dtfac := ADOTabXLS.FieldByName('��������������').AsDateTime;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ���').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dtFinish := ADOTabXLS.FieldByName('�ƻ��깤').AsDateTime;
      aSAPOPOAllocPtr^.sbillno_plan := ''; //ADOTabXLS.FieldByName('�ƻ�����').AsString;
      aSAPOPOAllocPtr^.dqtyorder := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sBUn := ADOTabXLS.FieldByName('BUn').AsString;
      aSAPOPOAllocPtr^.sstockname := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('����1').AsString;
      aSAPOPOAllocPtr^.dtneed := ADOTabXLS.FieldByName('��������').AsDateTime;
      aSAPOPOAllocPtr^.dqtyneed := ADOTabXLS.FieldByName('������').AsFloat;
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��').AsString;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sFix := ADOTabXLS.FieldByName('Fix').AsString;
      aSAPOPOAllocPtr^.dtChangeDate := ADOTabXLS.FieldByName('�������').AsDateTime;
      aSAPOPOAllocPtr^.dtChangeTime := ADOTabXLS.FieldByName('���ʱ��').AsDateTime;
      aSAPOPOAllocPtr^.dQtyIn := ADOTabXLS.FieldByName('�ջ�����').AsFloat;
           
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

      aSAPOPOAllocPtr^.sicmo_fac := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.stype := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.slang := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.swwpo := ADOTabXLS.FieldByName('ί��ɹ�����').AsString;
      aSAPOPOAllocPtr^.ssourceorder := ADOTabXLS.FieldByName('��Դ������').AsString;
      sdt := ADOTabXLS.FieldByName('��������������').AsString;
      aSAPOPOAllocPtr^.dtFac := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�����������').AsString;;
      aSAPOPOAllocPtr^.dtend := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('������ʼ����').AsString;;
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

      aSAPOPOAllocPtr^.sicmolbillno_fac := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.stype := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sfacno := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.dtdate_fac := ADOTabXLS.FieldByName('��������������').AsDateTime;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ���').AsString;
      aSAPOPOAllocPtr^.ssourcebillno := ADOTabXLS.FieldByName('��Դ������').AsString;
      aSAPOPOAllocPtr^.swwcontract1 := ADOTabXLS.FieldByName('ί���ͬ1').AsString;
      aSAPOPOAllocPtr^.dqty_contract_alloc1 := ADOTabXLS.FieldByName('��ͬ��������1').AsFloat;
      aSAPOPOAllocPtr^.swwcontract2 := ADOTabXLS.FieldByName('ί���ͬ2').AsString;
      aSAPOPOAllocPtr^.dqty_contract_alloc1 := ADOTabXLS.FieldByName('��ͬ��������2').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sall_transfer_flag := ADOTabXLS.FieldByName('��ȫת����־').AsString;
      aSAPOPOAllocPtr^.dtChangeDate := ADOTabXLS.FieldByName('�������').AsDateTime;
      aSAPOPOAllocPtr^.dtChangeTime := ADOTabXLS.FieldByName('���ʱ��').AsDateTime; 
            
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_PPBom);
      FList.Add(aSAPOPOAllocPtr);

      sdt := ADOTabXLS.FieldByName('�Ƶ�����').AsString;

      aSAPOPOAllocPtr^.dtdate := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('����/ί�ⶩ����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('����Ͷ�ϵ���').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('�������ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.dqtyplan := ADOTabXLS.FieldByName('�ƻ�Ͷ������').AsFloat;
      aSAPOPOAllocPtr^.dqtyshould := ADOTabXLS.FieldByName('Ӧ������').AsFloat;
      aSAPOPOAllocPtr^.sstockname := ADOTabXLS.FieldByName('�ֿ�').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('��λ����').AsFloat;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sworkshopname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sedi := ''; //ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
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
      if ADOTabXLS.FieldByName('��������').AsString = '' then
      begin            
        ADOTabXLS.Next;
        Continue;
      end;
      
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;


      aSAPOPOAlloc.sfacname := ADOTabXLS.FieldByName('�ӹ�������').AsString;
      aSAPOPOAlloc.sfac := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAlloc.sicmobillno := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAlloc.sicmotye := ADOTabXLS.FieldByName('��������').AsString;
      sdt := ADOTabXLS.FieldByName('�´�����').AsString;
      aSAPOPOAlloc.dtRelease := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�᰸����').AsString;
      if sdt = '' then
      begin
        aSAPOPOAlloc.dtClose := 0;
      end
      else
      begin
        aSAPOPOAlloc.dtClose := myStrToDateTime(sdt);
      end;
      sdt := ADOTabXLS.FieldByName('������ʼ����').AsString;
      aSAPOPOAlloc.dtBegin := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�����������').AsString;
      aSAPOPOAlloc.dtEnd := myStrToDateTime(sdt);
      aSAPOPOAlloc.splanbillno := ''; //ADOTabXLS.FieldByName('�ƻ�����').AsString;
      aSAPOPOAlloc.splanbillno_mz := ADOTabXLS.FieldByName('����ƻ�����').AsString;
      aSAPOPOAlloc.snumber_wt := ADOTabXLS.FieldByName('��̩���Ϻ�').AsString;
      aSAPOPOAlloc.snumber := ADOTabXLS.FieldByName('�ͻ����Ϻ�').AsString;
      aSAPOPOAlloc.svItemFlag := ADOTabXLS.FieldByName('������Ŀ��ʶ').AsString;
      aSAPOPOAlloc.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAlloc.sItemCode := ADOTabXLS.FieldByName('��Ŀ����').AsString;
      aSAPOPOAlloc.dICMOQty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAlloc.snote1 := ADOTabXLS.FieldByName('��ע1').AsString;
      aSAPOPOAlloc.iChangeCount := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAlloc.irowitem := ADOTabXLS.FieldByName('����Ŀ').AsString;
      aSAPOPOAlloc.snumber_item_wt := ADOTabXLS.FieldByName('��̩�����ϱ���').AsString;
      aSAPOPOAlloc.snumber_item := ADOTabXLS.FieldByName('�ͻ������ϱ���').AsString;
      aSAPOPOAlloc.sname_item := ADOTabXLS.FieldByName('��������1').AsString;
      aSAPOPOAlloc.dqtyplan := ADOTabXLS.FieldByName('������').AsFloat;
      aSAPOPOAlloc.dqtyout := ADOTabXLS.FieldByName('��Ͷ������').AsFloat;
      aSAPOPOAlloc.sstockname_wt := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAlloc.dqty0 := ADOTabXLS.FieldByName('���ǰ����').AsFloat;
      aSAPOPOAlloc.sgroup := ADOTabXLS.FieldByName('�����').AsString;
      aSAPOPOAlloc.sprioriry := ADOTabXLS.FieldByName('���ȼ�').AsString;
      aSAPOPOAlloc.dper := ADOTabXLS.FieldByName('�������').AsFloat;
      aSAPOPOAlloc.dqtyshould := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAlloc.sunit := ADOTabXLS.FieldByName('������λ').AsString;
      aSAPOPOAlloc.snote2 := ADOTabXLS.FieldByName('��ע2').AsString;
      aSAPOPOAlloc.schangelog := ADOTabXLS.FieldByName('������').AsString;

 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;                   
      sdt := ADOTabXLS.FieldByName('�Ƶ�����').AsString;
      if sdt = '' then Break;

      aSAPOPOAllocPtr := New(PDailyAccount_PPBom);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.dtdate := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.sicmobillno := ADOTabXLS.FieldByName('����/ί�ⶩ����').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('����Ͷ�ϵ���').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('�������ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('������������').AsString;
      aSAPOPOAllocPtr^.dqtyplan := ADOTabXLS.FieldByName('�ƻ�Ͷ������').AsFloat;
      aSAPOPOAllocPtr^.dqtyshould := ADOTabXLS.FieldByName('Ӧ������').AsFloat;
      aSAPOPOAllocPtr^.sstockname_yd := ADOTabXLS.FieldByName('�ֿ�').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('��λ����').AsFloat;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sworkshopname := ADOTabXLS.FieldByName('��������').AsString;      


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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if Trim(ADOTabXLS.FieldByName('�������').AsString) = '' then
      begin            
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_PPBomChange_mz);
      aSAPOPOAllocPtr^.bCalc := False;
      FList.Add(aSAPOPOAllocPtr);


      aSAPOPOAllocPtr^.schangebillno := ADOTabXLS.FieldByName('�������').AsString;//  : string; //
      aSAPOPOAllocPtr^.sarea := ADOTabXLS.FieldByName('MRP ��Χ').AsString;//  : string; //
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;//  : string; //
      aSAPOPOAllocPtr^.sorderbillno := ADOTabXLS.FieldByName('����').AsString;//  : string; //
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('����').AsString;//  : string; //
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��').AsString;//  : string; //
      aSAPOPOAllocPtr^.splanroder := ''; //ADOTabXLS.FieldByName('�ƻ�����').AsString;//  : string; //
      aSAPOPOAllocPtr^.sqtychangeflag := ADOTabXLS.FieldByName('���������־').AsString;//  : string; //
      aSAPOPOAllocPtr^.sreason := ADOTabXLS.FieldByName('���ԭ��').AsString;//  : string; //
      aSAPOPOAllocPtr^.schangetime := ADOTabXLS.FieldByName('���ʱ��').AsString;//  : string; //
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('���').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ITEM := ADOTabXLS.FieldByName('ZTIPP007B-ITEM').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_LGORT := ADOTabXLS.FieldByName('ZTIPP007B-LGORT').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ALPGR := ADOTabXLS.FieldByName('ZTIPP007B-ALPGR').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ALPRF := ADOTabXLS.FieldByName('ZTIPP007B-ALPRF').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_EWAHR := ADOTabXLS.FieldByName('ZTIPP007B-EWAHR').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_ITEM_FLAG := ADOTabXLS.FieldByName('ZTIPP007B-ITEM_FLAG').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_REMARK := ADOTabXLS.FieldByName('ZTIPP007B-REMARK').AsString;//  : string; //
      aSAPOPOAllocPtr^.sZTIPP007B_UPDKZ := ADOTabXLS.FieldByName('ZTIPP007B-UPDKZ').AsString;//  : string; //
      aSAPOPOAllocPtr^.sicmo_fac := ADOTabXLS.FieldByName('����������').AsString;//  : string; //
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;//  : Double;//
      aSAPOPOAllocPtr^.dqtyBefore := ADOTabXLS.FieldByName('�޸�ǰ����').AsFloat;//  : string; //
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtChange := myStrToDateTime(sdt);//  : TDateTime;//
      aSAPOPOAllocPtr^.sZTIPP007B_MENGE := ADOTabXLS.FieldByName('ZTIPP007B-MENGE').AsString;//  : string; //
      sdt := ADOTabXLS.FieldByName('��������').AsString;
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

      ADOTabXLS.TableName:='[Ͷ�ϱ����$]';

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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      if ADOTabXLS.FieldByName('��Ʒ����').AsString = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
      aSAPOPOAllocPtr := New(PDailyAccount_PPBomChange_yd);
      FList.Add(aSAPOPOAllocPtr);
 
      aSAPOPOAllocPtr^.sChangeFlag := ADOTabXLS.FieldByName('�����־').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.sppbombillno := ADOTabXLS.FieldByName('����Ͷ�ϵ���').AsString;
      aSAPOPOAllocPtr^.snumber_item := ADOTabXLS.FieldByName('���ϴ���').AsString;
      aSAPOPOAllocPtr^.sname_item := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.susage := ADOTabXLS.FieldByName('��׼����').AsString;
      aSAPOPOAllocPtr^.sstock_fac := ADOTabXLS.FieldByName('�ֿ�').AsString;
      aSAPOPOAllocPtr^.sChangeReason := ADOTabXLS.FieldByName('���ԭ��').AsString; 
      aSAPOPOAllocPtr^.sdt := ADOTabXLS.FieldByName('�Ƶ�����').AsString;
      aSAPOPOAllocPtr^.sdtCheck := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.sChangeVer := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.dQty := ADOTabXLS.FieldByName('�ƻ�Ͷ������').AsFloat;
                                     
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
      if ADOTabXLS.FieldByName('���ݱ��').AsString = '' then
      begin
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_DB_out);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ������').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;        
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);                      
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('���ϲ���').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('��;1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;       
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('��;2').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
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

      if ADOTabXLS.FieldByName('���ݱ��').AsString = '' then
      begin    
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_OutAInBC);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ������').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;        
      aSAPOPOAllocPtr^.dQty := - ADOTabXLS.FieldByName('����').AsFloat;
      sdt := ADOTabXLS.FieldByName('����').AsString;   
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);                      
      aSAPOPOAllocPtr^.sstock_out_ml := ADOTabXLS.FieldByName('�����ֿ�').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('���ϲ���').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.suse1 := ADOTabXLS.FieldByName('��;1').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;       
      aSAPOPOAllocPtr^.sunit := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.souttype := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.suse2 := ADOTabXLS.FieldByName('��;2').AsString;

      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_xout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sxoutbillno := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sxoutdept := ADOTabXLS.FieldByName('������λ').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('�Ϻ�').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;
      aSAPOPOAllocPtr^.sorder := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sproxy := ADOTabXLS.FieldByName('�����̼��').AsString;
      aSAPOPOAllocPtr^.sexp := ADOTabXLS.FieldByName('��ݹ�˾').AsString;
      aSAPOPOAllocPtr^.sebillno := ADOTabXLS.FieldByName('���ӵ���').AsString;
      aSAPOPOAllocPtr^.smnote := ADOTabXLS.FieldByName('������ע').AsString;
      aSAPOPOAllocPtr^.sddate := ADOTabXLS.FieldByName('����ʱ��').AsString;
      aSAPOPOAllocPtr^.sstock_fac := ADOTabXLS.FieldByName('��λ').AsString;
      aSAPOPOAllocPtr^.sdate := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;

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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);
                                                                             
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('�������񵥺�').AsString;

      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      sdt := ADOTabXLS.FieldByName('�������').AsString;
      aSAPOPOAllocPtr^.dtCheck := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.scostnumber := ADOTabXLS.FieldByName('�ɱ��������').AsString;
      aSAPOPOAllocPtr^.scostname := ADOTabXLS.FieldByName('�ɱ�����').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('���ϳ�����').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('ʵ������').AsFloat;
      aSAPOPOAllocPtr^.sstock_ml := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sdept := ADOTabXLS.FieldByName('���ϲ���').AsString;
      aSAPOPOAllocPtr^.suse := ADOTabXLS.FieldByName('������;').AsString;
      aSAPOPOAllocPtr^.sbatchno := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�����').AsString;
      aSAPOPOAllocPtr^.scheckflag := ADOTabXLS.FieldByName('��˱�־').AsString;
      aSAPOPOAllocPtr^.sbiller := ADOTabXLS.FieldByName('�Ƶ�').AsString;
      aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;


      if ADOTabXLS.FieldDefList.IndexOf('EDI�ύ') >= 0 then
        aSAPOPOAllocPtr^.sedi := ADOTabXLS.FieldByName('EDI�ύ').AsString;
 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('������').AsString;
      sdt := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_wt := ADOTabXLS.FieldByName('��Ʒ�Ϻ�').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('MZ').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dicmoqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.snote1 := ADOTabXLS.FieldByName('��ע1').AsString;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtout := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_child := ADOTabXLS.FieldByName('�����Ϻ�').AsString;
      aSAPOPOAllocPtr^.sname_child := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqtyout := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sstock_wt := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.sbomusage := ADOTabXLS.FieldByName('BOM����').AsString;
      aSAPOPOAllocPtr^.snote2 := ADOTabXLS.FieldByName('��ע2').AsString;
      aSAPOPOAllocPtr^.sicmotype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
 
 
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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      if ADOTabXLS.FieldByName('���ݱ��').AsString = '' then
      begin
        Break;
      end;

      aSAPOPOAllocPtr := New(PDailyAccount_sout);
      FList.Add(aSAPOPOAllocPtr);

      aSAPOPOAllocPtr^.sicmo := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('������').AsString;
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      sdt := ADOTabXLS.FieldByName('����').AsString; 
      aSAPOPOAllocPtr^.dt := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('��Ʒ�Ϻ�').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��Ʒ����').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('��������').AsFloat;
      sdt := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dtout := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.snumber_child := ADOTabXLS.FieldByName('�����Ϻ�').AsString;
      aSAPOPOAllocPtr^.sname_child := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dqtyout := ADOTabXLS.FieldByName('��������').AsFloat;
      aSAPOPOAllocPtr^.sstock_yd := ADOTabXLS.FieldByName('���ϲֿ�').AsString;
      aSAPOPOAllocPtr^.dusage := ADOTabXLS.FieldByName('��λ����').AsString;
      aSAPOPOAllocPtr^.snote := ADOTabXLS.FieldByName('��ע�����Ⱥ�飩').AsString;
      aSAPOPOAllocPtr^.sicmotype := ADOTabXLS.FieldByName('��������').AsString;

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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      ptrZPP_PRDORD_004Record := New(PZPP_PRDORD_004Record);
      FList.Add(ptrZPP_PRDORD_004Record);

      ptrZPP_PRDORD_004Record^.sicmobillno_fac := ADOTabXLS.FieldByName('����������').AsString;
      ptrZPP_PRDORD_004Record^.sicmobillno := ADOTabXLS.FieldByName('����').AsString;

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
      if Pos('�ϼ�', ADOTabXLS.Fields[0].AsString) > 0 then
      begin
        ADOTabXLS.Next;
        Continue;
      end;



      s := ADOTabXLS.FieldByName('��������ⵥ��').AsString;
      if s <> '' then
      begin             
        ptrCPINmz2facRecord := New(PCPINmz2facRecord);
        FList.Add(ptrCPINmz2facRecord);
        ptrCPINmz2facRecord^.scpinbillno_fac := s;
        ptrCPINmz2facRecord^.scpinbillno := ADOTabXLS.FieldByName('����ƾ֤').AsString;
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
