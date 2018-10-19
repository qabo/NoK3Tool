unit CommVars;

interface

uses
  ADODB;

const
  CSTitle75_E = 'M75料况表（BOM20140819 )---电子料';
  CIStartRow75_E = 15;
  CINumberCol75_E = 1;
  CIQtyFoxCol75_E = 4;  
  CIQtyWWCol75_E = 8;
  CIQtyMLCol75_E = 10; 
  CIQtyUnCheckCol75_E = 11;  
  CIQtySMTCol75_E = 12;
  CIQtyOrderCol75_E = 20;

  CSTitle75_S = 'M75料况表（BOM20140830 )---结构件';
  CIStartRow75_S = 16;
  CINumberCol75_S = 1;
  CIQtyFoxCol75_S = 4;  
  CIQtyWWCol75_S = 8;
  CIQtyMLCol75_S = 10;    
  CIQtyUnCheckCol75_S = 11;
  CIQtySMTCol75_S = 12;
  CIQtyOrderCol75_S = 20;

  CSTitle76_E = 'M76料况表---电子料部分';
  CIStartRow76_E = 11;
  CINumberCol76_E = 3;
  CIQtyMLCol76_E = 12;
  CIQtyFoxCol76_E = 14;
  CIQtyWWCol76_E = 16;
  CIQtyOrderStockCol76_E = 22;

  CSTitle79_S = 'M79料况表（BOM20141014 )---结构件';
  CIStartRow79_S = 16;
  CINumberCol79_S = 1;
  CIQtyFoxCol79_S = 5;
  CIQtyWWCol79_S = 9;
  CIQtyMLCol79_S = 11;
  CIQtyUnCheckCol79_S = 12;
  CIQtySMTCol79_S = 13;
  CIQtyOrderCol79_S = 22;

  CSINI_DivideOPO = 'DivideOPO.ini';

  CSBoolean: array[Boolean] of string = ('N', 'Y');

//  CSM75 = 'M75';
//  CSM76 = 'M76';
//  CSM71 = 'M71';

type   
  TLogEvent = procedure(const s: string) of object;

  TInvRecord = packed record
    fnumber: string;   
    fqty_ml: double;            //魅力原材料仓 (贴片仓+塑胶仓+五金仓+包材仓+辅料仓)
    fqty_smt: double;           //SMT仓
    fqty_fox_lf: double;           //廊坊富士康
    fqty_fox_bj: double;           //北京富士康
    fqty_wwsmt: double;         //外发SMT仓加总
    fqty_yh: double;            //易耗品
    fqty_uncheck: double;       //已收未验
    fqty_order: double;         //未结订单
    fqty_orderstock: double;    //订单入库
    fqty_os_fox_lf: double;        //订单入库 廊坊富士康
    fqty_os_fox_bj: double;        //订单入库 北京富士康
    fqty_os_ml: double;         //订单入库 魅力   
    fqty_qin: double;           //其他入库
    fqty_zz_zz: Double;         //总装车间
    fqty_zz_smt: Double;        //SMT车间
    fqty_zz_nd: Double;         //奈电车间
    fqty_zz_fox: Double;        //富士康车间
    fqty_zz_wcl: Double;        //伟创力车间
    fqty_zz_jw: Double;         //景旺车间
  end;
  PInvRecord = ^TInvRecord;

var
  gConnStr: string; 
  gUserID: Integer;
  gUserName: string;
  gToday: TDateTime;

  g880: Boolean = False;
  gTele: Boolean = False; // 通信
    
implementation

end.
