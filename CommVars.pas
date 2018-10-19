unit CommVars;

interface

uses
  ADODB;

const
  CSTitle75_E = 'M75�Ͽ���BOM20140819 )---������';
  CIStartRow75_E = 15;
  CINumberCol75_E = 1;
  CIQtyFoxCol75_E = 4;  
  CIQtyWWCol75_E = 8;
  CIQtyMLCol75_E = 10; 
  CIQtyUnCheckCol75_E = 11;  
  CIQtySMTCol75_E = 12;
  CIQtyOrderCol75_E = 20;

  CSTitle75_S = 'M75�Ͽ���BOM20140830 )---�ṹ��';
  CIStartRow75_S = 16;
  CINumberCol75_S = 1;
  CIQtyFoxCol75_S = 4;  
  CIQtyWWCol75_S = 8;
  CIQtyMLCol75_S = 10;    
  CIQtyUnCheckCol75_S = 11;
  CIQtySMTCol75_S = 12;
  CIQtyOrderCol75_S = 20;

  CSTitle76_E = 'M76�Ͽ���---�����ϲ���';
  CIStartRow76_E = 11;
  CINumberCol76_E = 3;
  CIQtyMLCol76_E = 12;
  CIQtyFoxCol76_E = 14;
  CIQtyWWCol76_E = 16;
  CIQtyOrderStockCol76_E = 22;

  CSTitle79_S = 'M79�Ͽ���BOM20141014 )---�ṹ��';
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
    fqty_ml: double;            //����ԭ���ϲ� (��Ƭ��+�ܽ���+����+���Ĳ�+���ϲ�)
    fqty_smt: double;           //SMT��
    fqty_fox_lf: double;           //�ȷ���ʿ��
    fqty_fox_bj: double;           //������ʿ��
    fqty_wwsmt: double;         //�ⷢSMT�ּ���
    fqty_yh: double;            //�׺�Ʒ
    fqty_uncheck: double;       //����δ��
    fqty_order: double;         //δ�ᶩ��
    fqty_orderstock: double;    //�������
    fqty_os_fox_lf: double;        //������� �ȷ���ʿ��
    fqty_os_fox_bj: double;        //������� ������ʿ��
    fqty_os_ml: double;         //������� ����   
    fqty_qin: double;           //�������
    fqty_zz_zz: Double;         //��װ����
    fqty_zz_smt: Double;        //SMT����
    fqty_zz_nd: Double;         //�ε糵��
    fqty_zz_fox: Double;        //��ʿ������
    fqty_zz_wcl: Double;        //ΰ��������
    fqty_zz_jw: Double;         //��������
  end;
  PInvRecord = ^TInvRecord;

var
  gConnStr: string; 
  gUserID: Integer;
  gUserName: string;
  gToday: TDateTime;

  g880: Boolean = False;
  gTele: Boolean = False; // ͨ��
    
implementation

end.
