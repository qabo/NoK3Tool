unit MakeFGReportCommon;

interface

type
  TFGNumberRecord = packed record
    snumber: string;
    sname: string;
    sBatchNo: string;
  end;
  PFGNumberRecord = ^TFGNumberRecord;

implementation

end.
