program NoK3Tool2;

uses
  Forms,
  SysUtils,
  MainWin in 'MainWin.pas' {frmMain},
  HWCinSumWin in 'HWCinSumWin.pas' {frmHWCinSum},
  CommUtils in 'CommUtils.pas', 
  MrpPRMergeWin in 'MrpPRMergeWin.pas' {frmMrpPRMerge},
  SOPAchievementWin in 'SOPAchievementWin.pas' {frmSOPAchievement},
  MergePlansWin in 'MergePlansWin.pas' {frmMergePlans},
  MergePlansAnalysisWin in 'MergePlansAnalysisWin.pas' {frmMergePlansAnalysis},
  SOPVSActWin in 'SOPVSActWin.pas' {frmSOPVSAct},
  SOPReaderUnit in 'SOPReaderUnit.pas',
  SOPVSActReaderUnit in 'SOPVSActReaderUnit.pas',
  BSFormatWin in 'BSFormatWin.pas' {frmBSFormat},
  BSDemandReader in 'BSDemandReader.pas',
  FGDemandWin in 'FGDemandWin.pas' {frmFGDemand},
  LocalFGDemandWin in 'LocalFGDemandWin.pas' {frmLocalFGDemand},
  ProjYearWin in 'ProjYearWin.pas' {frmProjYear},
  SelectFGDemandWin in 'SelectFGDemandWin.pas' {frmSelectFGDemand},
  DailyPlanVsActReader in 'DailyPlanVsActReader.pas',
  MPACompareWin in 'MPACompareWin.pas' {frmMPACompare},
  DataIntAnalysisReader in 'DataIntAnalysisReader.pas',
  WaterfallWin2 in 'WaterfallWin2.pas' {frmWaterfall2},
  ManMrpWin in 'ManMrpWin.pas' {frmManMrp},
  DataIntMergeWin in 'DataIntMergeWin.pas' {frmDataIntMerge},
  FGDemandConfigWin in 'FGDemandConfigWin.pas' {frmFGDemandConfig},
  DBConfigWin in 'DBConfigWin.pas' {frmDBConfig},
  ExcelConsts in 'ExcelConsts.pas',
  FGDemandManageWin in 'FGDemandManageWin.pas' {frmFGDemandManage},
  SOPvsMPSWin in 'SOPvsMPSWin.pas' {frmSOPvsMPS},
  KeyICItemSupplyReader in '..\..\erp\dev\K3Tool\KeyICItemSupplyReader.pas',
  SBomReader in 'SBomReader.pas',
  SOPSimReader in 'SOPSimReader.pas',
  MrpMPSReader in 'MrpMPSReader.pas',
  SopSimSumWin in 'SopSimSumWin.pas' {frmSopSimSum},
  NewSKUReader in 'NewSKUReader.pas',
  DOSReader in 'DOSReader.pas',
  FGPriorityReader in 'FGPriorityReader.pas',
  DOSPlanReader in 'DOSPlanReader.pas',
  StockBalReader in 'StockBalReader.pas',
  SEOutReader in 'SEOutReader.pas',
  ScheActExceptionWin in 'ScheActExceptionWin.pas' {frmScheActException},
  ProjOfICItemWin in 'ProjOfICItemWin.pas' {frmProjOfICItem},
  SOP2PPOrderWin in 'SOP2PPOrderWin.pas' {frmSOP2PPOrder},
  HWPkgStuffWin in 'HWPkgStuffWin.pas' {frmHWPkgStuff},
  FGTableReader in '..\..\erp\dev\K3Tool\FGTableReader.pas',
  FGStockReader in '..\..\erp\dev\K3Tool\FGStockReader.pas',
  WaterfallWin in 'WaterfallWin.pas' {frmWaterfall},
  SOP2SAPWin in 'SOP2SAPWin.pas' {frmSOP2SAP};

{$R *.res}

begin
  DateSeparator := '-';
  Application.Initialize;
  TfrmDBConfig.Load;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
