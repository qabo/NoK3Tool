{*******************************************************}
{                                                       }
{                       EhLib v5.6                      }
{                     EhLib v6.0.01                     }
{                   Resource of strings                 }
{                                                       }
{      Copyright (c) 2001-2012 by Dmitry V. Bolshakov   }
{                                                       }
{*******************************************************}

unit EhLibConsts;

interface

resourcestring
  SClearSelectedCellsEh = '���õ� ���� ����ðڽ��ϱ�?';
  SInvalidTextFormatEh = '�߸��� ���ڿ� �����Դϴ�!';
  SInvalidVCLDBIFFormatEh = '�߸��� VCLDBIF �����Դϴ�!';
  SErrorDuringInsertValueEh = '�ڷ� �߰� ���� ���� �߻�: ';
  SIgnoreErrorEh = '������ �����Ͻðڽ��ϱ�?';
  STabularInformationEh = 'ǥ ����';
  SPageOfPagesEh = '%d / %d ������';
  SPreviewEh = '�̸�����';
  SFieldNameNotFoundEh = '�ʵ�� "%s"��(��) ã�� �� �����ϴ�!';
  SFindDialogStringNotFoundMessageEh = '���ڿ� "%s"��(��) ã�� �� �����ϴ�!'; 
  SVisibleColumnsEh = 'ǥ�� �׸�';
  SCutEh = '�߶󳻱�';
  SCopyEh = '����';
  SPasteEh = '�ٿ��ֱ�';
  SDeleteEh = '����';
  SSelectAllEh = '��� ����';

  SSTFilterListItem_Empties = '(����)';
  SSTFilterListItem_NotEmpties = '(���� �ƴ�)';
  SSTFilterListItem_All = '(���)';
  SSTFilterListItem_ClearFilter = '(��������)';
  SSTFilterListItem_SortingByAscend = '(������������)';
  SSTFilterListItem_SortingByDescend = '(������������)';
  SSTFilterListItem_ApplyFilter = '(��������)';
  SSTFilterListItem_CustomFilter = '(����ڼ���)';

  SSimpFilter_equals              = '����(=)';
  SSimpFilter_does_not_equal      = '���� ����(<>)';
  SSimpFilter_is_greate_than      = 'ŭ(>)';
  SSimpFilter_is_greate_than_or_equall_to = 'ũ�ų� ����(>=)';
  SSimpFilter_is_less_than        = '����(<)';
  SSimpFilter_is_less_than_or_equall_to = '�۰ų� ����(<=)';
  SSimpFilter_begins_with         = '���� ��';
  SSimpFilter_does_not_begin_with = '������ ���� ��';
  SSimpFilter_ends_with           = '���� ��';
  SSimpFilter_does_not_end_with   = '������ ���� ��';
  SSimpFilter_contains            = '�����ϴ� ��';
  SSimpFilter_does_not_contain    = '�������� �ʴ� ��';
  SSimpFilter_like                = '������';
  SSimpFilter_not_like            = '�������� ����';
  SSimpFilter_in_list             = '���� ����';
  SSimpFilter_not_in_list         = '���� �� ����';
  SSimpFilter_is_blank            = '����';
  SSimpFilter_is_not_blank        = '���� �ƴ�';

  SGroupingPanelText = '���⿡ ������ ����� ������ �ش� �׸����� �׷�ȭ�˴ϴ�!';

  SNoDataEh = '<�ڷ� ����>';

  // Error message
  SQuoteIsAbsentEh = '�ο��ȣ �Ῡ: ';
  SLeftBracketExpectedEh = '''('' �ʿ�: ';
  SRightBracketExpectedEh = ''')'' �ʿ�: ';
  SErrorInExpressionEh = '���Ŀ���: ';
  SUnexpectedExpressionBeforeNullEh = 'NULL �� ����ġ ���� ����: ';
  SUnexpectedExpressionAfterOperatorEh = '������ �� ����ġ ���� ����: ';
  SIncorrectExpressionEh = '�߸��� ����: ';
  SUnexpectedANDorOREh = '�߸��� AND �Ǵ� OR: ';

  SGridSelectionInfo_Sum = '�հ�';
  SGridSelectionInfo_Cnt = '�Ǽ�';
  SGridSelectionInfo_Evg = '���';
  SGridSelectionInfo_Min = '�ּҰ�';
  SGridSelectionInfo_Max = '�ִ밪';

  SFirstRecordEh = 'ù��° ���ڵ�';
  SPriorRecordEh = '���� ���ڵ�';
  SNextRecordEh = '���� ���ڵ�';
  SLastRecordEh = '������ ���ڵ�';
  SInsertRecordEh = '���ڵ� �߰�';
  SDeleteRecordEh = '���ڵ� ����';
  SEditRecordEh = '���ڵ� ����';
  SPostEditEh = '����� ���� ����';
  SCancelEditEh = '���� ���';
  SConfirmCaptionEh = 'Ȯ��';
  SRefreshRecordEh = '���� ��ħ';

  SDeleteMultipleRecordsQuestionEh = '���õ� %d ���� ���ڵ带 �����Ͻðڽ��ϱ�?';
  SDeleteAllRecordsQuestionEh = '��� ���ڵ带 �����Ͻðڽ��ϱ�?';

  SDuplicateVarNameEh = '������ "%s"��(��) �ش� �÷��ǿ� �̹� �����մϴ�.';

  SSearchPanelEditorPromptText = '�˻�...';
  SUnsupportedFieldTypeEh = '�������� �ʴ� �ʵ� ����(%s)�� �ʵ� "%s"�� �ֽ��ϴ�!';

  SSearchPanelApplyFilterEh = '������ �˻��� ���ڵ� ����';
  SSearchPanelCancelFilterEh = '���� ���� �� �˻��� ���� ����';
  SSearchPanelFindNextEh = '���� ã��';
  SSearchPanelFindPrevEh = '���� ã��';
  SSearchPanelOptionsEh = '�˻� �ɼ� ����';

  SSortingByThisColumnIsNotPossibleEh = '������ �� ���� �÷��Դϴ�.';
  SGroupingByThisColumnIsNotPossibleEh = '�׷�ȭ�� �� ���� �÷��Դϴ�.';

  SFindItemNameAllEh = '<All>';

implementation

end.
