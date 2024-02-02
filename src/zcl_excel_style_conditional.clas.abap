class ZCL_EXCEL_STYLE_CONDITIONAL definition
  public
  final
  create public .

public section.

  constants C_CFVO_TYPE_FORMULA type ZEXCEL_CONDITIONAL_TYPE value 'formula' ##NO_TEXT.
  constants C_CFVO_TYPE_MAX type ZEXCEL_CONDITIONAL_TYPE value 'max' ##NO_TEXT.
  constants C_CFVO_TYPE_MIN type ZEXCEL_CONDITIONAL_TYPE value 'min' ##NO_TEXT.
  constants C_CFVO_TYPE_NUMBER type ZEXCEL_CONDITIONAL_TYPE value 'num' ##NO_TEXT.
  constants C_CFVO_TYPE_PERCENT type ZEXCEL_CONDITIONAL_TYPE value 'percent' ##NO_TEXT.
  constants C_CFVO_TYPE_PERCENTILE type ZEXCEL_CONDITIONAL_TYPE value 'percentile' ##NO_TEXT.
  constants C_ICONSET_3ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '3Arrows' ##NO_TEXT.
  constants C_ICONSET_3ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '3ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_3FLAGS type ZEXCEL_CONDITION_RULE_ICONSET value '3Flags' ##NO_TEXT.
  constants C_ICONSET_3SIGNS type ZEXCEL_CONDITION_RULE_ICONSET value '3Signs' ##NO_TEXT.
  constants C_ICONSET_3SYMBOLS type ZEXCEL_CONDITION_RULE_ICONSET value '3Symbols' ##NO_TEXT.
  constants C_ICONSET_3SYMBOLS2 type ZEXCEL_CONDITION_RULE_ICONSET value '3Symbols2' ##NO_TEXT.
  constants C_ICONSET_3TRAFFICLIGHTS type ZEXCEL_CONDITION_RULE_ICONSET value '' ##NO_TEXT.
  constants C_ICONSET_3TRAFFICLIGHTS2 type ZEXCEL_CONDITION_RULE_ICONSET value '3TrafficLights2' ##NO_TEXT.
  constants C_ICONSET_4ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '4Arrows' ##NO_TEXT.
  constants C_ICONSET_4ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '4ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_4RATING type ZEXCEL_CONDITION_RULE_ICONSET value '4Rating' ##NO_TEXT.
  constants C_ICONSET_4REDTOBLACK type ZEXCEL_CONDITION_RULE_ICONSET value '4RedToBlack' ##NO_TEXT.
  constants C_ICONSET_4TRAFFICLIGHTS type ZEXCEL_CONDITION_RULE_ICONSET value '4TrafficLights' ##NO_TEXT.
  constants C_ICONSET_5ARROWS type ZEXCEL_CONDITION_RULE_ICONSET value '5Arrows' ##NO_TEXT.
  constants C_ICONSET_5ARROWSGRAY type ZEXCEL_CONDITION_RULE_ICONSET value '5ArrowsGray' ##NO_TEXT.
  constants C_ICONSET_5QUARTERS type ZEXCEL_CONDITION_RULE_ICONSET value '5Quarters' ##NO_TEXT.
  constants C_ICONSET_5RATING type ZEXCEL_CONDITION_RULE_ICONSET value '5Rating' ##NO_TEXT.
  constants C_OPERATOR_BEGINSWITH type ZEXCEL_CONDITION_OPERATOR value 'beginsWith' ##NO_TEXT.
  constants C_OPERATOR_BETWEEN type ZEXCEL_CONDITION_OPERATOR value 'between' ##NO_TEXT.
  constants C_OPERATOR_CONTAINSTEXT type ZEXCEL_CONDITION_OPERATOR value 'containsText' ##NO_TEXT.
  constants C_OPERATOR_ENDSWITH type ZEXCEL_CONDITION_OPERATOR value 'endsWith' ##NO_TEXT.
  constants C_OPERATOR_EQUAL type ZEXCEL_CONDITION_OPERATOR value 'equal' ##NO_TEXT.
  constants C_OPERATOR_GREATERTHAN type ZEXCEL_CONDITION_OPERATOR value 'greaterThan' ##NO_TEXT.
  constants C_OPERATOR_GREATERTHANOREQUAL type ZEXCEL_CONDITION_OPERATOR value 'greaterThanOrEqual' ##NO_TEXT.
  constants C_OPERATOR_LESSTHAN type ZEXCEL_CONDITION_OPERATOR value 'lessThan' ##NO_TEXT.
  constants C_OPERATOR_LESSTHANOREQUAL type ZEXCEL_CONDITION_OPERATOR value 'lessThanOrEqual' ##NO_TEXT.
  constants C_OPERATOR_NONE type ZEXCEL_CONDITION_OPERATOR value '' ##NO_TEXT.
  constants C_OPERATOR_NOTCONTAINS type ZEXCEL_CONDITION_OPERATOR value 'notContains' ##NO_TEXT.
  constants C_OPERATOR_NOTEQUAL type ZEXCEL_CONDITION_OPERATOR value 'notEqual' ##NO_TEXT.
  constants C_RULE_CELLIS type ZEXCEL_CONDITION_RULE value 'cellIs' ##NO_TEXT.
  constants C_RULE_CONTAINSTEXT type ZEXCEL_CONDITION_RULE value 'containsText' ##NO_TEXT.
  constants C_RULE_DATABAR type ZEXCEL_CONDITION_RULE value 'dataBar' ##NO_TEXT.
  constants C_RULE_EXPRESSION type ZEXCEL_CONDITION_RULE value 'expression' ##NO_TEXT.
  constants C_RULE_ICONSET type ZEXCEL_CONDITION_RULE value 'iconSet' ##NO_TEXT.
  constants C_RULE_COLORSCALE type ZEXCEL_CONDITION_RULE value 'colorScale' ##NO_TEXT.
  constants C_RULE_NONE type ZEXCEL_CONDITION_RULE value 'none' ##NO_TEXT.
  constants C_RULE_TOP10 type ZEXCEL_CONDITION_RULE value 'top10' ##NO_TEXT.
  constants C_RULE_ABOVE_AVERAGE type ZEXCEL_CONDITION_RULE value 'aboveAverage' ##NO_TEXT.
  constants C_SHOWVALUE_FALSE type ZEXCEL_CONDITIONAL_SHOW_VALUE value 0 ##NO_TEXT.
  constants C_SHOWVALUE_TRUE type ZEXCEL_CONDITIONAL_SHOW_VALUE value 1 ##NO_TEXT.
  data MODE_CELLIS type ZEXCEL_CONDITIONAL_CELLIS .
  data MODE_COLORSCALE type ZEXCEL_CONDITIONAL_COLORSCALE .
  data MODE_DATABAR type ZEXCEL_CONDITIONAL_DATABAR .
  data MODE_EXPRESSION type ZEXCEL_CONDITIONAL_EXPRESSION .
  data MODE_ICONSET type ZEXCEL_CONDITIONAL_ICONSET .
  data MODE_TOP10 type ZEXCEL_CONDITIONAL_TOP10 .
  data MODE_ABOVE_AVERAGE type ZEXCEL_CONDITIONAL_ABOVE_AVG .
  data PRIORITY type ZEXCEL_STYLE_PRIORITY value 1 ##NO_TEXT.
  data RULE type ZEXCEL_CONDITION_RULE .

  methods CONSTRUCTOR .
  methods GET_DIMENSION_RANGE
    returning
      value(EP_DIMENSION_RANGE) type STRING .
  methods SET_RANGE
    importing
      !IP_START_ROW type ZEXCEL_CELL_ROW
      !IP_START_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_STOP_ROW type ZEXCEL_CELL_ROW
      !IP_STOP_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  class-methods FACTORY_COND_STYLE_ICONSET
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IV_ICON_TYPE type ZEXCEL_CONDITION_RULE_ICONSET default C_ICONSET_3TRAFFICLIGHTS2
      !IV_CFVO1_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO1_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO2_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO2_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO3_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO3_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO4_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO4_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_CFVO5_TYPE type ZEXCEL_CONDITIONAL_TYPE default C_CFVO_TYPE_PERCENT
      !IV_CFVO5_VALUE type ZEXCEL_CONDITIONAL_VALUE optional
      !IV_SHOWVALUE type ZEXCEL_CONDITIONAL_SHOW_VALUE default ZCL_EXCEL_STYLE_CONDITIONAL=>C_SHOWVALUE_TRUE
    returning
      value(RV_STYLE_CONDITIONAL) type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
protected section.
private section.

  data START_CELL type ZEXCEL_S_CELL_DATA .
  data STOP_CELL type ZEXCEL_S_CELL_DATA .
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_CONDITIONAL IMPLEMENTATION.


METHOD constructor.

  DATA: ls_iconset TYPE zexcel_conditional_iconset.
  ls_iconset-iconset     = zcl_excel_style_conditional=>c_iconset_3trafficlights.
  ls_iconset-cfvo1_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo1_value = '0'.
  ls_iconset-cfvo2_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo2_value = '20'.
  ls_iconset-cfvo3_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo3_value = '40'.
  ls_iconset-cfvo4_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo4_value = '60'.
  ls_iconset-cfvo5_type  = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo5_value = '80'.


  me->rule          = zcl_excel_style_conditional=>c_rule_none.
*  me->iconset->operator    = zcl_excel_style_conditional=>c_operator_none.
  me->mode_iconset  = ls_iconset.
  me->priority      = 1.

* inizialize dimension range
  me->stop_cell-cell_row     = 1.
  me->stop_cell-cell_column  = 1.
  me->start_cell-cell_row     = 1.
  me->start_cell-cell_column  = 1.
ENDMETHOD.


  METHOD factory_cond_style_iconset.
  ENDMETHOD.


METHOD get_dimension_range.
  IF stop_cell EQ start_cell. "only one cell
    ep_dimension_range = start_cell-cell_coords.
  ELSE.
    CONCATENATE start_cell-cell_coords ':' stop_cell-cell_coords INTO ep_dimension_range.
  ENDIF.
ENDMETHOD.


METHOD set_range.
  DATA: lv_column    TYPE zexcel_cell_column,
        lv_row_alpha TYPE string.

  lv_column = zcl_excel_common=>convert_column2int( ip_stop_column ).
  stop_cell-cell_row     = 1.
  stop_cell-cell_column  = lv_column.
  lv_row_alpha = ip_stop_row.
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE ip_stop_column lv_row_alpha INTO stop_cell-cell_coords.

  lv_column = zcl_excel_common=>convert_column2int( ip_start_column ).
  start_cell-cell_row     = 1.
  start_cell-cell_column  = lv_column.
  lv_row_alpha = ip_start_row.
  SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
  SHIFT lv_row_alpha LEFT DELETING LEADING space.
  CONCATENATE ip_start_column lv_row_alpha INTO start_cell-cell_coords.
ENDMETHOD.
ENDCLASS.
