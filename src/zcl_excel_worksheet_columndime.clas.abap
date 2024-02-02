class ZCL_EXCEL_WORKSHEET_COLUMNDIME definition
  public
  final
  create public .

public section.

  methods CONSTRUCTOR
    importing
      !IP_INDEX type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IP_EXCEL type ref to ZCL_EXCEL .
  methods GET_AUTO_SIZE
    returning
      value(R_AUTO_SIZE) type ABAP_BOOL .
  methods GET_COLLAPSED
    returning
      value(R_COLLAPSED) type ABAP_BOOL .
  methods GET_COLUMN_INDEX
    returning
      value(R_COLUMN_INDEX) type INT4 .
  methods GET_OUTLINE_LEVEL
    returning
      value(R_OUTLINE_LEVEL) type INT4 .
  methods GET_VISIBLE
    returning
      value(R_VISIBLE) type ABAP_BOOL .
  methods GET_WIDTH
    returning
      value(R_WIDTH) type FLOAT .
  methods GET_XF_INDEX
    returning
      value(R_XF_INDEX) type INT4 .
  methods SET_AUTO_SIZE
    importing
      !IP_AUTO_SIZE type ABAP_BOOL
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_COLLAPSED
    importing
      !IP_COLLAPSED type ABAP_BOOL
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_COLUMN_INDEX
    importing
      !IP_INDEX type ZEXCEL_CELL_COLUMN_ALPHA
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_OUTLINE_LEVEL
    importing
      !IP_OUTLINE_LEVEL type INT4 .
  methods SET_VISIBLE
    importing
      !IP_VISIBLE type ABAP_BOOL
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_WIDTH
    importing
      !IP_WIDTH type SIMPLE
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_XF_INDEX
    importing
      !IP_XF_INDEX type INT4
    returning
      value(R_WORKSHEET_COLUMNDIME) type ref to ZCL_EXCEL_WORKSHEET_COLUMNDIME .
  methods SET_COLUMN_STYLE_BY_GUID
    importing
      !IP_STYLE_GUID type ZEXCEL_CELL_STYLE .
  methods GET_COLUMN_STYLE_GUID
    returning
      value(EP_STYLE_GUID) type ZEXCEL_CELL_STYLE .
protected section.
private section.

  data COLUMN_INDEX type INT4 .
  data WIDTH type FLOAT .
  data AUTO_SIZE type ABAP_BOOL .
  data VISIBLE type ABAP_BOOL .
  data OUTLINE_LEVEL type INT4 .
  data COLLAPSED type ABAP_BOOL .
  data XF_INDEX type INT4 .
  data STYLE_GUID type ZEXCEL_CELL_STYLE .
  data EXCEL type ref to ZCL_EXCEL .
  data WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
ENDCLASS.



CLASS ZCL_EXCEL_WORKSHEET_COLUMNDIME IMPLEMENTATION.


METHOD constructor.
  me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
  me->width         = -1.
  me->auto_size     = abap_false.
  me->visible       = abap_true.
  me->outline_level	= 0.
  me->collapsed     = abap_false.
  me->excel         = ip_excel.        "ins issue #157 - Allow Style for columns
  me->worksheet     = ip_worksheet.    "ins issue #157 - Allow Style for columns

  " set default index to cellXf
  me->xf_index = 0.

ENDMETHOD.


METHOD get_auto_size.
  r_auto_size = me->auto_size.
ENDMETHOD.


METHOD get_collapsed.
  r_collapsed = me->collapsed.
ENDMETHOD.


METHOD get_column_index.
  r_column_index = me->column_index.
ENDMETHOD.


METHOD get_column_style_guid.
  IF me->style_guid IS NOT INITIAL.
    ep_style_guid = me->style_guid.
  ELSE.
    ep_style_guid = me->worksheet->zif_excel_sheet_properties~get_style( ).
  ENDIF.
ENDMETHOD.


METHOD get_outline_level.
  r_outline_level = me->outline_level.
ENDMETHOD.


METHOD get_visible.
  r_visible = me->visible.
ENDMETHOD.


METHOD get_width.
  r_width = me->width.
ENDMETHOD.


METHOD get_xf_index.
  r_xf_index = me->xf_index.
ENDMETHOD.


METHOD set_auto_size.
  me->auto_size = ip_auto_size.
  r_worksheet_columndime = me.
ENDMETHOD.


METHOD set_collapsed.
  me->collapsed = ip_collapsed.
  r_worksheet_columndime = me.
ENDMETHOD.


METHOD set_column_index.
  me->column_index = zcl_excel_common=>convert_column2int( ip_index ).
  r_worksheet_columndime = me.
ENDMETHOD.


METHOD set_column_style_by_guid.
  DATA: stylemapping TYPE zexcel_s_stylemapping.

  IF me->excel IS NOT BOUND.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Internal error - reference to ZCL_EXCEL not bound'.
  ENDIF.
  TRY.
      stylemapping = me->excel->get_style_to_guid( ip_style_guid ).
      me->style_guid = stylemapping-guid.

    CATCH zcx_excel .
      EXIT.  " leave as is in case of error
  ENDTRY.

ENDMETHOD.


METHOD set_outline_level.
  me->outline_level = ip_outline_level.
ENDMETHOD.


METHOD set_visible.
  me->visible = ip_visible.
  r_worksheet_columndime = me.
ENDMETHOD.


METHOD set_width.
  TRY.
      me->width = ip_width.
      r_worksheet_columndime = me.
    CATCH cx_sy_conversion_no_number.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Unable to interpret width as number'.
  ENDTRY.
ENDMETHOD.


METHOD set_xf_index.
  me->xf_index = ip_xf_index.
  r_worksheet_columndime = me.
ENDMETHOD.
ENDCLASS.
