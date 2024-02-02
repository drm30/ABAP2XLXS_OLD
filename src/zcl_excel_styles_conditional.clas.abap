class ZCL_EXCEL_STYLES_CONDITIONAL definition
  public
  final
  create public .

public section.

  methods ADD
    importing
      !IP_STYLE_CONDITIONAL type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type ZEXCEL_ACTIVE_WORKSHEET
    returning
      value(EO_STYLE_CONDITIONAL) type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_STYLE_CONDITIONAL type ref to ZCL_EXCEL_STYLE_CONDITIONAL .
  methods SIZE
    returning
      value(EP_SIZE) type I .
protected section.
PRIVATE SECTION.

  DATA styles_conditional TYPE REF TO cl_object_collection .
ENDCLASS.



CLASS ZCL_EXCEL_STYLES_CONDITIONAL IMPLEMENTATION.


METHOD add.
  styles_conditional->add( ip_style_conditional ).
ENDMETHOD.


METHOD clear.
  styles_conditional->clear( ).
ENDMETHOD.


METHOD constructor.

  CREATE OBJECT styles_conditional.

ENDMETHOD.


METHOD get.
  DATA lv_index TYPE i.
  lv_index = ip_index.
  eo_style_conditional ?= styles_conditional->if_object_collection~get( lv_index ).
ENDMETHOD.


METHOD get_iterator.
  eo_iterator ?= styles_conditional->if_object_collection~get_iterator( ).
ENDMETHOD.


METHOD is_empty.
  is_empty = styles_conditional->if_object_collection~is_empty( ).
ENDMETHOD.


METHOD remove.
  styles_conditional->remove( ip_style_conditional ).
ENDMETHOD.


METHOD size.
  ep_size = styles_conditional->if_object_collection~size( ).
ENDMETHOD.
ENDCLASS.
