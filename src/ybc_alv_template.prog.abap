" @TODO Populate header box
**********************************************************************
* Program Name  : YBC_ALV_TEMPLATE                                    *
* Program Title : Template Program                                    *
* Specification :                                                     *
* Created By    : Bryan Cain                                          *
* Create Date   : 12/10/2011                                          *
* DESCRIPTION   : Search for @TODO to find template areas that need changing
* SAP Version   : NW701                                               *
* ------------------------------------------------------------------- *
* Modification Log
*---------------------------------------------------------------------*
REPORT  ybc_alv_template
* @TODO set message ID
*MESSAGE-ID
NO STANDARD PAGE HEADING.

*----------------------------------------------------------------------*
*   Tables
*----------------------------------------------------------------------*

*----------------------------------------------------------------------*
*       CLASS lcl_routines DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_routines DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
* this will create the alv table object when you first call a method of the class
      class_constructor,
* this will call the F4 help for the layout.
      f4_help
        RETURNING VALUE(vari) TYPE disvariant-variant,
* this will build the output and display the alv
      display_alv
        IMPORTING vari TYPE disvariant-variant OPTIONAL PREFERRED PARAMETER vari,
      "double click event handler
      on_double_click FOR EVENT double_click OF cl_salv_events_table
        IMPORTING row column,
      "custom user commands
      user_command FOR EVENT added_function OF cl_salv_events_table
        IMPORTING e_salv_function,
      "export to xlsx
      export_excel.
    " @TODO - define other event handlers here
* @TODO - define your own methods here

  PRIVATE SECTION.
*----------------------------------------------------------------------*
*   Type definitions
*----------------------------------------------------------------------*
    TYPES: BEGIN OF ty_output,
             "@TODO - build your output table type here.  If you use data types that have
             "the descriptions you need (say, EBELN for PO number rather than NUMC 10),
             "the ALV classes will build your fieldcatalog for you.  If you use
             "generics, you'll need to update the column descriptions etc, later.

             ebeln TYPE char10, "just an example
             text  TYPE char40, "just an example
           END OF ty_output,
           "@TODO - consider changing the below table type to hashed or sorted, depending on requirements
           tt_output TYPE STANDARD TABLE OF ty_output INITIAL SIZE 0.

    CLASS-DATA:
      oref_table         TYPE REF TO cl_salv_table,
      oref_functions     TYPE REF TO cl_salv_functions,
      oref_std_functions TYPE REF TO cl_salv_functions_list,
      oref_columns       TYPE REF TO cl_salv_columns_table,
      oref_column        TYPE REF TO cl_salv_column_table,
      oref_layout        TYPE REF TO cl_salv_layout,
      oref_display       TYPE REF TO cl_salv_display_settings,
      oref_sorts         TYPE REF TO cl_salv_sorts,
      oref_aggregations  TYPE REF TO cl_salv_aggregations,
      oref_events        TYPE REF TO cl_salv_events_table,
      oref_header        TYPE REF TO cl_salv_form_layout_grid,
      oref_h_text        TYPE REF TO cl_salv_form_text,
      l_layout_key       TYPE salv_s_layout_key,
      l_layout           TYPE salv_s_layout,
      t_output           TYPE tt_output.

ENDCLASS.                    "lcl_routines DEFINITION


*----------------------------------------------------------------------*
*   Data definition
*----------------------------------------------------------------------*

*----------------------------------------------------------------------*
*   Constants
*----------------------------------------------------------------------*

*----------------------------------------------------------------------*
*   Selection-Screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE text-t01.
* @TODO create selection screen
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE text-t02.
PARAMETERS: p_vari TYPE disvariant-variant .
SELECTION-SCREEN END OF BLOCK b2.

*----------------------------------------------------------------------*
*   Initialization
*----------------------------------------------------------------------*
INITIALIZATION.

*----------------------------------------------------------------------*
*   At selection-screen
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_vari.

* call the f4 help
  p_vari = lcl_routines=>f4_help( ).

*----------------------------------------------------------------------*
*   Start of selection
*----------------------------------------------------------------------*
START-OF-SELECTION.

* @TODO data retrieval, etc

*----------------------------------------------------------------------*
*   END-OF-SELECTION
*----------------------------------------------------------------------*
END-OF-SELECTION.

* display the alv
  lcl_routines=>display_alv( p_vari ).

*----------------------------------------------------------------------*
*       CLASS lcl_routines IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_routines IMPLEMENTATION.
  METHOD class_constructor.

* instantiate the table object in the constructor so it is available to
* all other methods of the local class
    TRY .
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table   = oref_table
          CHANGING
            t_table        = t_output
               ).
      CATCH cx_salv_msg.

    ENDTRY.

* instantiate the layout class
    oref_layout = oref_table->get_layout( ).
    l_layout_key-report = sy-repid.
    oref_layout->set_key( l_layout_key ).

* only let the user save user-specific layouts
    oref_layout->set_save_restriction( cl_salv_layout=>restrict_user_dependant ).

  ENDMETHOD.                    "constructor
  METHOD f4_help.

* call the f4 help and set the layout variable.
    l_layout = oref_layout->f4_layouts( ).
    vari = l_layout-layout.

  ENDMETHOD.                                                "f4_help

  METHOD display_alv.

    DATA: stext TYPE scrtext_s,
          mtext TYPE scrtext_m,
          ltext TYPE scrtext_l.

    "@TODO remove this line.  it will add a blank record to your output.
    "it was put here so you can execute the template for traing purposes
    APPEND INITIAL LINE TO t_output.

    "@TODO - YOU MUST COPY the ALV_STANDARD gui status in SE41 in order to use the custom XLSX export routine.
    oref_table->set_screen_status(
        report        = sy-repid
        pfstatus      = 'ALV_STANDARD'
*        set_functions = C_FUNCTIONS_NONE
           ).

* create the fucntion object and force all to available.
    oref_functions = oref_table->get_functions( ).
    oref_functions->set_all( abap_true ).




* create the display settings object and set the various attributes
    oref_display = oref_table->get_display_settings( ).
    oref_display->set_striped_pattern( cl_salv_display_settings=>true ).
    oref_display->set_fit_column_to_table_size( cl_salv_display_settings=>true ).
* @TODO set list header
*  oref_display->set_list_header( '' ).

* create header object
    CREATE OBJECT oref_header.

*   Title in Bold
    oref_header->create_header_information( EXPORTING row = 1 column = 1 text = sy-title ). "@TODO change title text
*   other header information
    "@TODO change other info
*    oref_h_text = oref_header->create_text( row = 2 column = 1 text = '' ).


*   set the top of list using the header for Online.
    oref_table->set_top_of_list( oref_header ).

*   set the top of list using the header for Print.
    oref_table->set_top_of_list_print( oref_header ).

* create the columns object and optimize
    oref_columns = oref_table->get_columns( ).
    oref_columns->set_optimize( value  = cl_salv_columns_table=>if_salv_c_bool_sap~true  ).

* @TODO - set custom column headers
    TRY.
        stext = mtext = ltext = 'Description'.
        oref_column ?= oref_columns->get_column( 'TEXT' ).
        oref_column->set_short_text( value = stext   ). "also can set output length, etc by calling methods of this class.
        oref_column->set_medium_text( value = mtext ).
        oref_column->set_long_text( value = ltext ).
      CATCH cx_salv_not_found.
    ENDTRY.

* set the layout variant if one has been chosen.
    IF vari IS NOT INITIAL.
      oref_layout->set_initial_layout( value = vari   ).
    ELSE.

* set totals
      oref_aggregations = oref_table->get_aggregations( ).
* @TODO set totals
*    TRY .
*        oref_aggregations->add_aggregation( columnname  = '' ).
*      CATCH cx_salv_data_error.
*      CATCH cx_salv_not_found.
*      CATCH cx_salv_existing.
*
*    ENDTRY.


* update sort info
      oref_sorts = oref_table->get_sorts( ).
* @TODO set sorts
*    TRY .
*        oref_sorts->add_sort( columnname = ''
*                              subtotal   = if_salv_c_bool_sap=>true ).
*      CATCH cx_salv_data_error.
*      CATCH cx_salv_not_found.
*      CATCH cx_salv_existing.
*
*    ENDTRY.
    ENDIF.

    " set the event handler for double click
    oref_events = oref_table->get_event( ).
    SET HANDLER lcl_routines=>on_double_click FOR oref_events.
    " set the event handler for custom buttons
    SET HANDLER lcl_routines=>user_command FOR oref_events.

    "@TODO set other event handlers as necessary

* display the table.
    oref_table->display( ).

  ENDMETHOD.                    "display_alv

  METHOD on_double_click.
    "@TODO - define double click functionality

  ENDMETHOD.                    "on_double_click

  METHOD user_command.
    CASE e_salv_function.
      WHEN '&XLSX'.
        export_excel( ).
        "@TODO add other custom handlers if necessary.
    ENDCASE.
  ENDMETHOD.                    "user_command

  METHOD export_excel.
    DATA: lo_converter         TYPE REF TO zcl_excel_converter,
          l_path               TYPE string,  " local dir
          lv_default_file_name TYPE string,
          lv_file_name         TYPE string,
          lo_excel             TYPE REF TO zcl_excel,
          lv_action            TYPE i.

    "build default filename.
    CONCATENATE sy-repid '.xlsx' INTO lv_default_file_name.
    CONDENSE lv_default_file_name.
    "front slashes in windows file names causes issues
    IF lv_default_file_name(9) = '/LUMBERL/'.
      SHIFT lv_default_file_name LEFT BY 9 PLACES.
    ENDIF.

    "set default file
    cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = l_path ).
    cl_gui_cfw=>flush( ).

    "let user choose file name / location
    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
         default_extension    = '.xlsx'
         default_file_name    = lv_default_file_name
      CHANGING
        filename             = lv_default_file_name
        path                 = l_path
        fullpath             = lv_file_name
        user_action          = lv_action
      EXCEPTIONS
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        OTHERS               = 4
           ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    "@TODO add any necessary custom exit handling
    CASE lv_action.
      WHEN cl_gui_frontend_services=>action_ok OR cl_gui_frontend_services=>action_replace.
        "do nothing - class above set file variable
      WHEN cl_gui_frontend_services=>action_cancel.
        EXIT.
      WHEN OTHERS.
        EXIT.
    ENDCASE.

    "create excel object
    CREATE OBJECT lo_converter.
    TRY.
        lo_converter->convert(
          EXPORTING
            io_alv        = oref_table
            it_table      = t_output
            i_row_int     = 1
            i_column_int  = 1
            i_table       = abap_true
          CHANGING co_excel = lo_excel ).
      CATCH zcx_excel .
    ENDTRY.

    "output file
    lo_converter->write_file( i_path = lv_file_name ).
  ENDMETHOD.                    "export_excel

* @TODO implement other event handlers here
* @TODO implement your own methods here

ENDCLASS.                    "lcl_routines IMPLEMENTATION
