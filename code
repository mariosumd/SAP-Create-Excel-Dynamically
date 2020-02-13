  TYPES: BEGIN OF ltype_header,
           h(30) TYPE c,
         END OF ltype_header.

  DATA: l_t_fieldnames    TYPE TABLE OF ltype_header,
        l_wa_fieldnames   TYPE ltype_header,
        l_t_catalog       TYPE lvc_t_fcat,
        l_wa_catalog      TYPE lvc_s_fcat,
        l_file            TYPE rlgrap-filename,
        l_typ             TYPE REF TO cl_abap_elemdescr,
        l_t_tot_comp      TYPE cl_abap_structdescr=>component_table,
        l_wa_comp         LIKE LINE OF l_t_comp_comp,
        lo_new_type       TYPE REF TO cl_abap_structdescr,
        lo_table_type     TYPE REF TO cl_abap_tabledescr,
        l_tref            TYPE REF TO data,
        l_dy_line         TYPE REF TO data.

  FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE,
                 <fs_wa>,
                 <fs_field_out>,
                 <fs_field_exp>.

  " Getting live fieldcatalog
  CALL METHOD gv_grid->get_frontend_fieldcatalog
    IMPORTING
      et_fieldcatalog = l_t_catalog.

  " Looping the catalog conditioning by fields showing
  LOOP AT l_t_catalog INTO l_wa_catalog WHERE no_out = space.
    " Getting Excel column names
    l_wa_fieldnames-h = l_wa_catalog-reptext.

    APPEND l_wa_fieldnames TO l_t_fieldnames.
	
	" Creating dynamicaly the excel's structure
    CLEAR: l_typ, l_wa_comp.
	
	" Type: CHAR200
    l_typ = cl_abap_elemdescr=>get_c( p_length = '200' ).

    l_wa_comp-type = l_typ.                   "Tipo   del campo
    l_wa_comp-name = l_wa_catalog-fieldname. "Nombre del campo
    APPEND l_wa_comp TO l_t_tot_comp.
  ENDLOOP.

* Creating new type with the structure created
  lo_new_type = cl_abap_structdescr=>create( l_t_tot_comp ).

* Creating new table
  lo_table_type = cl_abap_tabledescr=>create( lo_new_type ).

* Creating field-symbol table with the dynamic structure
  CREATE DATA l_tref TYPE HANDLE lo_table_type.
  ASSIGN l_tref->* TO <fs_table>.

  "Data
  LOOP AT it_out.
    UNASSIGN: <fs_wa>, <fs_field_out>, <fs_field_exp>.

* Create the dinamic workarea
    CREATE DATA l_dy_line LIKE LINE OF <fs_table>.
    ASSIGN l_dy_line->* TO <fs_wa>.
	
	" Loop the catalog conditioning by fields showing
    LOOP AT l_t_catalog INTO l_wa_catalog WHERE no_out = space.
	  " Getting 
      ASSIGN COMPONENT l_wa_catalog-fieldname OF STRUCTURE it_out TO <fs_field_out>.

      IF sy-subrc = 0.
        ASSIGN COMPONENT l_wa_catalog-fieldname OF STRUCTURE <fs_wa> TO <fs_field_exp>.

        IF sy-subrc = 0 AND <fs_field_out> IS NOT INITIAL.
           WRITE <fs_field_out> TO <fs_field_exp>.
        ENDIF.
      ENDIF.
    ENDLOOP.

    APPEND <fs_wa> TO <fs_table>.
  ENDLOOP.
  
  " Getting filename
  CALL FUNCTION 'KD_GET_FILENAME_ON_F4'
    CHANGING
      file_name     = l_file
    EXCEPTIONS
      mask_too_long = 1
      OTHERS        = 2.

  IF sy-subrc = 0.
	
	" Downloading excel sheet
    CALL FUNCTION 'EXCEL_OLE_STANDARD_DAT'
      EXPORTING
        file_name                 = l_file
      TABLES
        data_tab                  = <fs_table>
        fieldnames                = l_t_fieldnames
      EXCEPTIONS
        file_not_exist            = 1
        filename_expected         = 2
        communication_error       = 3
        ole_object_method_error   = 4
        ole_object_property_error = 5
        invalid_pivot_fields      = 6
        download_problem          = 7
        OTHERS                    = 8.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE 'S' NUMBER sy-msgno
         WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
         DISPLAY LIKE sy-msgty.
    ENDIF.

  ENDIF.
