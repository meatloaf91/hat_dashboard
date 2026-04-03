# Neo dark dashboard style (inspired by attached reference)
# Core palette
# bg: #070809
# card: #141518
# card_soft: #1B1D21
# stroke: #2B2F35
# text: #F4F6F8
# muted: #A6ADB5
# lime: #B8F35A
# orange: #F6A62E

ss_qpushbutton_01 = """
    #MainWindow {
        background-color: #070809;
        color: #F4F6F8;
    }

    QWidget {
        color: #F4F6F8;
        font-family: "Segoe UI";
        font-size: 12px;
    }

    QFrame {
        background-color: #0C0D10;
        border: 1px solid #1F2328;
        border-radius: 16px;
        padding: 10px;
    }

    QPushButton {
        background-color: #24282E;
        border: 1px solid #323841;
        border-radius: 12px;
        color: #F4F6F8;
        font-size: 12px;
        font-weight: 600;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
        padding: 4px 10px;
    }

    QPushButton:hover {
        background-color: #2A2F36;
        border: 1px solid #404854;
    }

    QPushButton:pressed {
        background-color: #191D22;
        border: 1px solid #4D5561;
    }

    QCheckBox {
        color: #D6DBE0;
        font-size: 12px;
        font-weight: 500;
        padding: 2px 0px;
        background: transparent;
    }

    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border-radius: 4px;
        border: 1px solid #555F6A;
        background: #111419;
    }

    QCheckBox::indicator:checked {
        background: #B8F35A;
        border-color: #B8F35A;
    }

    QRadioButton {
        color: #DDE1E6;
        font-size: 12px;
        font-weight: 600;
        background: transparent;
    }

    QRadioButton::indicator {
        width: 15px;
        height: 15px;
        border: 1.2px solid #7B8491;
        border-radius: 8px;
        background: #111419;
        margin-right: 6px;
    }

    QRadioButton::indicator:checked {
        background: #B8F35A;
        border: 1.2px solid #B8F35A;
    }

    QTextEdit,
    QLineEdit {
        background-color: #0B0D10;
        color: #FFFFFF;
        border: 1px solid #3A414B;
        border-radius: 14px;
        padding: 7px 10px;
        selection-background-color: #B8F35A;
        selection-color: #0A0A0A;
    }

    QTextEdit:focus,
    QLineEdit:focus {
        border: 1px solid #B8F35A;
    }

    QScrollBar:vertical {
        border: none;
        background: #15181D;
        width: 7px;
        margin: 0px;
        border-radius: 4px;
    }

    QScrollBar::handle:vertical {
        background: #58606B;
        min-height: 24px;
        border-radius: 4px;
    }

    QScrollBar::handle:vertical:hover {
        background: #B8F35A;
    }

    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
    }

    QScrollBar:horizontal {
        border: none;
        background: #15181D;
        height: 7px;
        margin: 0px;
        border-radius: 4px;
    }

    QScrollBar::handle:horizontal {
        background: #58606B;
        min-width: 24px;
        border-radius: 4px;
    }

    QScrollBar::handle:horizontal:hover {
        background: #B8F35A;
    }

    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
        width: 0px;
    }
"""

ss_frame_01 = """
    #frame {
        background-color: #070809;
        border: none;
    }

    #panel_left_top,
    #panel_left_btm {
        background-color: #111318;
        color: #FFFFFF;
        border-radius: 16px;
        padding: 8px;
        border: 1px solid #262B31;
        outline: none;
    }

    #panel_main {
        background-color: #13161B;
        border-radius: 16px;
        border: 1px solid #2A2F36;
        padding: 10px;
    }

    #panel_main QWidget {
        background: transparent;
        border: none;
    }

    #btn_menu_generate_packshot_naming,
    #btn_menu_generate_thumbnails,
    #btn_menu_generate_data_from_images,
    #btn_menu_bulk_rename,
    #btn_menu_all_trackers_status_collector,
    #btn_menu_reformat_sap_data,
    #btn_menu_compare_sap_data,
    #btn_menu_reference_collector,
    #btn_menu_description_of_matched_basics,
    #btn_menu_one_step_run {
        background-color: transparent;
        color: #E9EDF2;
        border-radius: 12px;
        border: 1px solid transparent;
        padding: 7px 12px;
        font-family: "Segoe UI";
        font-size: 13px;
        font-weight: 600;
        text-align: left;
        letter-spacing: 0.2px;
        min-width: 205px;
        max-width: 205px;
        min-height: 24px;
        max-height: 24px;
    }

    #btn_menu_generate_packshot_naming:hover,
    #btn_menu_generate_thumbnails:hover,
    #btn_menu_generate_data_from_images:hover,
    #btn_menu_bulk_rename:hover,
    #btn_menu_all_trackers_status_collector:hover,
    #btn_menu_reformat_sap_data:hover,
    #btn_menu_compare_sap_data:hover,
    #btn_menu_reference_collector:hover,
    #btn_menu_description_of_matched_basics:hover,
    #btn_menu_one_step_run:hover {
        background-color: #1E232A;
        border: 1px solid #323A44;
    }

    #btn_menu_generate_packshot_naming:checked,
    #btn_menu_generate_thumbnails:checked,
    #btn_menu_generate_data_from_images:checked,
    #btn_menu_bulk_rename:checked,
    #btn_menu_all_trackers_status_collector:checked,
    #btn_menu_reformat_sap_data:checked,
    #btn_menu_compare_sap_data:checked,
    #btn_menu_reference_collector:checked,
    #btn_menu_description_of_matched_basics:checked,
    #btn_menu_one_step_run:checked {
        background-color: rgba(184, 243, 90, 0.16);
        border: 1px solid #B8F35A;
        color: #F9FFF0;
    }

    #label_title_generate_packshot_naming,
    #label_title_generate_thumbnails,
    #label_title_generate_data_from_images,
    #label_title_bulk_rename,
    #label_title_all_trackers_status_collector,
    #label_title_reformat_sap_data,
    #label_title_compare_sap_data,
    #label_title_reference_collector,
    #label_title_description_of_matched_basics,
    #label_title_one_step_run {
        color: #FFFFFF;
        font-family: "Segoe UI Semibold";
        font-size: 30px;
        font-weight: 700;
        text-align: left;
        letter-spacing: 0.8px;
        min-width: 340px;
        max-width: 340px;
        min-height: 36px;
        max-height: 36px;
    }

    #label_note_generate_packshot_naming,
    #label_note_generate_thumbnails,
    #label_note_generate_data_from_images,
    #label_note_bulk_rename,
    #label_note_all_trackers_status_collector,
    #label_note_reformat_sap_data,
    #label_note_compare_sap_data,
    #label_note_reference_collector,
    #label_note_description_of_matched_basics,
    #label_note_title_one_step_run {
        color: #A9B0B8;
        font-family: "Segoe UI";
        font-size: 12px;
        padding: 1px 0px;
        background: transparent;
        qproperty-wordWrap: true;
    }

    #label_menu_title_body_mapper,
    #label_menu_title_basic_tools {
        background: transparent;
        color: #FFFFFF;
        border: none;
        font-family: "Segoe UI Semibold";
        font-size: 14px;
        font-weight: 700;
        text-align: left;
        padding: 2px;
        letter-spacing: 2px;
        min-width: 210px;
        max-width: 210px;
        min-height: 20px;
        max-height: 20px;
    }
"""

ss_panelmain_01 = """
    #btn_run_process_generate_packshot_naming,
    #btn_run_process_generate_thumbnails,
    #btn_run_process_generate_data_from_images,
    #btn_run_process_bulk_rename,
    #btn_run_process_all_trackers_status_collector,
    #btn_format_run_process,
    #btn_run_process_reformat_sap_data,
    #btn_run_process_compare_sap_data,
    #btn_run_process_reference_collector,
    #btn_run_process_description_of_matched_basics,
    #btn_run_process_title_one_step_run,
    #btn_collect_image_names {
        background-color: #F6A62E;
        border: 1px solid #F6A62E;
        border-radius: 12px;
        font-size: 14px;
        font-weight: 700;
        color: #101012;
        min-width: 170px;
        max-width: 170px;
        min-height: 38px;
        max-height: 38px;
        text-align: center;
    }

    #btn_run_process_generate_packshot_naming:hover,
    #btn_run_process_generate_thumbnails:hover,
    #btn_run_process_generate_data_from_images:hover,
    #btn_run_process_bulk_rename:hover,
    #btn_run_process_all_trackers_status_collector:hover,
    #btn_format_run_process:hover,
    #btn_run_process_reformat_sap_data:hover,
    #btn_run_process_compare_sap_data:hover,
    #btn_run_process_reference_collector:hover,
    #btn_run_process_description_of_matched_basics:hover,
    #btn_run_process_title_one_step_run:hover,
    #btn_collect_image_names:hover {
        background-color: #FDC059;
        border: 1px solid #FDC059;
    }

    #btn_run_process_generate_packshot_naming:pressed,
    #btn_run_process_generate_thumbnails:pressed,
    #btn_run_process_generate_data_from_images:pressed,
    #btn_run_process_bulk_rename:pressed,
    #btn_run_process_all_trackers_status_collector:pressed,
    #btn_format_run_process:pressed,
    #btn_run_process_reformat_sap_data:pressed,
    #btn_run_process_compare_sap_data:pressed,
    #btn_run_process_reference_collector:pressed,
    #btn_run_process_description_of_matched_basics:pressed,
    #btn_run_process_title_one_step_run:pressed,
    #btn_collect_image_names:pressed {
        background-color: #C57E16;
        border: 1px solid #C57E16;
        color: #0E0F11;
    }

    QTextEdit {
        background-color: #090B0F;
        color: #FDFDFD;
        border: 1px solid #3B424D;
        border-radius: 14px;
        padding: 8px 10px;
        min-width: 220px;
        max-width: 220px;
        min-height: 38px;
        max-height: 38px;
        margin-left: 10px;
    }
"""

ss_page_B_compare_sap_data = """
    #btn_pg2_existing_data,
    #btn_pg2_new_data {
        background-color: #24282E;
        border: 1px solid #343B45;
        color: #F2F5F8;
        font-size: 12px;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }
"""

ss_page_B_reference_collector = """
    #btn_pg3_images_folder,
    #btn_pg3_list_of_idhs {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }

    #label_pg3_column_letter {
        color: #EEF2F6;
        font-family: "Segoe UI";
        font-size: 13px;
        font-weight: 600;
        text-align: left;
        letter-spacing: 0.2px;
        min-width: 110px;
        max-width: 110px;
        min-height: 20px;
        max-height: 20px;
    }

    #lineEdit_pg3_column_letter {
        background-color: #0A0D10;
        border: 1px solid #3B424D;
        border-radius: 10px;
        color: #FFFFFF;
        font-weight: 700;
        min-width: 70px;
        max-width: 70px;
        min-height: 40px;
        max-height: 40px;
    }
"""

ss_page_A_generate_thumbnails = """
    #btn_pg7_images_folder,
    #btn_pg7_output {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }
"""

ss_page_B_description_of_matched_basics = """
    #btn_pg4_mapped_data,
    #btn_pg4_output,
    #btn_pg4_existing_data {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }

    #label_title_description_of_matched_basics {
        min-width: 340px;
        max-width: 340px;
    }
"""

ss_page_B_one_step_run = """
    #btn_one_step_run_master_data,
    #btn_one_step_run_new_sap_data,
    #btn_one_step_run_images_folder,
    #btn_one_step_run_output_location {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 11px;
        color: #F2F5F8;
        min-width: 128px;
        max-width: 128px;
        min-height: 28px;
        max-height: 28px;
        text-align: center;
    }

    QTextEdit {
        background-color: #0A0D10;
        color: #FFFFFF;
        border: 1px solid #3B424D;
        border-radius: 10px;
        min-width: 220px;
        max-width: 220px;
        min-height: 28px;
        max-height: 28px;
        padding: 5px 8px;
    }
"""

ss_page_A_generate_packshot_naming = """
    #btn_pg6_excel_tracker,
    #btn_pg6_output {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }

    #btn_pg6_idh_num_col,
    #btn_pg6_pack_size_col,
    #btn_pg6_pack_type_col,
    #btn_pg6_product_name_col,
    #btn_pg6_view,
    #btn_pg6_starting_row,
    #btn_pg6_ending_row {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 28px;
        max-height: 28px;
        text-align: center;
    }

    #lineEdit_pg6_idh_num_col,
    #lineEdit_pg6_pack_size_col,
    #lineEdit_pg6_pack_type_col,
    #lineEdit_pg6_product_name_col,
    #lineEdit_pg6_view,
    #lineEdit_pg6_starting_row,
    #lineEdit_pg6_ending_row {
        background-color: #0A0D10;
        color: #FFFFFF;
        letter-spacing: 0.2px;
        border: 1px solid #3B424D;
        border-radius: 10px;
        min-width: 150px;
        max-width: 150px;
        min-height: 36px;
        max-height: 36px;
        margin-left: 10px;
    }

    #lineEdit_pg6_idh_num_col:focus,
    #lineEdit_pg6_pack_size_col:focus,
    #lineEdit_pg6_pack_type_col:focus,
    #lineEdit_pg6_product_name_col:focus,
    #lineEdit_pg6_view:focus,
    #lineEdit_pg6_starting_row:focus,
    #lineEdit_pg6_ending_row:focus {
        border: 1px solid #B8F35A;
    }
"""

ss_page_A_status_collector = """
    #btn_sc_select_trackers,
    #btn_sc_output_location {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }

    #textEdit_sc_selected_trackers,
    #textEdit_sc_output_location {
        font-size: 11px;
        color: #FFFFFF;
    }
"""

ss_page_A_bulk_rename = """
    #btn_pg9_images_folder,
    #btn_pg9_excel_file_to_follow,
    #btn_pg9_output {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }
"""

ss_page_A_generate_data_from_images = """
    #btn_pg8_images_folder,
    #btn_pg8_output {
        background-color: #24282E;
        border: 1px solid #343B45;
        font-size: 12px;
        color: #F2F5F8;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }
"""

ss_reformat_sap_data = """
    #btn_format_files_to_reformat,
    #btn_format_output_location {
        background-color: #24282E;
        border: 1px solid #343B45;
        color: #F2F5F8;
        font-size: 12px;
        min-width: 152px;
        max-width: 152px;
        min-height: 36px;
        max-height: 36px;
        text-align: center;
    }

    #textEdit_format_files_to_format,
    #textEdit_format_output_location {
        font-size: 11px;
        color: #FFFFFF;
    }
"""

msg_stylesheet = """
    QMessageBox {
        background-color: #111F35;
        border: 1px solid #2D3E58;
    }

    QLabel {
        color: #F3F6F8;
        font-size: 12px;
    }

    QPushButton {
        background-color: #7A808A;
        border: 1px solid #7A808A;
        border-radius: 10px;
        color: #F3F6F8;
        font-weight: 700;
        min-width: 150px;
        max-width: 150px;
        min-height: 28px;
        max-height: 28px;
        text-align: center;
    }

    QPushButton:hover {
        background-color: #8B929D;
        border: 1px solid #8B929D;
    }

    QPushButton:pressed {
        background-color: #666C75;
        border: 1px solid #666C75;
        color: #F3F6F8;
    }
"""

add_on_btn = """
    #btn_menu_clear_all_fields {
        background-color: #B8F35A;
        border: 1px solid #B8F35A;
        border-radius: 12px;
        font-size: 14px;
        font-weight: 700;
        color: #101012;
        min-width: 170px;
        max-width: 170px;
        min-height: 38px;
        max-height: 38px;
        text-align: center;
    }

    #btn_menu_clear_all_fields:hover {
        background-color: #CBF880;
        border: 1px solid #CBF880;
    }

    #btn_menu_clear_all_fields:pressed {
        background-color: #95CC41;
        border: 1px solid #95CC41;
    }
"""
