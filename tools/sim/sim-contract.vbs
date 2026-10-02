' ---------------------------------------------------------------------------
'  Service_Manager - the contract of the workbook (GENERATED)
'
'  Written by tools\make-template.vbs: the sheets and every workbook scoped
'  name with the cell it points at. tools\sim (the model under cscript) builds
'  its workbook from this list, so the model and the template cannot disagree
'  about a name. Do not edit - rerun the generator after a layout change.
'
'  generated : 26.09.2026 20:11:34
'  names     : 114
' ---------------------------------------------------------------------------
Option Explicit

' The sheets of the workbook, in the order of the tabs.
Function ContractSheets()
    ContractSheets = Array("Dashboard", "Booking", "ClientsAdmin", "ServicesAdmin", "Appointments", "Reports", _
        "Config", "Log", "About", "_Lists", "db_clients", "db_services", _
        "db_appointments")
End Function

' Every named cell as "name|sheet|row|column".
Function ContractNames()
    ContractNames = Array("b_view_daily|Dashboard|3|1", "b_view_weekly|Dashboard|3|2", "b_view_monthly|Dashboard|3|3", "b_prev|Dashboard|3|5", "b_today|Dashboard|3|6", "b_next|Dashboard|3|7", _
        "o_view|Dashboard|4|2", "o_period|Dashboard|4|4", "o_count|Dashboard|4|7", "o_revenue|Dashboard|4|9", "o_free_slots|Dashboard|5|2", "o_status_dashboard|Dashboard|5|4", _
        "b_new_appointment|Dashboard|6|1", "b_go_clients|Dashboard|6|2", "b_go_services|Dashboard|6|3", "b_go_appointments|Dashboard|6|4", "b_go_reports|Dashboard|6|5", "b_go_config|Dashboard|6|6", _
        "b_go_about|Dashboard|6|7", "cal_title|Dashboard|8|1", "cal_first|Dashboard|10|1", "cal_last|Dashboard|160|9", "s_view|Dashboard|4|12", "s_date|Dashboard|5|12", _
        "b_refresh|Dashboard|7|11", "i_booking_date|Booking|4|2", "b_booking_date_pick|Booking|4|3", "i_booking_client|Booking|5|2", "b_booking_client_pick|Booking|5|4", "o_booking_client_id|Booking|5|3", _
        "i_booking_service|Booking|6|2", "o_booking_price|Booking|6|3", "o_booking_service_duration|Booking|6|4", "i_booking_start_h|Booking|7|2", "i_booking_start_m|Booking|7|3", "b_booking_start_pick|Booking|7|4", _
        "o_booking_start|Booking|7|6", "i_booking_dur_h|Booking|8|2", "i_booking_dur_m|Booking|8|3", "b_booking_dur_pick|Booking|8|4", "o_booking_dur|Booking|8|6", "o_booking_end|Booking|9|6", _
        "o_booking_mode|Booking|10|2", "s_edit_id|Booking|10|3", "b_booking_save|Booking|12|1", "b_booking_clear|Booking|12|2", "b_booking_cancel|Booking|12|3", "b_home_booking|Booking|12|4", _
        "o_status_booking|Booking|14|2", "i_client_name|ClientsAdmin|5|2", "i_client_phone|ClientsAdmin|6|2", "i_client_email|ClientsAdmin|7|2", "i_client_comment|ClientsAdmin|8|2", "b_clients_add|ClientsAdmin|5|4", _
        "b_clients_clear|ClientsAdmin|7|4", "b_home_clients|ClientsAdmin|9|4", "o_status_clients|ClientsAdmin|10|2", "o_clients_title|ClientsAdmin|12|1", "grid_clients_first|ClientsAdmin|14|1", "grid_clients_last|ClientsAdmin|63|6", _
        "i_service_name|ServicesAdmin|5|2", "i_service_price|ServicesAdmin|6|2", "i_service_duration|ServicesAdmin|7|2", "b_services_add|ServicesAdmin|5|4", "b_services_clear|ServicesAdmin|7|4", "b_home_services|ServicesAdmin|9|4", _
        "o_status_services|ServicesAdmin|9|2", "o_services_title|ServicesAdmin|11|1", "grid_services_first|ServicesAdmin|13|1", "grid_services_last|ServicesAdmin|62|5", "b_appointments_refresh|Appointments|3|1", "b_home_appointments|Appointments|3|3", _
        "o_status_appointments|Appointments|4|2", "o_appointments_title|Appointments|6|1", "grid_appointments_first|Appointments|8|1", "i_report_from|Reports|4|2", "i_report_to|Reports|5|2", "b_report_render|Reports|4|4", _
        "b_report_prev|Reports|5|4", "b_report_next|Reports|6|4", "b_home_reports|Reports|7|4", "b_report_from_pick|Reports|4|3", "b_report_to_pick|Reports|5|3", "o_status_reports|Reports|7|2", _
        "o_report_title|Reports|9|1", "grid_reports_first|Reports|11|1", "grid_reports_last|Reports|60|3", "i_config_day_start|Config|5|2", "i_config_day_end|Config|6|2", "i_config_step|Config|7|2", _
        "i_config_duration|Config|8|2", "i_config_currency|Config|9|2", "i_config_number_format|Config|10|2", "i_config_nonwork_days|Config|11|2", "i_config_nonwork_dates|Config|12|2", "b_config_apply|Config|5|4", _
        "b_config_defaults|Config|7|4", "b_config_nonwork_pick|Config|12|4", "b_config_nonwork_clear|Config|12|5", "b_lock_data|Config|9|4", "b_unlock_data|Config|11|4", "b_home_config|Config|13|4", _
        "o_status_config|Config|14|2", "s_schema_version|Config|6|8", "s_period_from|Config|7|8", "s_period_to|Config|8|8", "log_first|Log|2|1", "o_about_version|About|3|2", _
        "o_about_file|About|4|2", "o_about_sheets|About|5|2", "b_home_about|About|9|1", "list_clients|_Lists|2|1", "list_services|_Lists|2|2", "list_times|_Lists|2|3")
End Function
