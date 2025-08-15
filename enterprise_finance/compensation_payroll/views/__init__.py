# views/__init__.py
from .dashboard import dashboard

from .personnel import personnel_list, create_personnel, update_personnel, delete_personnel, export_personnels_to_excel, \
    personnel_graph_view

from .organization import organizational_profile, create_organizational_profile, update_organizational_profile, \
    delete_organizational_profile

from .payroll_period import payroll_period_list, create_payroll_period, \
    update_payroll_period, delete_payroll_period

from .payroll_month_and_component import payroll_month_and_component_list, create_payroll_month_and_component, \
    update_payroll_month_and_component, delete_payroll_month_and_component

from .regular_payroll import payroll_month_list, regular_payroll_list, regular_payroll_detail, create_regular_payroll, \
    update_regular_payroll, \
    delete_regular_payroll, export_regular_payroll_to_excel, payroll_processing_graphs, get_regular_payroll_by_month_report, \
    payroll_month_for_summary_report, payroll_month_for_journal_entry_report, \
    regular_payroll_summary, payroll_by_year_summary_report, regular_payroll_journal_entry, university_cost_sharing_view

from .earning_adjustment import earning_object_list, earnings_object_detail, \
    create_earning_adjustment, update_earning_adjustment, delete_earning_adjustment, earning_per_adjusted_month, \
    monthly_earning_adjustment, monthly_earning_adjustment_total, export_earning_adjustment_list_to_excel, \
    export_earning_per_adjusted_month_to_excel, export_monthly_earning_adjustment_to_excel, \
    export_monthly_earning_adjustment_aggregate

from .deduction_adjustment import deduction_object_list, deduction_object_detail, \
    create_deduction_adjustment, update_deduction_adjustment, delete_deduction_adjustment, deduction_per_adjusted_month, \
    monthly_deduction_adjustment, monthly_deduction_adjustment_total, export_deduction_adjustment_list_to_excel, \
    export_deduction_per_adjusted_month_to_excel, export_monthly_deduction_adjustment_to_excel, \
    export_monthly_deduction_adjustment_aggregate

from .combined_personnel import get_combined_personnel_payroll_context, combined_personnel_payroll_list, \
    combined_personnel_adjustment_list, combined_personnel_total, combined_personnel_expense, combined_personnel_net_income, \
    combined_employee_pension, combined_personnel_employment_income_tax, export_combined_personnel_list, export_personnel_total_adjustment, \
    export_combined_personnel_total, export_combined_personnel_expense, \
    export_combined_personnel_net_income, export_combined_personnel_employment_tax, export_combined_personnel_pension, \
    combined_personnel_detail, export_combined_personnel_detail

from .combined_monthly import get_combined_monthly_detail, monthly_combined_detail, \
    monthly_combined_adjustment_journal, monthly_combined_adjustment_summary, export_combined_monthly_detail_to_excel, \
    combined_monthly_graph_view, monthly_combined_summary, export_combined_monthly_summary_to_excel

from .combined_yearly import get_combined_yearly_detail, yearly_combined_detail_view, \
    export_combined_yearly_detail_to_excel, yearly_combined_summary_view, export_combined_yearly_summary_to_excel

from .severance_pay import severance_pay_list,  severance_pay_detail, create_severance_pay, update_severance_pay, \
    delete_severance_pay, severance_pay_report, export_severance_pay_to_excel

from .adsence_deduction import absence_deduction_list, create_absence_deduction, update_absence_deduction, \
    delete_absence_deduction
