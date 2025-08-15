from django.urls import path
from . import views


urlpatterns = [
    path('', views.dashboard, name='dashboard'),

    # organizational_profile
    path('organizational-profile/', views.organizational_profile, name='organizational_profile'),
    path('organizational-profile/create/', views.create_organizational_profile, name='create_organizational_profile'),
    path('organizational-profile/update/<int:pk>/', views.update_organizational_profile, name='update_organizational_profile'),
    path('organizational-profile/delete/<int:pk>/', views.delete_organizational_profile, name='delete_organizational_profile'),

    # View the personnel_list list
    path('personnel-list/', views.personnel_list, name='personnel_list'),
    path('export-personnel/', views.export_personnels_to_excel, name='export_personnel'),
    path('personnel-list/create/', views.create_personnel, name='create_personnel'),
    path('personnel-list/update/<int:pk>/', views.update_personnel, name='update_personnel'),
    path('personnel-list/delete/<int:pk>/', views.delete_personnel, name='delete_personnel'),
    path('personnel-graph', views.personnel_graph_view, name='personnel_graph'),

    # payroll period URLs
    path('payroll-period/', views.payroll_period_list, name='payroll_period_list'),
    path('payroll-period/create/', views.create_payroll_period, name='create_payroll_period'),
    path('payroll-period/update/<int:pk>/', views.update_payroll_period, name='update_payroll_period'),
    path('payroll-period/delete/<int:pk>/', views.delete_payroll_period, name='delete_payroll_period'),

    # payroll month component
    path('payroll-month-and-component/', views.payroll_month_and_component_list, name='payroll_month_and_component_list'),
    path('payroll-month-and-component/create/', views.create_payroll_month_and_component, name='create_payroll_month_and_component'),
    path('payroll-month-and-component/update/<int:pk>/', views.update_payroll_month_and_component, name='update_payroll_month_and_component'),
    path('payroll-month-and-component/delete/<int:pk>/', views.delete_payroll_month_and_component, name='delete_payroll_month_and_component'),

    # urls.py for EarningAdjustment
    path('earning-adjustment/list', views.earning_object_list, name='earning_adjustment_list'),
    path('earning-adjustment/', views.earnings_object_detail, name='earning_adjustment_detail'),
    path('earning-adjustment/per-adjusted-month/', views.earning_per_adjusted_month, name='earning_per_adjusted_month'),
    path('earning-adjustment/monthly-earning-adjustment/', views.monthly_earning_adjustment, name='monthly_earning_adjustment'),
    path('earning-adjustment/create/', views.create_earning_adjustment, name='create_earning_adjustment'),
    path('earning-adjustment/update/<int:pk>/', views.update_earning_adjustment, name='update_earning_adjustment'),
    path('earning-adjustment/delete/<int:pk>/', views.delete_earning_adjustment, name='delete_earning_adjustment'),
    #total
    path('monthly-earning-adjustment-total', views.monthly_earning_adjustment_total, name='monthly_earning_adjustment_total'),
    # exel
    path('export/earning-adjustment-list/', views.export_earning_adjustment_list_to_excel, name='export_earning_adjustment_list_to_excel'),
    path('export/earning-per-adjusted-month/', views.export_earning_per_adjusted_month_to_excel, name='export_earning_per_adjusted_month_to_excel'),
    path('export/monthly-earning-adjustment/', views.export_monthly_earning_adjustment_to_excel, name='export_monthly_earning_adjustment_to_excel'),
    path('export_monthly_earning_adjustment_aggregate', views.export_monthly_earning_adjustment_aggregate, name='export_monthly_earning_adjustment_aggregate'),

    #

    # urls.py for DeductionAdjustment
    path('deduction-adjustment/list', views.deduction_object_list, name='deduction_adjustment_list'),
    path('deduction-adjustment/detail', views.deduction_object_detail, name='deduction_adjustment_detail'),
    path('deduction-adjustment/per-adjusted-month/', views.deduction_per_adjusted_month, name='deduction_per_adjusted_month'),
    path('deduction-adjustment/monthly-adjustment/', views.monthly_deduction_adjustment, name='monthly_deduction_adjustment'),
    path('deduction-adjustment/create/', views.create_deduction_adjustment, name='create_deduction_adjustment'),
    path('deduction-adjustment/update/<int:pk>/', views.update_deduction_adjustment, name='update_deduction_adjustment'),
    path('deduction-adjustment/delete/<int:pk>/', views.delete_deduction_adjustment, name='delete_deduction_adjustment'),
    #total
    path('monthly-deduction-adjustment-total', views.monthly_deduction_adjustment_total, name='monthly_deduction_adjustment_total'),

    #export
    path('export-deduction-adjustment/', views.export_deduction_adjustment_list_to_excel, name='export_deduction_adjustment_list_to_excel'),
    path('export-deduction-per-adjusted-month/', views.export_deduction_per_adjusted_month_to_excel, name='export_deduction_per_adjusted_month_to_excel'),
    path('export-monthly-deduction-adjustment/', views.export_monthly_deduction_adjustment_to_excel, name='export_monthly_deduction_adjustment_to_excel'),
    path('export_monthly_deduction_adjustment_aggregate', views.export_monthly_deduction_adjustment_aggregate, name='export_monthly_deduction_adjustment_aggregate'),

    #
    # combined personnel
    path('combined-personnel-detail/', views.combined_personnel_detail, name='combined_personnel_detail'),
    path('combined-personnel-adjustment/', views.combined_personnel_adjustment_list, name='combined_personnel_adjustment'),
    path('combined-personnel-payroll/', views.combined_personnel_payroll_list, name='combined_personnel_payroll'),
    path('combined-personnel-total/', views.combined_personnel_total, name='combined_personnel_total'),
    path('combined-personnel-expense/', views.combined_personnel_expense, name='combined_personnel_expense'),
    path('combined-personnel-net-income/', views.combined_personnel_net_income, name='combined_personnel_net_income'),
    path('combined-personnel-employment-income-tax/', views.combined_personnel_employment_income_tax, name='combined_personnel_employment_income_tax'),
    path('combined-personnel-pension/', views.combined_employee_pension, name='combined_personnel_pension'),
    #export
    path('export-combined-personnel-detail/', views.export_combined_personnel_detail, name='export_combined_personnel_detail'),

    path('combined-personnel-list/export/', views.export_combined_personnel_list, name='export_combined_personnel_list'),
    path('export-combined-payroll-expense/', views.export_combined_personnel_expense, name='export_combined_personnel_expense'),
    path('export-combined-net-income/', views.export_combined_personnel_net_income, name='export_combined_personnel_net_income'),
    path('export-combined-employment-tax/', views.export_combined_personnel_employment_tax, name='export_combined_personnel_employment_tax'),
    path('export-combined-personnel-pension/', views.export_combined_personnel_pension, name='export_combined_personnel_pension'),
    path('export-total-adjustment/', views.export_personnel_total_adjustment, name='export_personnel_total_adjustment'),
    path('combined-total-payroll/export/', views.export_combined_personnel_total, name='export_combined_personnel_total'),

    #combined monthly and yearly
    path('combined-monthly-detail/', views.monthly_combined_detail, name='combined_monthly_detail'),
    path('combined-monthly-summary/', views.monthly_combined_summary, name='combined_monthly_summary'),
    path('combined-monthly-adjustment-journal/', views.monthly_combined_adjustment_journal, name='monthly_combined_adjustment_journal'),
    path('monthly-combined-adjustment-summary/', views.monthly_combined_adjustment_summary, name='monthly_combined_adjustment_summary'),

    #yearly
    path('combined-yearly-detail/', views.yearly_combined_detail_view, name='combined_yearly_detail'),
    path('combined-yearly-summary/', views.yearly_combined_summary_view, name='combined_yearly_summary'),

    #export
    path('export-combined-monthly-detail/', views.export_combined_monthly_detail_to_excel, name='export_combined_monthly_detail_to_excel'),
    path('export-combined-monthly-summary/', views.export_combined_monthly_summary_to_excel, name='export_combined_monthly_summary_to_excel'),
    path('export-combined-yearly-summary/', views.export_combined_yearly_summary_to_excel,name='export_combined_yearly_summary_to_excel'),

    path('combined-yearly-detail/export/', views.export_combined_yearly_detail_to_excel,name='export_combined_yearly_detail_to_excel'),
    #graph
    path('combined-monthly-graph/', views.combined_monthly_graph_view, name='combined_monthly_graph'),

    #

    # severance pay
    path('severance-pay/', views.severance_pay_list, name='severance-pay-list'),
    path('export-severance-pay-to-excel/', views.export_severance_pay_to_excel, name='export_severance_pay_to_excel'),
    path('severance-pay/<int:pk>/', views.severance_pay_detail, name='severance_pay_detail'),
    path('severance-pay/create/', views.create_severance_pay, name='create_severance_pay'),
    path('severance-pay/update/<int:pk>/', views.update_severance_pay, name='update_severance_pay'),
    path('severance-pay/delete/<int:pk>/', views.delete_severance_pay, name='delete_severance_pay'),
    path('severance-pay-summary/', views.severance_pay_report, {'template_name': 'severance_pay/summary.html'}, name='severance_pay_summary'),
    path('severance-pay-journal-entry/', views.severance_pay_report, {'template_name': 'severance_pay/journal_entry.html'}, name='severance_pay_journal_entry'),

    # absence deduction
    path('absence-deduction', views.absence_deduction_list, name='absence_deduction_list'),
    path('absence-deduction/create/', views.create_absence_deduction, name='create_absence_deduction'),
    path('absence-deduction/update/<int:pk>/', views.update_absence_deduction, name='update_absence_deduction'),
    path('absence-deduction/delete/<int:pk>', views.delete_absence_deduction, name='delete_absence_deduction'),

    # payroll - process
    path('payroll-month/', views.payroll_month_list, name='payroll_month_list'),
    path('regular-payroll/<slug:payroll_month_slug>/', views.regular_payroll_list, name='regular_payroll_list'),
    path('regular-payroll/<slug:payroll_month_slug>/<int:regular_payroll_pk>/', views.regular_payroll_detail, name='regular_payroll_detail'),
    path('regular-payroll/<slug:payroll_month_slug>/create/', views.create_regular_payroll, name='create_regular_payroll'),
    path('regular-payroll/<slug:payroll_month_slug>/update/<int:pk>/', views.update_regular_payroll, name='update_regular_payroll'),
    path('regular-payroll/<slug:payroll_month_slug>/delete/<int:pk>/', views.delete_regular_payroll, name='delete_regular_payroll'),

    # export regular payroll
    path('export-regular-payroll/<slug:payroll_month_slug>/export-payments/', views.export_regular_payroll_to_excel, name='export_regular_payroll_to_excel'),

    # analysis
    path('payroll-month-for-summary-report/', views.payroll_month_for_summary_report, name='payroll_month_for_summary_report'),
    path('payroll-month-for-journal-entry/', views.payroll_month_for_journal_entry_report, name='payroll_month_for_journal_entry'),
    path('payroll-by-month-journal-entry-report/<slug:payroll_month_slug>/', views.regular_payroll_journal_entry, name='payroll_by_month_journal_entry'),
    path('payroll-summary-by-month/<slug:payroll_month_slug>/', views.regular_payroll_summary, name='payroll_by_month_summary_report'),
    path('payroll-summary-by-year/', views.payroll_by_year_summary_report, name='payroll_by_year_summary_report'),

    # cost share pay
    path('university-cost-sharing', views.university_cost_sharing_view, name='university_cost-sharing_list'),

    # graph
    path('regular-payroll-graph/', views.payroll_processing_graphs, name='regular_payroll_graph'),
]
