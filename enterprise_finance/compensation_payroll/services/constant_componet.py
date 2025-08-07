

FULLY_TAXABLE = [
        'basic_salary', 'overtime', 'housing_allowance', 'position_allowance',
        'commission', 'telephone_allowance', 'one_time_bonus', 'casual_labor_wage'
    ]

PARTIALLY_TAXABLE = [
    'transport_home_to_office', 'transport_for_work', 'fuel_home_to_office',
    'fuel_for_work', 'per_diem', 'hardship_allowance'
]

NON_TAXABLE = [
    'public_cash_award', 'incidental_operation_allowance', 'medical_allowance',
    'cash_gift', 'personal_injury', 'child_support_payment', 'tuition_fees'
]

DEFERRED_EARNINGS = [
    'leave_encashment', 'quarterly_bonus',
    'semi_annual_bonus', 'annual_bonus', 'performance_based_bonus',
    'project_completion_bonuses', 'holiday_bonus', 'other_bonus'
]

DEDUCTIONS = [
    'charitable_donation', 'saving_plan', 'loan_payment', 'court_order',
    'workers_association', 'personnel_insurance_saving',
    'university_cost_share_pay', 'red_cross', 'party_contribution', 'other_deduction'
]

PENSIONABLE = [
    'basic_salary',
]
