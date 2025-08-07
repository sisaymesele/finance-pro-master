import datetime
from django.db.models import Q

current_year = datetime.datetime.now().year
YEAR_CHOICES = [(str(year), str(year)) for year in range(2017, 2041)]

MONTH_CHOICES = [
    ('01', 'September'), ('02', 'October'), ('03', 'November'), ('04', 'December'),
    ('05', 'January'), ('06', 'February'), ('07', 'March'), ('08', 'April'),
    ('09', 'May'), ('10', 'June'), ('11', 'July'), ('12', 'August'),
]

