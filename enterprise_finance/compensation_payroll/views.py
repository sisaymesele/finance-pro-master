# payroll process is on last
from django.shortcuts import render, get_object_or_404, redirect
from .models import *
from django.core.paginator import Paginator
from .forms import *

from django.urls import reverse
from django.contrib import messages
from django.db.models import Sum, Count, F, Case, When, Value, DecimalField, IntegerField
# graph
import plotly.graph_objs as go
# export
import openpyxl
from io import BytesIO
import io
from decimal import Decimal
from django.http import HttpResponse
from openpyxl import Workbook
#
from django.db.models import Q
from datetime import datetime
# total
from collections import defaultdict
from django.contrib.auth.decorators import login_required



#
#