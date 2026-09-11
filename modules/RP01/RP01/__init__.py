print("LOADED:", __file__)

from flask import Blueprint

MODULE_CODE = "RP01"
MODULE_INFO = {
    "code": MODULE_CODE,
    "name": "Reports"
}

bp = Blueprint(
    "RP01",
    __name__,
    template_folder="."
)

# Import after blueprint creation
from . import views
from . import vessel_call_report
from . import service_record_report
from .JJLTPL import jjltpl
from .report1 import report1
from .report2 import report2
from .report3 import report3 as report_03_views
from .report7 import report7 as report_07_views
from .report_06 import views as report_06_views
from .Berth_plan import view as berth_plan_view
from .report4 import report4

from .custom_report import views as custom_report_views
from .report_08 import report8 as report_08_views
from .report9 import report09 as report_09_views
from .report13 import report13 as report_13_views
from.report11 import report11 as report_11_views 
from .report12 import report12 as report_12_views
from .report5 import report5 as report_05_views
from .report10 import report10 as report_10_views
from .report_budget import report_budget as report_budget_views
from .dpr import dpr as dpr_views
from .vessel_timing import vessel_timing as vessel_timing_views
from .vessel_delay import vessel_delay as vessel_delay_views


