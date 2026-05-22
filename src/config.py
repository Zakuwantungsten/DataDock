import json
from typing import Dict, List


COLUMN_ALIASES_JSON = """
{
    "Position": ["Current Position", "Location", "Position"],
    "Truck": ["Truck LP", "Truck ID", "Truck"],
    "Trailer": ["Trailer"],
    "Status": ["Status", "Remarks"],
    "Departure Date": ["Depart Date", "Date of Departure", "Dept Date"]
}
"""


def get_column_aliases() -> Dict[str, List[str]]:
    return json.loads(COLUMN_ALIASES_JSON)
    
