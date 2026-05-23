import json
import re
from pathlib import Path
from typing import Dict, List, Optional


COLUMN_ALIASES_JSON = """
{
    "Position": [
        "Current Position",
        "Location",
        "Position"
    ],
    "Truck": [
        "Truck LP",
        "Truck ID",
        "Truck"
    ],
    "Trailer": [
        "Trailer"
    ],
    "Status": [
        "Status",
        "Remarks"
    ],
    "Departure Date": [
        "Depart Date",
        "Date of Departure",
        "Dept Date"
    ]
}
"""


def get_column_aliases() -> Dict[str, List[str]]:
    return json.loads(COLUMN_ALIASES_JSON)


def save_column_aliases(
    aliases_map: Dict[str, List[str]],
    config_path: Optional[Path] = None,
) -> None:
    if config_path is None:
        config_path = Path(__file__).resolve()

    json_text = json.dumps(aliases_map, indent=4, sort_keys=False)
    replacement = f'COLUMN_ALIASES_JSON = """\n{json_text}\n"""'

    text = config_path.read_text(encoding="utf-8")
    pattern = r'COLUMN_ALIASES_JSON\s*=\s*"""[\s\S]*?"""'
    if not re.search(pattern, text):
        raise ValueError("COLUMN_ALIASES_JSON block not found in config.py")

    updated = re.sub(pattern, replacement, text, count=1)
    config_path.write_text(updated, encoding="utf-8")
    
