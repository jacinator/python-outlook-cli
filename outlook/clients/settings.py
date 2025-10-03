from pathlib import Path
from typing import Final

ROOT: Final[Path] = Path(__file__).parent.parent.parent
AUTH: Final[Path] = ROOT / ".auth.json"
AUTH_RECORD: Final[Path] = ROOT / ".auth_record.json"
