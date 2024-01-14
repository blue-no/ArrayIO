from __future__ import annotations

from pathlib import Path

VALUES = [[100 * (i + j + 1) for i in range(10)] for j in range(20000)]
CELLS = ("B3", "K20002")
LINES = (3, 20002)
DIR = Path("./tests")
