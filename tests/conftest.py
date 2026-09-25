import sys
from pathlib import Path

# Let tests import xlgit (repo root) and each other's helpers (tests/).
sys.path[:0] = [str(Path(__file__).resolve().parents[1]), str(Path(__file__).resolve().parent)]
