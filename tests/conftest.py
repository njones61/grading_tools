import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1] / "plugins/grading/skills/homework-grader/scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(Path(__file__).resolve().parent))
