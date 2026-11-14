# -*- coding: utf-8 -*-
"""把技能 scripts 目录加入 sys.path，使测试可直接 import 各模块。"""
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL / "scripts"))
