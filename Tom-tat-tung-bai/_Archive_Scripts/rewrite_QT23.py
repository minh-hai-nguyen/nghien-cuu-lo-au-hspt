# -*- coding: utf-8 -*-
"""QT23 JAACAP — viết lại với 3 bảng từ PDF gốc"""
import sys, io
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
# Copy code from rewrite_QT21_24.py QT23 section but add 2 more tables
# TODO: Run in next session - context near limit
# Key data from PDF:
# Table 2: Anxiety 9.6%->19.2% AOR=2.17, Depression 13.4%->17.0% AOR=1.20
# Table 4: Sex trends
# Add: cross-study comparison table + screening vs diagnostic table
sys.stderr.write('QT23 TODO - run in next session\n')
