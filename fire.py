"""
fire.py — thin wrapper around template/pdf.py
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'template'))

from pdf import build_resume


def generate_resume_pdf(data, main_color="#1a56db", secondary_color="#6b7280", font="Helvetica"):
    return build_resume(data, main_color=main_color, secondary_color=secondary_color, font=font)