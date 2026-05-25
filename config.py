# -*- coding: utf-8 -*-
"""Application configuration helpers."""
from db.models import Settings


def get_company_info():
    """Return a dict with company header info for delivery notes."""
    return {
        "company_name": Settings.get("company_name"),
        "company_address": Settings.get("company_address"),
        "company_email": Settings.get("company_email"),
        "footer_text": Settings.get("footer_text"),
        "maker_name": Settings.get("maker_name"),
    }


def get_default_excel_folder():
    return Settings.get("default_excel_folder")


def get_app_version():
    return Settings.get("app_version", "1.0.0")
