# -*- coding: utf-8 -*-
"""Delivery number generation.

Format: {customer_prefix}{YY}{MM}{3-digit-seq}
Example: HR202605001

Sequence is per-customer per-month, independent of order creation.
Only assigned when generating a delivery note.
"""
from datetime import datetime
from db.database import get_connection


def generate_delivery_number(customer_id, delivery_date=None):
    """Generate the next delivery number for a customer.

    Args:
        customer_id: customer ID
        delivery_date: date string "YYYY-MM-DD", defaults to today

    Returns:
        str: delivery number like "HR202605001", or None on error
    """
    if delivery_date is None:
        today = datetime.now()
    else:
        parts = delivery_date.split("-")
        today = datetime(int(parts[0]), int(parts[1]), int(parts[2]))

    yyyy = today.strftime("%Y")
    mm = today.strftime("%m")
    prefix = _get_customer_prefix(customer_id)
    if not prefix:
        return None
    like_pattern = "{}{}{}%".format(prefix, yyyy, mm)

    c = get_connection()
    try:
        row = c.execute(
            "SELECT delivery_number FROM delivery_note"
            " WHERE delivery_number LIKE ?"
            " ORDER BY delivery_number DESC LIMIT 1",
            (like_pattern,)).fetchone()

        if row:
            last_seq = int(row["delivery_number"][-3:])
            seq = last_seq + 1
        else:
            seq = 1

        return "{}{}{}{:03d}".format(prefix, yyyy, mm, seq)
    finally:
        c.close()


def _get_customer_prefix(customer_id):
    c = get_connection()
    try:
        row = c.execute(
            "SELECT prefix FROM customer WHERE id=?",
            (customer_id,)).fetchone()
        return row["prefix"] if row else None
    finally:
        c.close()
