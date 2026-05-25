# -*- coding: utf-8 -*-
"""ORM-style model classes for all database tables."""
from datetime import datetime
from .database import get_connection


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ============================================================
# Customer
# ============================================================
class Customer(object):

    FIELDS = [
        "name", "prefix", "contact_person", "phone", "email",
        "address", "notes", "has_model_number", "has_item_code",
        "excel_file_path", "print_template", "dn_headers"
    ]

    @staticmethod
    def create(name, prefix, **kwargs):
        c = get_connection()
        now = _now()
        vals = {
            "name": name, "prefix": prefix,
            "contact_person": "", "phone": "", "email": "",
            "address": "", "notes": "",
            "has_model_number": 1, "has_item_code": 1,
            "excel_file_path": "", "print_template": "",
            "dn_headers": ",订单号,,款号",
            "created_at": now, "updated_at": now
        }
        vals.update(kwargs)
        cols = [k for k in vals if k != "created_at" and k != "updated_at"]
        cols += ["created_at", "updated_at"]
        placeholders = ",".join(["?"] * len(cols))
        sql = "INSERT INTO customer ({}) VALUES ({})".format(
            ",".join(cols), placeholders)
        params = [vals[k] for k in cols]
        try:
            cur = c.execute(sql, params)
            c.commit()
            return cur.lastrowid
        finally:
            c.close()

    @staticmethod
    def get_all():
        c = get_connection()
        try:
            rows = c.execute("SELECT * FROM customer ORDER BY name").fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def get_by_id(customer_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT * FROM customer WHERE id=?", (customer_id,)).fetchone()
            return dict(row) if row else None
        finally:
            c.close()

    @staticmethod
    def update(customer_id, **kwargs):
        allowed = set(Customer.FIELDS)
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        sets.append("updated_at = ?")
        vals.append(_now())
        vals.append(customer_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE customer SET {} WHERE id=?".format(", ".join(sets)),
                vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(customer_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM customer WHERE id=?", (customer_id,))
            c.commit()
        finally:
            c.close()


# ============================================================
# Product
# ============================================================
class Product(object):

    @staticmethod
    def create(model_number="", item_code="", product_name="",
               specification=""):
        c = get_connection()
        try:
            cur = c.execute(
                "INSERT INTO product (model_number, item_code, product_name,"
                " specification) VALUES (?,?,?,?)",
                (model_number, item_code, product_name, specification))
            c.commit()
            return cur.lastrowid
        finally:
            c.close()

    @staticmethod
    def get_all():
        c = get_connection()
        try:
            rows = c.execute(
                "SELECT * FROM product ORDER BY product_name").fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def get_by_id(product_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT * FROM product WHERE id=?", (product_id,)).fetchone()
            return dict(row) if row else None
        finally:
            c.close()

    @staticmethod
    def search(keyword):
        c = get_connection()
        kw = "%{}%".format(keyword)
        try:
            rows = c.execute(
                "SELECT * FROM product WHERE product_name LIKE ?"
                " OR model_number LIKE ? OR item_code LIKE ?"
                " ORDER BY product_name",
                (kw, kw, kw)).fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def find_or_create(model_number, item_code, product_name):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT * FROM product WHERE product_name=?"
                " AND model_number=? AND item_code=?",
                (product_name, model_number, item_code)).fetchone()
            if row:
                return dict(row)
            cur = c.execute(
                "INSERT INTO product (model_number, item_code, product_name)"
                " VALUES (?,?,?)",
                (model_number, item_code, product_name))
            c.commit()
            pid = cur.lastrowid
            return Product.get_by_id(pid)
        finally:
            c.close()

    @staticmethod
    def update(product_id, **kwargs):
        allowed = {"model_number", "item_code", "product_name", "specification"}
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        vals.append(product_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE product SET {} WHERE id=?".format(", ".join(sets)),
                vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(product_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM product WHERE id=?", (product_id,))
            c.commit()
        finally:
            c.close()


# ============================================================
# Color
# ============================================================
class Color(object):

    @staticmethod
    def get_all():
        c = get_connection()
        try:
            rows = c.execute(
                "SELECT * FROM color ORDER BY name").fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def get_by_id(color_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT * FROM color WHERE id=?", (color_id,)).fetchone()
            return dict(row) if row else None
        finally:
            c.close()

    @staticmethod
    def find_or_create(name):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT * FROM color WHERE name=?", (name,)).fetchone()
            if row:
                return dict(row)
            cur = c.execute(
                "INSERT INTO color (name) VALUES (?)", (name,))
            c.commit()
            pid = cur.lastrowid
            return Color.get_by_id(pid)
        finally:
            c.close()

    @staticmethod
    def delete(color_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM color WHERE id=?", (color_id,))
            c.commit()
        finally:
            c.close()


# ============================================================
# Order
# ============================================================
class Order(object):

    @staticmethod
    def create(customer_id, display_name, order_date, notes=""):
        c = get_connection()
        now = _now()
        try:
            cur = c.execute(
                "INSERT INTO orders (customer_id, display_name, order_date,"
                " notes, status, created_at, updated_at)"
                " VALUES (?,?,?,?,'draft',?,?)",
                (customer_id, display_name, order_date, notes, now, now))
            c.commit()
            return cur.lastrowid
        finally:
            c.close()

    @staticmethod
    def get_all(page=1, page_size=50, filters=None, order_by=None):
        """Paginated order list with optional filters."""
        c = get_connection()
        try:
            query = (
                "SELECT o.*, c.name as customer_name, c.prefix as customer_prefix"
                " FROM orders o JOIN customer c ON o.customer_id = c.id"
                " WHERE 1=1"
            )
            params = []
            if filters:
                if filters.get("customer_id"):
                    query += " AND o.customer_id = ?"
                    params.append(filters["customer_id"])
                if filters.get("customer_ids"):
                    placeholders = ",".join(["?"] * len(filters["customer_ids"]))
                    query += " AND o.customer_id IN ({})".format(placeholders)
                    params.extend(filters["customer_ids"])
                if filters.get("status"):
                    query += " AND o.status = ?"
                    params.append(filters["status"])
                if filters.get("status_in"):
                    placeholders = ",".join(["?"] * len(filters["status_in"]))
                    query += " AND o.status IN ({})".format(placeholders)
                    params.extend(filters["status_in"])
                if filters.get("status_not_in"):
                    placeholders = ",".join(
                        ["?"] * len(filters["status_not_in"]))
                    query += " AND o.status NOT IN ({})".format(placeholders)
                    params.extend(filters["status_not_in"])
                if filters.get("date_from"):
                    query += " AND o.order_date >= ?"
                    params.append(filters["date_from"])
                if filters.get("date_to"):
                    query += " AND o.order_date <= ?"
                    params.append(filters["date_to"])
                if filters.get("keyword"):
                    kw = "%{}%".format(filters["keyword"])
                    query += (" AND (o.display_name LIKE ?"
                              " OR o.delivery_number LIKE ?"
                              " OR c.name LIKE ?)")
                    params.extend([kw, kw, kw])

            # count total
            count_sql = query.replace(
                "SELECT o.*, c.name as customer_name, c.prefix as customer_prefix",
                "SELECT COUNT(*)")
            total = c.execute(count_sql, params).fetchone()[0]

            # pagination
            query += " ORDER BY {}".format(
                order_by if order_by else "o.created_at DESC")
            offset = (page - 1) * page_size
            query += " LIMIT ? OFFSET ?"
            params.extend([page_size, offset])

            rows = c.execute(query, params).fetchall()
            return {
                "items": [dict(r) for r in rows],
                "total": total,
                "page": page,
                "page_size": page_size,
                "total_pages": max(1, (total + page_size - 1) // page_size)
            }
        finally:
            c.close()

    @staticmethod
    def get_by_id(order_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT o.*, c.name as customer_name, c.prefix as customer_prefix"
                " FROM orders o JOIN customer c ON o.customer_id = c.id"
                " WHERE o.id=?", (order_id,)).fetchone()
            return dict(row) if row else None
        finally:
            c.close()

    @staticmethod
    def get_by_display_name(display_name):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT o.*, c.name as customer_name, c.prefix as customer_prefix"
                " FROM orders o JOIN customer c ON o.customer_id = c.id"
                " WHERE o.display_name=?", (display_name,)).fetchone()
            return dict(row) if row else None
        finally:
            c.close()

    @staticmethod
    def update(order_id, **kwargs):
        allowed = {"display_name", "order_date", "delivery_number",
                   "delivery_date", "status", "notes"}
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        sets.append("updated_at = ?")
        vals.append(_now())
        vals.append(order_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE orders SET {} WHERE id=?".format(", ".join(sets)),
                vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(order_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM orders WHERE id=?", (order_id,))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def get_item_count(order_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT COUNT(*) as cnt FROM order_item"
                " WHERE order_id=?", (order_id,)).fetchone()
            return row["cnt"] if row else 0
        finally:
            c.close()

    @staticmethod
    def next_display_name(customer_id, order_date):
        """Generate display name like '宏润5.4'.
        If same customer+date already exists, append (2), (3)...
        """
        c = get_connection()
        try:
            row = c.execute(
                "SELECT name FROM customer WHERE id=?",
                (customer_id,)).fetchone()
            if not row:
                return None
            customer_name = row["name"]
            parts = order_date.split("-")
            if len(parts) >= 2:
                month_day = "{}.{}".format(
                    int(parts[1]), int(parts[2]))
            else:
                month_day = order_date
            base = "{}{}".format(customer_name, month_day)

            rows = c.execute(
                "SELECT display_name FROM orders"
                " WHERE customer_id=? AND display_name LIKE ?"
                " ORDER BY display_name DESC",
                (customer_id, "{}%".format(base))).fetchall()
            existing = set(r["display_name"] for r in rows)

            if base not in existing:
                return base
            n = 2
            while True:
                candidate = "{}({})".format(base, n)
                if candidate not in existing:
                    return candidate
                n += 1
        finally:
            c.close()


# ============================================================
# OrderItem
# ============================================================
class OrderItem(object):

    @staticmethod
    def create(order_id, product_id, color_id, model_number, item_code,
               product_name, color_name, quantity_formula="", quantity=0,
               print_count=1, is_printed=0, push_count=0,
               unit_price="", amount=0, notes="", sort_order=0,
               order_number="", mfg_number="", customer_code=""):
        c = get_connection()
        now = _now()
        try:
            cur = c.execute(
                "INSERT INTO order_item (order_id, product_id, color_id,"
                " model_number, item_code, customer_code,"
                " product_name, color_name,"
                " quantity_formula, quantity, print_count, is_printed,"
                " push_count, unit_price, amount, notes, sort_order,"
                " order_number, mfg_number,"
                " created_at, updated_at)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (order_id, product_id, color_id, model_number, item_code,
                 customer_code,
                 product_name, color_name, quantity_formula, quantity,
                 print_count, is_printed, push_count,
                 unit_price, amount, notes, sort_order,
                 order_number, mfg_number,
                 now, now))
            c.commit()
            return cur.lastrowid
        finally:
            c.close()

    @staticmethod
    def get_by_order(order_id):
        c = get_connection()
        try:
            rows = c.execute(
                "SELECT * FROM order_item WHERE order_id=?"
                " ORDER BY sort_order, id",
                (order_id,)).fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def update(item_id, **kwargs):
        allowed = {
            "model_number", "item_code", "product_name", "color_name",
            "quantity_formula", "quantity", "print_count", "is_printed",
            "push_count", "unit_price", "amount", "notes", "sort_order",
            "order_number", "mfg_number", "customer_code"
        }
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        sets.append("updated_at = ?")
        vals.append(_now())
        vals.append(item_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE order_item SET {} WHERE id=?".format(", ".join(sets)),
                vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(item_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM order_item WHERE id=?", (item_id,))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete_by_order(order_id):
        c = get_connection()
        try:
            c.execute(
                "DELETE FROM order_item WHERE order_id=?", (order_id,))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def reorder(order_id, item_ids):
        c = get_connection()
        try:
            for i, item_id in enumerate(item_ids):
                c.execute(
                    "UPDATE order_item SET sort_order=?"
                    " WHERE id=? AND order_id=?",
                    (i, item_id, order_id))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def mark_printed(item_ids):
        c = get_connection()
        try:
            for item_id in item_ids:
                c.execute(
                    "UPDATE order_item SET is_printed=1, updated_at=?"
                    " WHERE id=?",
                    (_now(), item_id))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def mark_unprinted(item_ids):
        c = get_connection()
        try:
            for item_id in item_ids:
                c.execute(
                    "UPDATE order_item SET is_printed=0, updated_at=?"
                    " WHERE id=?",
                    (_now(), item_id))
            c.commit()
        finally:
            c.close()


# ============================================================
# Settings
# ============================================================
class DeliveryNote(object):

    @staticmethod
    def create(order_id, delivery_number, delivery_date="", notes="",
               customer_id=None):
        c = get_connection()
        now = _now()
        try:
            # Resolve customer_id from order if not provided
            if customer_id is None and order_id is not None:
                row = c.execute(
                    "SELECT customer_id FROM orders WHERE id=?",
                    (order_id,)).fetchone()
                if row:
                    customer_id = row["customer_id"]
            cur = c.execute(
                "INSERT INTO delivery_note (order_id, customer_id,"
                " delivery_number, delivery_date, notes, status,"
                " created_at, updated_at)"
                " VALUES (?,?,?,?,?,'draft',?,?)",
                (order_id, customer_id, delivery_number, delivery_date,
                 notes, now, now))
            c.commit()
            dn_id = cur.lastrowid
            # If order_id provided, insert into junction table
            if order_id is not None:
                c.execute(
                    "INSERT OR IGNORE INTO delivery_note_order"
                    " (delivery_note_id, order_id) VALUES (?,?)",
                    (dn_id, order_id))
                c.commit()
            return dn_id
        finally:
            c.close()

    @staticmethod
    def get_by_id(dn_id):
        c = get_connection()
        try:
            row = c.execute(
                "SELECT dn.*,"
                " c.name as customer_name, c.id as customer_id"
                " FROM delivery_note dn"
                " LEFT JOIN customer c ON dn.customer_id = c.id"
                " WHERE dn.id=?", (dn_id,)).fetchone()
            if not row:
                return None
            result = dict(row)
            # Get source orders
            orders = c.execute(
                "SELECT o.id, o.display_name, o.order_date"
                " FROM delivery_note_order dno"
                " JOIN orders o ON dno.order_id = o.id"
                " WHERE dno.delivery_note_id=?"
                " ORDER BY o.order_date, o.id",
                (dn_id,)).fetchall()
            result["source_orders"] = [dict(o) for o in orders]
            # Backward compat: if single source order
            if len(result["source_orders"]) == 1:
                o = result["source_orders"][0]
                result["order_display_name"] = o["display_name"]
                result["order_date"] = o["order_date"]
                result["order_id"] = result.get("order_id") or o["id"]
            elif result["source_orders"]:
                names = [o["display_name"] for o in result["source_orders"]]
                result["order_display_name"] = ",".join(names)
                result["order_date"] = result["source_orders"][0]["order_date"]
            return result
        finally:
            c.close()

    @staticmethod
    def get_by_order(order_id, page=1, page_size=50):
        c = get_connection()
        try:
            count_row = c.execute(
                "SELECT COUNT(*) FROM delivery_note_order"
                " WHERE order_id=?", (order_id,)).fetchone()
            total = count_row[0] if count_row else 0
            offset = (page - 1) * page_size
            rows = c.execute(
                "SELECT dn.*,"
                " c.name as customer_name"
                " FROM delivery_note dn"
                " JOIN delivery_note_order dno ON dn.id = dno.delivery_note_id"
                " JOIN customer c ON dn.customer_id = c.id"
                " WHERE dno.order_id=?"
                " ORDER BY dn.created_at DESC"
                " LIMIT ? OFFSET ?",
                (order_id, page_size, offset)).fetchall()
            return {
                "items": [dict(r) for r in rows],
                "total": total,
                "page": page,
                "page_size": page_size,
                "total_pages": max(1, (total + page_size - 1) // page_size)
            }
        finally:
            c.close()

    @staticmethod
    def get_all(page=1, page_size=50, filters=None, order_by=None):
        c = get_connection()
        try:
            query = (
                "SELECT dn.*,"
                " c.name as customer_name, c.id as customer_id,"
                " o.display_name as order_display_name"
                " FROM delivery_note dn"
                " LEFT JOIN customer c ON dn.customer_id = c.id"
                " LEFT JOIN orders o ON dn.order_id = o.id"
                " WHERE 1=1"
            )
            params = []
            if filters:
                if filters.get("customer_id"):
                    query += " AND dn.customer_id = ?"
                    params.append(filters["customer_id"])
                if filters.get("customer_ids"):
                    ph = ",".join(["?"] * len(filters["customer_ids"]))
                    query += " AND dn.customer_id IN ({})".format(ph)
                    params.extend(filters["customer_ids"])
                if filters.get("status"):
                    query += " AND dn.status = ?"
                    params.append(filters["status"])
                if filters.get("status_in"):
                    ph = ",".join(["?"] * len(filters["status_in"]))
                    query += " AND dn.status IN ({})".format(ph)
                    params.extend(filters["status_in"])
                if filters.get("status_not_in"):
                    ph = ",".join(["?"] * len(filters["status_not_in"]))
                    query += " AND dn.status NOT IN ({})".format(ph)
                    params.extend(filters["status_not_in"])
                if filters.get("date_from"):
                    query += " AND dn.delivery_date >= ?"
                    params.append(filters["date_from"])
                if filters.get("date_to"):
                    query += " AND dn.delivery_date <= ?"
                    params.append(filters["date_to"])
                if filters.get("keyword"):
                    kw = "%{}%".format(filters["keyword"])
                    query += (" AND (dn.delivery_number LIKE ?"
                              " OR c.name LIKE ?)")
                    params.extend([kw, kw])

            count_sql = query.replace(
                "SELECT dn.*,"
                " c.name as customer_name, c.id as customer_id,"
                " o.display_name as order_display_name",
                "SELECT COUNT(*)")
            total = c.execute(count_sql, params).fetchone()[0]

            query += " ORDER BY {}".format(
                order_by if order_by else "dn.created_at DESC")
            offset = (page - 1) * page_size
            query += " LIMIT ? OFFSET ?"
            params.extend([page_size, offset])

            rows = c.execute(query, params).fetchall()
            items = [dict(r) for r in rows]
            # Enrich multi-order DNs: if order_display_name is NULL,
            # fetch from junction table
            null_ids = [it["id"] for it in items
                       if not it.get("order_display_name")]
            if null_ids:
                ph = ",".join(["?"] * len(null_ids))
                orows = c.execute(
                    "SELECT dno.delivery_note_id, o.display_name"
                    " FROM delivery_note_order dno"
                    " JOIN orders o ON dno.order_id = o.id"
                    " WHERE dno.delivery_note_id IN ({})"
                    " ORDER BY dno.delivery_note_id, o.order_date".format(ph),
                    null_ids).fetchall()
                omap = {}
                for r in orows:
                    omap.setdefault(r["delivery_note_id"], []).append(
                        r["display_name"])
                for it in items:
                    if not it.get("order_display_name"):
                        names = omap.get(it["id"], [])
                        it["order_display_name"] = (
                            ",".join(names) if names else "")
            return {
                "items": items,
                "total": total,
                "page": page,
                "page_size": page_size,
                "total_pages": max(1, (total + page_size - 1) // page_size)
            }
        finally:
            c.close()

    @staticmethod
    def update(dn_id, **kwargs):
        allowed = {"delivery_number", "delivery_date", "status", "notes",
                   "export_count"}
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        sets.append("updated_at = ?")
        vals.append(_now())
        vals.append(dn_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE delivery_note SET {} WHERE id=?".format(
                    ", ".join(sets)), vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(dn_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM delivery_note WHERE id=?", (dn_id,))
            c.commit()
        finally:
            c.close()


class DeliveryNoteOrder(object):

    @staticmethod
    def create(delivery_note_id, order_id):
        c = get_connection()
        try:
            c.execute(
                "INSERT OR IGNORE INTO delivery_note_order"
                " (delivery_note_id, order_id) VALUES (?,?)",
                (delivery_note_id, order_id))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def get_by_dn(delivery_note_id):
        c = get_connection()
        try:
            rows = c.execute(
                "SELECT o.id, o.display_name, o.order_date, o.customer_id"
                " FROM delivery_note_order dno"
                " JOIN orders o ON dno.order_id = o.id"
                " WHERE dno.delivery_note_id=?"
                " ORDER BY o.order_date, o.id",
                (delivery_note_id,)).fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def delete_by_dn(delivery_note_id):
        c = get_connection()
        try:
            c.execute(
                "DELETE FROM delivery_note_order"
                " WHERE delivery_note_id=?", (delivery_note_id,))
            c.commit()
        finally:
            c.close()


class DNItem(object):

    @staticmethod
    def create(delivery_note_id, model_number, item_code, product_name,
               color_name, quantity_formula="", quantity=0, unit_price="",
               amount=0, notes="", sort_order=0, order_item_id=None,
               mfg_number="", customer_code=""):
        c = get_connection()
        now = _now()
        try:
            cur = c.execute(
                "INSERT INTO delivery_note_item (delivery_note_id,"
                " order_item_id, model_number, item_code, customer_code,"
                " product_name,"
                " color_name, quantity_formula, quantity, unit_price,"
                " amount, notes, sort_order, mfg_number,"
                " created_at, updated_at)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (delivery_note_id, order_item_id, model_number, item_code,
                 customer_code,
                 product_name, color_name, quantity_formula, quantity,
                 unit_price, amount, notes, sort_order, mfg_number,
                 now, now))
            c.commit()
            return cur.lastrowid
        finally:
            c.close()

    @staticmethod
    def get_by_dn(delivery_note_id):
        c = get_connection()
        try:
            rows = c.execute(
                "SELECT * FROM delivery_note_item"
                " WHERE delivery_note_id=? ORDER BY sort_order, id",
                (delivery_note_id,)).fetchall()
            return [dict(r) for r in rows]
        finally:
            c.close()

    @staticmethod
    def update(item_id, **kwargs):
        allowed = {"model_number", "item_code", "customer_code",
                   "product_name", "color_name",
                   "quantity_formula", "quantity", "unit_price", "amount",
                   "notes", "sort_order", "mfg_number"}
        sets = []
        vals = []
        for k, v in kwargs.items():
            if k in allowed:
                sets.append("{} = ?".format(k))
                vals.append(v)
        if not sets:
            return
        sets.append("updated_at = ?")
        vals.append(_now())
        vals.append(item_id)
        c = get_connection()
        try:
            c.execute(
                "UPDATE delivery_note_item SET {} WHERE id=?".format(
                    ", ".join(sets)), vals)
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete(item_id):
        c = get_connection()
        try:
            c.execute("DELETE FROM delivery_note_item WHERE id=?", (item_id,))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def delete_by_dn(dn_id):
        c = get_connection()
        try:
            c.execute(
                "DELETE FROM delivery_note_item"
                " WHERE delivery_note_id=?", (dn_id,))
            c.commit()
        finally:
            c.close()


def deduplicate_dn_items(items):
    """Clear 客户号/订单号/制单号/款号 on rows where all four match the group leader.

    Compared fields: customer_code, model_number, mfg_number, item_code.
    Modifies items in place.  The first of each consecutive same-group keeps its
    values; subsequent rows whose four fields all match the *last non-cleared row*
    above are cleared.
    """
    for i in range(1, len(items)):
        cur = items[i]
        prev = None
        for j in range(i - 1, -1, -1):
            p = items[j]
            if (p.get("customer_code") or p.get("model_number")
                    or p.get("mfg_number") or p.get("item_code")):
                prev = p
                break
        if prev is None:
            continue
        if (str(cur.get("customer_code", "")).strip()
                == str(prev.get("customer_code", "")).strip()
                and str(cur.get("model_number", "")).strip()
                == str(prev.get("model_number", "")).strip()
                and str(cur.get("mfg_number", "")).strip()
                == str(prev.get("mfg_number", "")).strip()
                and str(cur.get("item_code", "")).strip()
                == str(prev.get("item_code", "")).strip()):
            cur["customer_code"] = ""
            cur["model_number"] = ""
            cur["mfg_number"] = ""
            cur["item_code"] = ""
    return items


class Settings(object):

    DEFAULTS = {
        "company_name": u"佛山市和诚内衣辅料有限公司",
        "company_address": (
            u"地址：盐步联安路97号F座    电话：18988539567"
            u"        传真：0757-81101784"),
        "company_email": u"邮箱：ZWQ197819@126.COM",
        "footer_text": (
            u"客户须知：如对产品质量有异议，请于收到货后三天内"
            u"与本公司联系解决，逾期则视为合格。谢谢合作！"),
        "maker_name": u"制表人13829133080",
        "default_excel_folder": "D:/xxm/送货单",
        "template_dir": "D:/xxm/templates",
        "printer_name": "",
        "app_version": "1.0.0",
    }

    @staticmethod
    def get(key, default=None):
        if default is None:
            default = Settings.DEFAULTS.get(key, "")
        c = get_connection()
        try:
            row = c.execute(
                "SELECT value FROM settings WHERE key=?", (key,)).fetchone()
            return row["value"] if row else default
        finally:
            c.close()

    @staticmethod
    def set(key, value):
        c = get_connection()
        try:
            c.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?,?)",
                (key, str(value)))
            c.commit()
        finally:
            c.close()

    @staticmethod
    def get_all():
        c = get_connection()
        try:
            rows = c.execute("SELECT * FROM settings").fetchall()
            result = dict(Settings.DEFAULTS)
            for r in rows:
                result[r["key"]] = r["value"]
            return result
        finally:
            c.close()
