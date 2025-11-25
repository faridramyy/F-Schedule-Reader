from datetime import datetime, timedelta


def excel_date_to_string(serial):
    """Convert Excel serial date to human-readable format."""
    try:
        serial = float(serial)
        if serial > 40000:
            dt = datetime(1899, 12, 30) + timedelta(days=serial)
            return dt.strftime("%A, %B %d, %Y")
        return None
    except:
        return None


def format_hour(t):
    """Convert hour integer (10, 15, 24) to AM/PM string."""
    t = int(float(t))

    if t == 12:
        return "12:00 PM"
    if t == 24 or t == 0:
        return "12:00 AM"
    if t > 12:
        return f"{t - 12}:00 PM"

    return f"{t}:00 AM"
