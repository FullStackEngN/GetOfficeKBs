from datetime import datetime


class msp_file_item:
    def __init__(
        self,
        filename: str,
        product: str,
        non_security_release_date: str,
        non_security_KB: str,
        security_release_date: str,
        security_KB: str,
    ):
        self.filename = filename
        self.product = product
        self.non_security_release_date = non_security_release_date
        self.non_security_KB = non_security_KB
        self.security_release_date = security_release_date
        self.security_KB = security_KB
        self.security_greater_than_non_security = self._compare_release_dates()

    def _parse_date(self, date_str: str):
        """Try to parse a date string with multiple formats."""
        for fmt in ("%B %d, %Y", "%b %d, %Y"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return None

    def _compare_release_dates(self) -> bool:
        """Compare release dates and determine if security date is greater."""
        ns_date = self.non_security_release_date
        s_date = self.security_release_date

        # Check for empty or 'Not applicable' values
        if ns_date and ns_date != "Not applicable":
            if s_date and s_date != "Not applicable":
                ns_dt = self._parse_date(ns_date)
                s_dt = self._parse_date(s_date)
                if ns_dt and s_dt:
                    return ns_dt < s_dt
                return False
            else:
                return False
        else:
            if s_date and s_date != "Not applicable":
                return True
            return False

    def tostring(self) -> str:
        """Return a string representation of the object."""
        return (
            f"filename: {self.filename}; "
            f"product: {self.product}; "
            f"non_security_release_date: {self.non_security_release_date}; "
            f"non_security_KB: {self.non_security_KB}; "
            f"security_release_date: {self.security_release_date}; "
            f"security_KB: {self.security_KB}; "
            f"security_greater_than_non_security: {self.security_greater_than_non_security}"
        )
