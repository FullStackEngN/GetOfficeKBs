from datetime import datetime


class MspFile:
    filename = ""
    product = ""
    non_security_release_date = ""
    non_security_KB = ""
    security_release_date = ""
    security_KB = ""
    security_greater_than_non_security = False

    def __init__(
        self,
        filename,
        product,
        non_security_release_date,
        non_security_KB,
        security_release_date,
        security_KB,
    ):
        self.filename = filename
        self.product = product
        self.non_security_release_date = non_security_release_date
        self.non_security_KB = non_security_KB
        self.security_release_date = security_release_date
        self.security_KB = security_KB
        self.security_greater_than_non_security = False

        if (
            self.non_security_release_date != ""
            and self.non_security_release_date != "Not applicable"
        ):
            if (
                self.security_release_date != ""
                and self.security_release_date != "Not applicable"
            ):
                try:
                    non_security_release_date_var = datetime.strptime(
                        self.non_security_release_date, "%B %d, %Y"
                    )

                except ValueError:
                    non_security_release_date_var = datetime.strptime(
                        self.non_security_release_date, "%b %d, %Y"
                    )

                try:
                    security_release_date_var = datetime.strptime(
                        self.security_release_date, "%B %d, %Y"
                    )

                except ValueError:
                    security_release_date_var = datetime.strptime(
                        self.security_release_date, "%b %d, %Y"
                    )

                if non_security_release_date_var < security_release_date_var:
                    self.security_greater_than_non_security = True
            else:
                self.security_greater_than_non_security = False
        else:
            if (
                self.security_release_date != ""
                and self.security_release_date != "Not applicable"
            ):
                self.security_greater_than_non_security = True

    def tostring(self):
        return (
            "filename: "
            + self.filename
            + "; product: "
            + self.product
            + "; non_security_release_date: "
            + self.non_security_release_date
            + "; non_security_KB: "
            + self.non_security_KB
            + "; security_release_date: "
            + self.security_release_date
            + "; security_KB: "
            + self.security_KB
            + "; security_greater_than_non_security: "
            + str(self.security_greater_than_non_security)
        )
