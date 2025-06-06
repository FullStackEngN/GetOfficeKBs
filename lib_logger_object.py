import logging


def create_logger_object(filename: str) -> logging.Logger:
    """
    Create and return a logger object with both console and file handlers.

    :param filename: The log file name.
    :return: Configured logger object.
    """
    logger = logging.getLogger("Download_Office_KB")
    logger.setLevel(logging.DEBUG)

    # Prevent adding handlers multiple times
    if not logger.handlers:
        # Create console handler with DEBUG level
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)

        # Create file handler with INFO level, overwrite mode
        file_handler = logging.FileHandler(filename, mode="w", encoding="utf-8")
        file_handler.setLevel(logging.INFO)

        # Create formatter and add it to the handlers
        formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%a, %d %b %Y %H:%M:%S")
        console_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)

        # Add the handlers to logger
        logger.addHandler(console_handler)
        logger.addHandler(file_handler)

    return logger
