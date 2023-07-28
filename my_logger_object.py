import logging


def create_logger_object(FILENAME):
    logger = logging.getLogger("Download_Office_KB")
    logger.setLevel(logging.DEBUG)

    # create console handler with a higher log level
    consoleHandler = logging.StreamHandler()
    consoleHandler.setLevel(logging.DEBUG)

    # create file handler which logs even debug messages
    fileHandler = logging.FileHandler(FILENAME, mode="w")
    fileHandler.setLevel(logging.INFO)

    # create formatter and add it to the handlers
    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s %(message)s", datefmt="%a, %d %b %Y %H:%M:%S"
    )

    consoleHandler.setFormatter(formatter)
    fileHandler.setFormatter(formatter)

    # add the handlers to logger
    logger.addHandler(consoleHandler)
    logger.addHandler(fileHandler)
    return logger
