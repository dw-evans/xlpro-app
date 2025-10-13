class ServerClosedException(Exception):
    pass


class ArugmentNotReadyException(Exception):
    """Specifies that the argument is not ready tof calculation.
    Use this error to signal for deferred calculation"""

    pass


class ExcelArugmentIsNoneException(Exception):
    """Specifies that the argument is not ready tof calculation.
    Use this error to signal for deferred calculation"""

    pass


class ExcelNotAccessibleError(Exception):
    pass


class xlproUnhandledException(Exception):
    pass


class xlproArgumentExceptionError(Exception):
    pass


class xlproLikelyCOMAccessError(Exception):
    pass
