class ServerClosedException(Exception):
    def __init__(self, *args):
        super().__init__(*args)


class ArugmentNotReadyException(Exception):
    """Specifies that the argument is not ready tof calculation. 
    Use this error to signal for deferred calculation"""
    def __init__(self, *args):
        super().__init__(*args)