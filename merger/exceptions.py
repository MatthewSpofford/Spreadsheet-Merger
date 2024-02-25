
class MergeException(Exception):
    """Raised when an expected failure occurs during the merge"""

    def __init__(self, message):
        super().__init__(message)
