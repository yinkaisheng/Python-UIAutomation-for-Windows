from enum import IntEnum


class DpiAwarenessBehavior(IntEnum):
    ProcessDpiAwareness = 0
    ThreadDpiAwareness = 1
    NoDpiAwareness = 3


DPI_AWARENESS = DpiAwarenessBehavior.ProcessDpiAwareness
SEARCH_INTERVAL = 0.5
MAX_MOVE_SECOND = 1
TIME_OUT_SECOND = 10
OPERATION_WAIT_TIME = 0.5
PRINT_LOG = True
WRITE_LOG = True
DEBUG_SEARCH_TIME = False
DEBUG_EXIST_DISAPPEAR = False
