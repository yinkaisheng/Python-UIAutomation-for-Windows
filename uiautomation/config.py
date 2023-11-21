from enum import IntEnum


class DpiAwarenessBehavior(IntEnum):
    ProcessDpiAwareness = 0
    ThreadDpiAwareness = 1
    NoDpiAwareness = 3


class ProcessDpiAwareness(IntEnum):
    DpiUnaware = 0
    SystemDpiAware = 1
    PerMonitorDpiAware = 2


class DpiAwarenessContext(IntEnum):
    Unaware = -1
    SystemAware = -2
    PerMonitorAware = -3
    PerMonitorAwareV2 = -4
    UnawareGdiScaled = -5


DPI_AWARENESS_BEHAVIOR = DpiAwarenessBehavior.ProcessDpiAwareness
DPI_AWARENESS_VALUE = ProcessDpiAwareness.PerMonitorDpiAware # use DpiAwarenessContext value if it's ThreadDpiAwareness
SEARCH_INTERVAL = 0.5
MAX_MOVE_SECOND = 1
TIME_OUT_SECOND = 10
OPERATION_WAIT_TIME = 0.5
PRINT_LOG = True
WRITE_LOG = True
DEBUG_SEARCH_TIME = False
DEBUG_EXIST_DISAPPEAR = False
