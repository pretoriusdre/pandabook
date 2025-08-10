import re
from uuid import UUID
import numpy
from pandas import Timestamp
import numbers

from pandabook.styles.defaults import DATE_ISO_STYLE, DATETIME_ISO_STYLE, SHRINK_TO_FIT


def format_by_value(value):

    if value is None:
        return None

    elif isinstance(value, (Timestamp, numpy.datetime64)):
        ts = Timestamp(value)
        if ts.hour == 0 and ts.minute == 0 and ts.second == 0 and ts.microsecond == 0:
            return DATE_ISO_STYLE
        else:
            return DATETIME_ISO_STYLE

    elif isinstance(value, UUID):
        return SHRINK_TO_FIT



