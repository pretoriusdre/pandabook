import re
from uuid import UUID
import numpy
from pandas import Timestamp
import numbers

EXCEL_ILLEGAL_CHARACTERS = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')


def sanitise_value(value):

    if value is None:
        return None
    
    if isinstance(value, str):
        value = re.sub(EXCEL_ILLEGAL_CHARACTERS, '', value)
        if value.startswith('=') and not value.startswith('=HYPERLINK'):
            # Append apostrophe before leading equals signs to prevent being interpreted as forumla
            value = "'" + value
            
    elif isinstance(value, bytes):
        # Attempt to decode bytes to string
        try:
            value = value.decode('utf-8')
        except Exception:
            value = str(value)

    elif isinstance(value, numpy.datetime64):
        value = Timestamp(value)

    elif isinstance(value, Timestamp):
        return value
        
    elif isinstance(value, numbers.Number):
        # Leave numeric types as is
        # Boolean is also captured here.
        pass

    else:
        # Convert arbitrary objects to strings, eg list, dict, UUID
        value = str(value)

    return value

