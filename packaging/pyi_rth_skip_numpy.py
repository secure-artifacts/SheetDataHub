"""Make a missing or incomplete numpy look like ImportError.

openpyxl optionally imports numpy and then reads numpy.short. A half-bundled
numpy raises AttributeError instead of ImportError and crashes startup.
"""
from __future__ import annotations

import sys


def _numpy_usable() -> bool:
    try:
        import numpy
        return all(hasattr(numpy, name) for name in ("short", "int16", "ndarray"))
    except Exception:
        return False


if not _numpy_usable():
    for name in list(sys.modules):
        if name == "numpy" or name.startswith("numpy."):
            sys.modules.pop(name, None)

    class _NumpyBlocker:
        def find_spec(self, fullname, path, target=None):
            if fullname == "numpy" or fullname.startswith("numpy."):
                raise ImportError("numpy is not bundled with this application")
            return None

    sys.meta_path.insert(0, _NumpyBlocker())
