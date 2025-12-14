"""Compatibility wrapper.

Prefer running the module:

  python -m PickliPy.Bluetable ...

This file is kept so existing pipelines that call ``bluetable_picklist.py``
continue to work.
"""

from PickliPy.Bluetable import main


if __name__ == "__main__":
    main()
