#!/usr/bin/env python3
from __future__ import annotations

import sys

from readerctl import main


if __name__ == "__main__":
    sys.argv.insert(1, "merge-draft")
    raise SystemExit(main())
