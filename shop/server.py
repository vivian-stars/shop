"""启动电商工具站。"""

from __future__ import annotations

import uvicorn


def main() -> None:
    uvicorn.run("shop.shopapp.api:app", host="0.0.0.0", port=3460, log_level="info")


if __name__ == "__main__":
    main()
