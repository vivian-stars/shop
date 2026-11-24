"""平台注册表。新平台在这里登记，页面不用改。"""

from shop.shopapp.platforms.aliexpress import AliExpressPlatform
from shop.shopapp.platforms.base import Platform

PLATFORMS: dict[str, Platform] = {
    AliExpressPlatform.id: AliExpressPlatform(),
}


def get_platform(platform_id: str) -> Platform:
    try:
        return PLATFORMS[platform_id]
    except KeyError as exc:
        known = "、".join(item.name for item in PLATFORMS.values())
        raise KeyError(f"没有这个平台。目前可用：{known}") from exc
