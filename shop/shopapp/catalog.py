"""站点目录。先平台，再功能。通用功能不挂在平台下面。"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ToolSpec:
    id: str
    name: str
    ready: bool
    summary: str


@dataclass(frozen=True)
class PlatformSpec:
    id: str
    name: str
    tools: tuple[ToolSpec, ...]


def _tools(freight_ready: bool) -> tuple[ToolSpec, ...]:
    return (
        ToolSpec("freight", "运费计算", freight_ready, "按物流费表计算各渠道运费，并导出模板。"),
        ToolSpec("assets", "素材生成", False, "生成商品素材。"),
    )


PLATFORMS: tuple[PlatformSpec, ...] = (
    PlatformSpec("aliexpress", "速卖通", _tools(True)),
    PlatformSpec("shopee", "虾皮", _tools(False)),
    PlatformSpec("shein", "希音", _tools(False)),
    PlatformSpec("amazon", "亚马逊", _tools(False)),
)

COMMON: tuple[ToolSpec, ...] = (
    ToolSpec("text-translate", "文字翻译", True, "选择输入语言和输出语言，翻译填写的文字。"),
    ToolSpec("video-translate", "视频翻译", True, "选择输入语言和输出语言，上传视频后翻译其中的说话。"),
)


def platform_by_id(platform_id: str) -> PlatformSpec | None:
    for item in PLATFORMS:
        if item.id == platform_id:
            return item
    return None


def tool_by_id(platform: PlatformSpec, tool_id: str) -> ToolSpec | None:
    for item in platform.tools:
        if item.id == tool_id:
            return item
    return None


def catalog_payload() -> dict:
    def dump_tool(item: ToolSpec) -> dict:
        return {"id": item.id, "name": item.name, "ready": item.ready, "summary": item.summary}

    return {
        "platforms": [
            {"id": item.id, "name": item.name, "tools": [dump_tool(tool) for tool in item.tools]}
            for item in PLATFORMS
        ],
        "common": [dump_tool(item) for item in COMMON],
    }
