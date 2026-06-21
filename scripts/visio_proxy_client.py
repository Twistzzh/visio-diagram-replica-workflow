"""Visio COM 代理客户端 - Python 封装

在 Codex 沙箱用户中调用，通过文件型 IPC 与 ASUS 交互用户中的 Visio 代理通信。

用法:
    from visio_proxy_client import VisioProxy

    v = VisioProxy()
    v.new_document("NVH Test Setup")
    v.set_page(width_px=1180, height_px=780)
    v.draw_rect(x1=80, y1=30, x2=600, y2=100, fill="RGB(245,245,245)", text="LMS 数据采集前端")
    v.save_as(r"D:\codex\projects\draw\outputs\test.vsdx")
"""

import json
import os
import time
import uuid
from pathlib import Path

DEFAULT_PROXY_DIR = Path(r"D:\codex\projects\draw\scratch\visio_proxy")
DEFAULT_TIMEOUT = 120


class VisioProxyError(Exception):
    """Visio 代理操作异常"""
    pass


class VisioProxy:
    """Visio COM 代理客户端"""

    def __init__(self, proxy_dir: Path | str = DEFAULT_PROXY_DIR, timeout: int = DEFAULT_TIMEOUT):
        self.proxy_dir = Path(proxy_dir)
        self.cmd_dir = self.proxy_dir / "commands"
        self.rsp_dir = self.proxy_dir / "responses"
        self.lock_file = self.proxy_dir / "server.lock"
        self.timeout = timeout

        if not self.lock_file.exists():
            raise VisioProxyError(
                "Visio COM 代理未运行！请让 ASUS 用户双击运行:\n"
                "  D:\\codex\\projects\\draw\\scripts\\start_visio_proxy.bat"
            )

    def _send_command(self, action: str, args: dict | None = None) -> dict:
        """发送命令到代理并等待响应"""
        cmd_id = uuid.uuid4().hex
        cmd_file = self.cmd_dir / f"{cmd_id}.json"
        rsp_file = self.rsp_dir / f"{cmd_id}.json"

        cmd = {"action": action, "args": args or {}}
        cmd_file.write_text(json.dumps(cmd, ensure_ascii=False), encoding="utf-8")

        deadline = time.time() + self.timeout
        while time.time() < deadline:
            if rsp_file.exists():
                result = json.loads(rsp_file.read_text(encoding="utf-8"))
                rsp_file.unlink(missing_ok=True)
                if not result.get("success"):
                    raise VisioProxyError(result.get("error", "未知错误"))
                return result
            time.sleep(0.2)

        raise VisioProxyError(f"Visio 命令超时 ({self.timeout}s): {action}")

    def ping(self) -> dict:
        """检查代理是否存活"""
        return self._send_command("ping")

    def new_document(self, name: str = "Visio 复刻图") -> dict:
        """创建新文档"""
        return self._send_command("new_document", {"name": name})

    def set_page(self, width_px: float, height_px: float,
                 scale_px_per_inch: float = 100.0) -> dict:
        """设置页面尺寸"""
        return self._send_command("set_page", {
            "widthPx": width_px,
            "heightPx": height_px,
            "scalePxPerInch": scale_px_per_inch,
        })

    def draw_rect(self, x1: float, y1: float, x2: float, y2: float,
                  fill: str = "RGB(255,255,255)", line: str = "RGB(35,35,35)",
                  weight: float = 1.0, dash: bool = False,
                  text: str = "", font_size: float = 8.0,
                  text_color: str = "RGB(40,40,40)", bold: bool = False,
                  align: int = 1, page_height_px: float = 800,
                  scale_px_per_inch: float = 100.0) -> dict:
        """绘制矩形"""
        style = {
            "fill": fill, "line": line, "weight": weight, "dash": dash,
            "fontSize": font_size, "textColor": text_color, "bold": bold, "align": align,
        }
        return self._send_command("draw_shape", {
            "type": "rect",
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "style": style, "text": text,
            "pageHeightPx": page_height_px,
            "scalePxPerInch": scale_px_per_inch,
        })

    def draw_oval(self, x1: float, y1: float, x2: float, y2: float,
                  fill: str = "RGB(255,255,255)", line: str = "RGB(35,35,35)",
                  weight: float = 1.0, text: str = "", font_size: float = 8.0,
                  text_color: str = "RGB(40,40,40)", bold: bool = False,
                  page_height_px: float = 800,
                  scale_px_per_inch: float = 100.0) -> dict:
        """绘制椭圆"""
        style = {
            "fill": fill, "line": line, "weight": weight,
            "fontSize": font_size, "textColor": text_color, "bold": bold,
        }
        return self._send_command("draw_shape", {
            "type": "oval",
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "style": style, "text": text,
            "pageHeightPx": page_height_px,
            "scalePxPerInch": scale_px_per_inch,
        })

    def draw_line(self, x1: float, y1: float, x2: float, y2: float,
                  line: str = "RGB(35,35,35)", weight: float = 1.0,
                  arrow: str | None = "end",
                  page_height_px: float = 800,
                  scale_px_per_inch: float = 100.0) -> dict:
        """绘制线条"""
        style = {"line": line, "weight": weight}
        return self._send_command("draw_shape", {
            "type": "line",
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "style": style, "arrow": arrow,
            "pageHeightPx": page_height_px,
            "scalePxPerInch": scale_px_per_inch,
        })

    def draw_text(self, x1: float, y1: float, x2: float, y2: float,
                  text: str, font_size: float = 8.0,
                  text_color: str = "RGB(40,40,40)", bold: bool = False,
                  align: int = 1, page_height_px: float = 800,
                  scale_px_per_inch: float = 100.0) -> dict:
        """绘制文字标签"""
        style = {
            "fontSize": font_size, "textColor": text_color, "bold": bold, "align": align,
        }
        return self._send_command("draw_shape", {
            "type": "text",
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "style": style, "text": text,
            "pageHeightPx": page_height_px,
            "scalePxPerInch": scale_px_per_inch,
        })

    def save_as(self, path: str) -> dict:
        """保存为 .vsdx"""
        return self._send_command("save_as", {"path": path})

    def export_page(self, path: str) -> dict:
        """导出当前页为 .emf"""
        return self._send_command("export_page", {"path": path})

    def close_document(self) -> dict:
        """关闭当前文档"""
        return self._send_command("close_document")

    def get_active_page(self) -> dict:
        """获取当前页面信息"""
        return self._send_command("get_active_page")
