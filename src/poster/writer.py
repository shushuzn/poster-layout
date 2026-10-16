# -*- coding: utf-8 -*-
"""把求解结果写成 PPTX。

模板本身**不是空白**：占位文本框里残留着旧文字，必须先清空；
模板上的其余字段（作品编号 / 本科生组 / 作品名称 / 项目成员）**一律不动**。
占位框与内容区的判定都交给 :mod:`poster.template`，本模块只管往版面里画。
"""
from __future__ import annotations

import os
import time
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from . import config as C
from . import typography as T
from .layout import Layout, cap_size, cap_rows_of, fig_rows_of
from .repro import freeze_zip
from .template import is_placeholder          # noqa: F401  （对外保留同一入口）
from .typography import line_spacing_pct

_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def _sub(tag_xml):
    return parse_xml(tag_xml)


def _lin_angle(grad_fill, ang: int) -> None:
    """直接写 ``a:lin/@ang``。

    注意：**不要用 python-pptx 的 ``fill.gradient_angle`` setter** —— 它内部做了
    ``360 - value`` 换算，写 90 会落成 ``ang=16200000``（PowerPoint 的「线性向上」），
    与「自上而下」的意图相反。直接写 XML 才是可控的。
    映射：0=左→右，5400000=上→下，10800000=右→左，16200000=下→上。
    """
    lin = grad_fill.find(qn("a:lin"))
    if lin is None:
        lin = _sub('<a:lin %s ang="%d" scaled="0"/>' % (_NS, ang))
        grad_fill.append(lin)
    lin.set("ang", str(ang))
    lin.set("scaled", "0")


def _grad(sh, c1, c2, ang: int = 5400000) -> None:
    f = sh.fill
    f.gradient()
    st = f.gradient_stops
    st[0].color.rgb = RGBColor(*c1)
    st[0].position = 0.0
    st[1].color.rgb = RGBColor(*c2)
    st[1].position = 1.0
    gf = f._xPr.find(qn("a:gradFill"))
    if gf is not None:
        _lin_angle(gf, ang)


def _noshadow(sh) -> None:
    try:
        sh.shadow.inherit = False
    except Exception:                                          # noqa: BLE001
        pass


def _force_no_autofit(bp) -> None:
    """把 ``a:bodyPr`` 的自动调整钉成 ``noAutofit``，且**只留一个**。

    ``bodyPr`` 里 noAutofit / normAutofit / spAutoFit 是**三选一**。同时挂两个
    （新增文本框自带 spAutoFit，我们再 append 一个 noAutofit）会让 PowerPoint
    判定整份文件损坏、连打开都拒绝 —— 这种错最难查，所以这里先删干净再挂。
    """
    for t in ("a:normAutofit", "a:spAutoFit", "a:noAutofit"):
        for e in bp.findall(qn(t)):
            bp.remove(e)
    bp.append(_sub('<a:noAutofit %s/>' % _NS))


def _set_punct_hanging(bp) -> None:
    """让行尾标点「悬挂」在行外（PowerPoint 的「允许标点溢出边界」）。

    开了它，渲染器就不会把行尾放不下的「，。、」推到下一行，断行结果与本引擎
    的 ``hang`` 策略一致 —— 版面上更规范（行首不出现标点）。
    注意：这只是让"画出来更好看"；求解尺寸始终按两种策略里更悲观的那个算，
    所以哪怕某个渲染器不支持这个属性，也不会溢出压字。
    """
    if C.HANG_PUNCT:
        bp.set("hangingPunct", "1")


def _set_line_spacing(par, factor: float) -> None:
    """写 ``spcPct`` 行距。

    **不要用 ``paragraph.line_spacing = Pt(x)``**（那是 ``spcPts``）：PowerPoint 会把
    绝对点值行距四舍五入到整点，5.52pt 会被按 6.0pt 排，一行多出 8.7% 高度，
    整块多一行，末行就钻到下面的配图底下被盖住。``spcPct`` 是精确的。
    """
    pPr = par._p.get_or_add_pPr()
    old = pPr.find(qn("a:lnSpc"))
    if old is not None:
        pPr.remove(old)
    pPr.insert(0, _sub('<a:lnSpc %s><a:spcPct val="%d"/></a:lnSpc>'
                       % (_NS, line_spacing_pct(factor))))


class PosterWriter:
    """承载一次构建的全部形状操作。"""

    def __init__(self, template_path):
        self.prs = Presentation(str(template_path))
        self.slide = self.prs.slides[0]

    # ------------------------------------------------------------ 基础形状
    def rrect(self, x, y, w, h, dashed=True, line=C.BLOCK_LINE, lw=C.BORDER_W,
              grad=True, fill=None):
        s = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Cm(x), Cm(y), Cm(w), Cm(h))
        if grad:
            _grad(s, C.BLOCK_C1, C.BLOCK_C2, 5400000)
        else:
            s.fill.solid()
            s.fill.fore_color.rgb = RGBColor(*(fill or C.PLATE))
        if line:
            s.line.color.rgb = RGBColor(*line)
            s.line.width = Pt(lw)
            if dashed:
                s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        else:
            s.line.fill.background()
        _noshadow(s)
        s.adjustments[0] = max(0.0, min(0.5, C.RADIUS / max(0.02, min(w, h))))
        s.text_frame.text = ""
        return s

    def chip(self, x, y, w, h, text):
        """胶囊小标题：全圆角 + 左→右渐变 + 白字粗体，文字真垂直居中。"""
        s = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Cm(x), Cm(y), Cm(w), Cm(h))
        _grad(s, C.CHIP_C1, C.CHIP_C2, 0)
        s.line.fill.background()
        _noshadow(s)
        s.adjustments[0] = 0.5

        tf = s.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        # 光设垂直锚点不够：主题默认的行距/段距会把字顶偏。
        # 行距写成 1 倍（= 字号 x 1.20 x 83.3% = 字号），且一律走 spcPct。
        _set_line_spacing(p, 1.0)
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        if pPr.find(qn("a:buNone")) is None:
            pPr.append(_sub('<a:buNone %s/>' % _NS))
        bp = tf._txBody.find(qn("a:bodyPr"))
        if bp is not None:
            _force_no_autofit(bp)

        r = p.add_run()
        r.text = text
        r.font.size = Pt(C.TITLE_SZ)
        r.font.name = C.FONT_FAMILY
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        self._set_ea_font(r._r.get_or_add_rPr())
        return s

    def dashed_link(self, x1, x2, y):
        """胶囊背后的整条虚线（从块左内边铺到右内边）。"""
        c = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                            Cm(x1), Cm(y), Cm(x2), Cm(y))
        c.line.color.rgb = RGBColor(*C.CHIP_LINK)
        c.line.width = Pt(0.6)
        c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        _noshadow(c)
        return c

    def add_text(self, x, y, w, h, paras, size, color=C.BLACK,
                 align=PP_ALIGN.LEFT, line_factor=C.LINE_FACTOR):
        tb = self.slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        bp = tf._txBody.find(qn("a:bodyPr"))
        if bp is not None:
            _force_no_autofit(bp)
            _set_punct_hanging(bp)

        first = True
        for para in paras:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            _set_line_spacing(p, line_factor)      # spcPct，不用 spcPts
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            if pPr.find(qn("a:buNone")) is None:
                pPr.append(_sub('<a:buNone %s/>' % _NS))
            for seg in para:
                text, col, script = T.unpack(seg)
                r = p.add_run()
                r.text = text
                r.font.size = Pt(size)
                r.font.name = C.FONT_FAMILY
                r.font.color.rgb = RGBColor(*col)
                rPr = r._r.get_or_add_rPr()
                # 字符间距（1/100 pt）：汉字密排时透一口气。量宽侧同口径，见 typography._spc_px
                rPr.set("spc", str(int(round(C.CHAR_SPC_PT * 100))))
                if script in C.SCRIPT_BASELINE:      # 化学式上/下标：只位移，不缩放
                    rPr.set("baseline", str(C.SCRIPT_BASELINE[script]))
                self._set_ea_font(rPr)
        return tb

    @staticmethod
    def _set_ea_font(rPr) -> None:
        """同步东亚/复杂文种字体，避免中文回落到宋体。"""
        for tag in ("a:ea", "a:cs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = _sub('<%s %s typeface="%s"/>' % (tag, _NS, C.FONT_FAMILY))
                rPr.append(el)
            el.set("typeface", C.FONT_FAMILY)

    # ------------------------------------------------------------ 组装
    def clear_placeholder(self) -> bool:
        """清空模板里残留的占位文本框。"""
        for sh in self.slide.shapes:
            if is_placeholder(sh):
                sh.text_frame.clear()
                sh.width = Cm(0.01)
                sh.height = Cm(0.01)
                return True
        return False

    def set_number(self, number: str) -> bool:
        """把模板的「作品编号」框改成 ``作品编号：<number>``。

        模板里这个框预置的是占位文案（例如「作品编号：对应缴费编号」）。只有
        任务文档写了 ``number:`` 才动它 —— 没写就**一字不改**，与作品名称 /
        项目成员 / 本科生组三个字段一样。格式沿用原框第一个 run 的字号与颜色，
        只换文字，不新建 run。
        """
        if not number:
            return False
        for sh in self.slide.shapes:
            if not sh.has_text_frame:
                continue
            if not sh.text_frame.text.strip().startswith(C.NUMBER_PREFIX):
                continue
            runs = sh.text_frame.paragraphs[0].runs
            if not runs:
                continue
            runs[0].text = "%s：%s" % (C.NUMBER_PREFIX, number)
            for r in runs[1:]:                  # 占位文案若被拆成多段 run，清掉尾巴
                r.text = ""
            return True
        return False

    def place_figs(self, bf, caps_rows, ix: float, iy: float, ihs: list,
                   caphs: list, cap_sz: float, fw_max: float) -> None:
        """把一块的配图画成**图网格**：行内等高横排，行间上下叠。

        每行的图在块内居中（以最宽那条图条为准对齐），
        并**逐张**在各自图下写图注 —— 一张图一条注，不合并。
        """
        rows = fig_rows_of(bf) or [bf]
        y = iy
        for k, rb in enumerate(rows):
            ih = ihs[k] if k < len(ihs) else ihs[-1]
            caph = caphs[k] if k < len(caphs) else caphs[-1]
            caps = caps_rows[k] if k < len(caps_rows) else []
            fw = sum(ih * s for s in rb.spans) + max(0, len(rb.paths) - 1) * C.FIG_GAP_IN
            x = ix + max(0.0, (fw_max - fw)) / 2
            for path, span, cap in zip(rb.paths, rb.spans, caps):
                w = ih * span
                pic = self.slide.shapes.add_picture(str(path), Cm(x), Cm(y),
                                                    width=Cm(w), height=Cm(ih))
                try:
                    pic.line.color.rgb = RGBColor(*C.IMG_LINE)
                    pic.line.width = Pt(0.5)
                except Exception:                          # noqa: BLE001
                    pass
                if cap:
                    self.add_text(x, y + ih + C.CAP_LINEGAP, w, caph,
                                  [[(cap, C.GRAY)]], cap_sz,
                                  align=PP_ALIGN.CENTER, line_factor=C.CAP_LINE_FACTOR)
                x += w + C.FIG_GAP_IN
            y += ih + C.CAP_LINEGAP + caph + C.FIG_ROW_GAP

    def emit(self, lay: Layout, task, figs, with_plate: bool = True) -> int:
        """按版面把所有形状写进幻灯片，返回最终形状数。"""
        from .typography import font as mfont

        def pill_w(text):
            return mfont(C.TITLE_SZ, True).getlength(text) / C.PX_PER_CM + 2 * C.CHIP_PAD_X

        if with_plate:
            self.rrect(C.X0, C.Y0, C.W, C.H, dashed=False, line=None, grad=False)
        self.clear_placeholder()
        self.set_number(getattr(task, "number", ""))

        cells = lay.cells
        for c in cells:                                   # 块底
            self.rrect(c.x, c.y, c.bw, c.h)
        for c in cells:                                   # 虚线（先画，垫在胶囊下）
            self.dashed_link(c.x + C.BP_X, c.x + c.bw - C.BP_X,
                             c.y + C.TITLE_TOP + C.CHIP_H / 2)
        for c in cells:                                   # 胶囊（后画，骑在线上）
            cw = min(pill_w(task.titles[c.r]), c.bw - 2 * C.BP_X)
            self.chip(c.x + C.BP_X, c.y + C.TITLE_TOP, cw, C.CHIP_H, task.titles[c.r])

        cap_sz = cap_size(lay.body_pt)
        for c in cells:                                   # 正文 + 图 + 图注
            row_top = c.y + C.TITLE_TOP + C.CHIP_H + C.TITLE_GAP + c.off
            if c.style == "stack":
                self.add_text(c.x + C.BP_X, row_top, c.tw, c.th + 0.02,
                              task.paras[c.r], lay.body_pt)
                iy = row_top + c.th + C.TEXT_IMG_GAP
                ix = c.x + C.BP_X + (c.inner - c.fw) / 2
            elif c.style == "side":
                iy = row_top + max(0.0, (c.ch - (c.ih + C.CAP_LINEGAP + c.caph)) / 2)
                t_y = row_top + max(0.0, (c.ch - c.th) / 2)
                self.add_text(c.x + C.BP_X, t_y, c.tw, c.th + 0.02,
                              task.paras[c.r], lay.body_pt)
                ix = c.x + c.bw - C.BP_X - c.fw
            else:                                          # 纯文本块
                self.add_text(c.x + C.BP_X, row_top, c.tw, c.th + 0.02,
                              task.paras[c.r], lay.body_pt)
                continue

            self.place_figs(figs[c.r], cap_rows_of(task, c.r, figs[c.r]),
                            ix, iy, c.ihs, c.caphs, cap_sz, c.fw)
        return len(self.slide.shapes)

    def save(self, path):
        """原子落盘：先写同目录临时文件，再 ``os.replace`` 顶掉目标。

        为什么不直接 ``prs.save(目标)``：目标随时可能被外部程序（PowerPoint 查看、
        资源管理器预览、同步 / 索引 / 查杀服务）以共享方式打开着，python-pptx 会
        直接抛 ``PermissionError``，**整个构建随之失败**（F3 复现性也跟着红）。
        改成"写临时文件 + 原子替换"只动目录项，这类占用基本都能绕开；仍失败时
        退避重试几轮，且**原产物中途不会被写坏**。
        """
        path = Path(path)
        tmp = path.with_name(path.stem + ".tmp" + path.suffix)
        last = None
        for i in range(5):
            try:
                self.prs.save(str(tmp))
                freeze_zip(tmp)      # 归一化 zip 时间戳，保证同输入同字节
                os.replace(tmp, path)
                return
            except OSError as e:     # 目标被占用 / 共享冲突：等一下再试
                last = e
                time.sleep(0.4 * (i + 1))
        raise RuntimeError("产物写入失败（反复被占用）：%s" % last)


def build_pptx(lay: Layout, task, figs, template_path, out_path) -> int:
    w = PosterWriter(template_path)
    n = w.emit(lay, task, figs)
    w.save(out_path)
    return n
