'''
file: IYATT_FillRandomFormula.py
description: 动态随机公式填充工具（常驻交互面板，支持手动修正规格参数、自定义任意步长、动态实时预览与连续填充）
author: IYATT-yx
copyright:   Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import re
import atexit
import ctypes
import logging
import unicodedata
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, messagebox

user32 = ctypes.windll.user32
GetAsyncKeyState = user32.GetAsyncKeyState
GetAsyncKeyState.argtypes = [ctypes.c_int]
GetAsyncKeyState.restype = wintypes.SHORT

VK_RETURN = 0x0D  # 回车键
ENTER_INTERVAL = 500  # 连续填充间隔（毫秒）

pluginInfo = {
    'name': 'FillRandomFormula',
    'author': 'IYATT',
    'description': '解析或手动调整规格，生成带边界缩减与自定义步长的动态随机公式',
    'version': '0.0.1',
}


def normalizeText(text: str) -> str:
    """全角转半角及符号标准化"""
    if not text:
        return ""
    text = unicodedata.normalize('NFKC', str(text))
    return text.lower().strip()

def getDecimalsFromStep(step: float) -> int:
    """根据步长推算显示的小数位数"""
    s = f"{step:.8f}".rstrip('0')
    if '.' in s:
        return len(s.split('.')[1])
    return 0

def parseSpecComponents(specStr: str):
    """
    解析规格文本，提取：名义值(Nominal), 上公差(UpperDev), 下公差(LowerDev)
    """
    normStr = normalizeText(specStr)
    if not normStr:
        return None, None, None

    # 1. 形位公差 / 单向上限公差（名义值为 0，下公差为 0，公差值全部作为上公差）
    if any(kw in normStr for kw in ["位置度", "垂直度", "同轴度", "平行度", "跳动", "倾斜度", "平面度", "直线度"]):
        match = re.search(r'(\d+\.?\d*)', normStr)
        if match:
            tolVal = float(match.group(1))
            return 0.0, tolVal, 0.0  # 名义值=0, 上公差=+tolVal, 下公差=0

    # 2. 正负对称公差（例：50 ± 0.05）
    matchSym = re.search(r'(\d+\.?\d*)\s*(?:±|\+\/-|\+-)\s*(\d+\.?\d*)', normStr)
    if matchSym:
        nominal = float(matchSym.group(1))
        tol = float(matchSym.group(2))
        return nominal, tol, -tol

    # 3. 双向极限偏差（例：50 (+0.03 / -0.01)）
    matchLim = re.search(r'(\d+\.?\d*)\s*\(\s*([+-]?\d+\.?\d*)\s*[,/]\s*([+-]?\d+\.?\d*)\s*\)', normStr)
    if matchLim:
        nominal = float(matchLim.group(1))
        dev1 = float(matchLim.group(2))
        dev2 = float(matchLim.group(3))
        return nominal, max(dev1, dev2), min(dev1, dev2)

    # 4. 单纯纯数字
    matchNum = re.search(r'^(\d+\.?\d*)$', normStr)
    if matchNum:
        return float(matchNum.group(1)), 0.0, 0.0

    return None, None, None

class ContinuousFillWindow:
    """常驻交互面板：支持手动规格调整、自定义可填步长、响应式实时计算与连续填充"""

    def __init__(self, appComHandle, logger=None):
        self._app = appComHandle
        self._logger = logger or logging.getLogger(pluginInfo['name'])

        self._activeFocus = "source"  # "source" 或 "target"
        self._srcCell = None
        self._destRng = None

        self._isListening = False
        self._lastEnterState = False
        self._afterId = None
        self._singleEnterTimer = None

        self._origMoveAfterReturn = True
        try:
            self._origMoveAfterReturn = self._app.MoveAfterReturn
        except Exception:
            pass

        self._root = tk.Tk()
        self._root.title("动态随机公式填充工具 (自定义可填步长版)")
        self._root.geometry("530x510")
        self._root.attributes("-topmost", True)
        self._root.resizable(False, False)

        atexit.register(self._stopListening)

        self._buildUi()
        self._bindEvents()

    def _buildUi(self):
        # 1. 顶部提示
        lblTip = ttk.Label(
            self._root,
            text="快捷操作：在 Excel 中选好单元格单击【回车】自动录入到源，快速双击录入到目标。\n识别公差后可手动修改名义值或偏差，步长支持下拉选择或直接手动输入任意数值。",
            justify="left",
            wraplength=500,
            font=("Microsoft YaHei", 9)
        )
        lblTip.pack(padx=12, pady=8, anchor="w")

        # 2. 选区关联面板
        frameSel = ttk.LabelFrame(self._root, text=" 1. 选区配置 ", padding=8)
        frameSel.pack(padx=12, pady=2, fill="x")

        ttk.Label(frameSel, text="源规格单元格:").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_src = ttk.Entry(frameSel, width=36)
        self.ent_src.grid(row=0, column=1, sticky="w", pady=3, padx=5)
        btnPickSrc = ttk.Button(frameSel, text="识别", width=8, command=self._parseSourceCell)
        btnPickSrc.grid(row=0, column=2, sticky="e", pady=3)

        ttk.Label(frameSel, text="目标填充区域:").grid(row=1, column=0, sticky="w", pady=3)
        self.ent_dest = ttk.Entry(frameSel, width=36)
        self.ent_dest.grid(row=1, column=1, sticky="w", pady=3, padx=5)
        btnPickDest = ttk.Button(frameSel, text="提取", width=8, command=self._captureDestRange)
        btnPickDest.grid(row=1, column=2, sticky="e", pady=3)

        # 3. 规格参数解析与手动修正面板
        frameSpec = ttk.LabelFrame(self._root, text=" 2. 规格分解与手动调整 ", padding=8)
        frameSpec.pack(padx=12, pady=6, fill="x")

        ttk.Label(frameSpec, text="名义值 (基准):").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_nominal = ttk.Entry(frameSpec, width=10)
        self.ent_nominal.grid(row=0, column=1, sticky="w", pady=3, padx=(0, 15))

        ttk.Label(frameSpec, text="上公差 (+):").grid(row=0, column=2, sticky="w", pady=3)
        self.ent_upper_dev = ttk.Entry(frameSpec, width=10)
        self.ent_upper_dev.grid(row=0, column=3, sticky="w", pady=3, padx=(0, 15))

        ttk.Label(frameSpec, text="下公差 (-):").grid(row=0, column=4, sticky="w", pady=3)
        self.ent_lower_dev = ttk.Entry(frameSpec, width=10)
        self.ent_lower_dev.grid(row=0, column=5, sticky="w", pady=3)

        # 4. 收缩策略与全过程极限实时计算显示
        frameCalc = ttk.LabelFrame(self._root, text=" 3. 边界收缩与控制极限实时计算 ", padding=8)
        frameCalc.pack(padx=12, pady=6, fill="x")

        # 策略参数
        ttk.Label(frameCalc, text="上限收缩(%):").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_upper_gap = ttk.Entry(frameCalc, width=8)
        self.ent_upper_gap.insert(0, "30")
        self.ent_upper_gap.grid(row=0, column=1, sticky="w", pady=3)

        ttk.Label(frameCalc, text="下限收缩(%):").grid(row=0, column=2, sticky="w", pady=3, padx=(12, 0))
        self.ent_lower_gap = ttk.Entry(frameCalc, width=8)
        self.ent_lower_gap.insert(0, "20")
        self.ent_lower_gap.grid(row=0, column=3, sticky="w", pady=3)

        ttk.Label(frameCalc, text="步长精度:").grid(row=0, column=4, sticky="w", pady=3, padx=(12, 0))
        
        # 允许直接敲击键盘手动输入任意步长数值
        self.cbo_step = ttk.Combobox(
            frameCalc,
            values=[
                "0.01", "0.001", "0.02", "0.05", 
                "0.002", "0.005", "0.0005", "0.0001", "0.1"
            ],
            width=8
        )
        self.cbo_step.current(0)
        self.cbo_step.grid(row=0, column=5, sticky="w", pady=3)

        # 分隔线
        sep = ttk.Separator(frameCalc, orient="horizontal")
        sep.grid(row=1, column=0, columnspan=6, sticky="ew", pady=6)

        # 源头理论极限显示（整数字号 9）
        ttk.Label(frameCalc, text="源头理论上极限:").grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        self.lbl_orig_max = ttk.Label(frameCalc, text="-", font=("Consolas", 9, "bold"))
        self.lbl_orig_max.grid(row=2, column=2, sticky="w", pady=2)

        ttk.Label(frameCalc, text="源头理论下极限:").grid(row=2, column=3, columnspan=2, sticky="w", pady=2)
        self.lbl_orig_min = ttk.Label(frameCalc, text="-", font=("Consolas", 9, "bold"))
        self.lbl_orig_min.grid(row=2, column=5, sticky="w", pady=2)

        # 缩减后实际控制范围显示（整数字号 9）
        ttk.Label(frameCalc, text="实际控制上极限:").grid(row=3, column=0, columnspan=2, sticky="w", pady=2)
        self.lbl_real_max = ttk.Label(frameCalc, text="-", font=("Consolas", 9, "bold"), foreground="#0055BB")
        self.lbl_real_max.grid(row=3, column=2, sticky="w", pady=2)

        ttk.Label(frameCalc, text="实际控制下极限:").grid(row=3, column=3, columnspan=2, sticky="w", pady=2)
        self.lbl_real_min = ttk.Label(frameCalc, text="-", font=("Consolas", 9, "bold"), foreground="#0055BB")
        self.lbl_real_min.grid(row=3, column=5, sticky="w", pady=2)

        # 5. 底部控制按钮
        frameBottom = ttk.Frame(self._root)
        frameBottom.pack(padx=12, pady=8, fill="x")

        btnFill = ttk.Button(frameBottom, text="🚀 确认填充", command=self._executeFill)
        btnFill.pack(side="right", padx=5)

        btnExit = ttk.Button(frameBottom, text="退出 (Esc)", command=self._closeWindow)
        btnExit.pack(side="right", padx=5)

    def _bindEvents(self):
        # 焦点监听，判断 Enter 键填充到哪个文本框
        self.ent_src.bind("<FocusIn>", lambda e: self._setActiveFocus("source"))
        self.ent_dest.bind("<FocusIn>", lambda e: self._setActiveFocus("target"))

        # 任意参数（包含步长编辑框）在键盘敲击释放时均重新计算
        for ent in [self.ent_nominal, self.ent_upper_dev, self.ent_lower_dev, self.ent_upper_gap, self.ent_lower_gap, self.cbo_step]:
            ent.bind("<KeyRelease>", lambda e: self._recalculateAll())

        # 步长从下拉菜单选中时也触发重新计算
        self.cbo_step.bind("<<ComboboxSelected>>", lambda e: self._recalculateAll())
        self._root.bind("<Escape>", lambda e: self._closeWindow())

    def _setActiveFocus(self, focusType: str):
        self._activeFocus = focusType

    def _startListening(self):
        self._isListening = True
        try:
            self._app.MoveAfterReturn = False
        except Exception:
            pass
        self._pollEnterKey()

    def _stopListening(self):
        self._isListening = False
        try:
            self._app.MoveAfterReturn = self._origMoveAfterReturn
        except Exception:
            pass

        if self._afterId and self._root:
            try:
                self._root.after_cancel(self._afterId)
            except Exception:
                pass
            self._afterId = None

        # 新增：清理回车定时器
        if self._singleEnterTimer and self._root:
            try:
                self._root.after_cancel(self._singleEnterTimer)
            except Exception:
                pass
            self._singleEnterTimer = None

    def _pollEnterKey(self):
        if not self._isListening:
            return

        try:
            if not self._root or not self._root.winfo_exists():
                self._isListening = False
                return

            state = GetAsyncKeyState(VK_RETURN)
            isPressed = bool(state & 0x8000)

            if isPressed and not self._lastEnterState:
                if self._singleEnterTimer is not None:
                    self._root.after_cancel(self._singleEnterTimer)
                    self._singleEnterTimer = None
                    self._captureDestRange()
                else:
                    self._singleEnterTimer = self._root.after(ENTER_INTERVAL, self._handleSingleEnterTimeout)

            self._lastEnterState = isPressed

        except Exception as e:
            if isinstance(e, tk.TclError):
                self._isListening = False
                return
            self._logger.error(f"按键检测异常: {e}")

        if self._isListening and self._root:
            try:
                if self._root.winfo_exists():
                    self._afterId = self._root.after(30, self._pollEnterKey)
            except tk.TclError:
                self._isListening = False

    def _handleSingleEnterTimeout(self):
        """超时未按第二下，执行单次 Enter 逻辑：选择源规格单元格"""
        self._singleEnterTimer = None
        self._parseSourceCell()

    def _parseSourceCell(self):
        """选中单元格并自动解析填充名义值、上公差、下公差（校验只允许单选，合并的单元格算1个）"""
        try:
            sel = self._app.Selection
            if not sel:
                return

            # 校验：只允许选择单个单元格或单个合并单元格
            try:
                if sel.Areas.Count > 1 or sel.Address != sel.Cells(1, 1).MergeArea.Address:
                    messagebox.showwarning("提示", "源规格单元格只允许选择单个单元格（合并的单元格算1个）！", parent=self._root)
                    return
            except Exception:
                messagebox.showwarning("提示", "源规格单元格只允许选择单个单元格！", parent=self._root)
                return

            self._srcCell = sel.Cells(1, 1)
            sheetName = self._srcCell.Worksheet.Name
            addr = self._srcCell.Address
            
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, f"[{sheetName}]!{addr}")

            # 解析规格文本
            specStr = str(self._srcCell.Value or "")
            nominal, upperDev, lowerDev = parseSpecComponents(specStr)

            # 更新编辑框
            self.ent_nominal.delete(0, tk.END)
            self.ent_upper_dev.delete(0, tk.END)
            self.ent_lower_dev.delete(0, tk.END)

            if nominal is not None:
                self.ent_nominal.insert(0, f"{nominal}")
                self.ent_upper_dev.insert(0, f"{upperDev}")
                self.ent_lower_dev.insert(0, f"{lowerDev}")

            # 触发重新计算
            self._recalculateAll()

        except Exception as e:
            self._logger.error(f"识别源单元格发生异常: {e}")

    def _captureDestRange(self):
        """提取目标填充区域"""
        try:
            sel = self._app.Selection
            if not sel:
                return

            self._destRng = sel
            sheetName = self._destRng.Worksheet.Name
            addr = self._destRng.Address

            self.ent_dest.delete(0, tk.END)
            self.ent_dest.insert(0, f"[{sheetName}]!{addr}")

        except Exception as e:
            self._logger.error(f"提取目标区域发生异常: {e}")

    def _recalculateAll(self):
        """核心函数：根据【名义值、上公差、下公差】、【收缩比例】及【自定义步长】，实时计算并刷新界面"""
        try:
            # 1. 获取名义值与偏差
            nominalStr = self.ent_nominal.get().strip()
            upperDevStr = self.ent_upper_dev.get().strip()
            lowerDevStr = self.ent_lower_dev.get().strip()

            if not nominalStr or not upperDevStr or not lowerDevStr:
                self._clearPreview()
                return

            nominal = float(nominalStr)
            upperDev = float(upperDevStr)
            lowerDev = float(lowerDevStr)

            # 2. 计算源头理论上下极限
            lMax = nominal + upperDev
            lMin = nominal + lowerDev

            if lMax < lMin:
                lMax, lMin = lMin, lMax

            # 3. 计算收缩后的实际控制极限
            uGapStr = self.ent_upper_gap.get().strip()
            lGapStr = self.ent_lower_gap.get().strip()

            uGap = float(uGapStr) / 100.0 if uGapStr else 0.0
            lGap = float(lGapStr) / 100.0 if lGapStr else 0.0

            bandwidth = lMax - lMin
            realMax = lMax - bandwidth * uGap
            realMin = lMin + bandwidth * lGap

            # 4. 刷新 UI 显示
            self.lbl_orig_min.config(text=f"{lMin:.4f}")
            self.lbl_orig_max.config(text=f"{lMax:.4f}")
            self.lbl_real_min.config(text=f"{realMin:.4f}")
            self.lbl_real_max.config(text=f"{realMax:.4f}")

        except ValueError:
            # 打字中途不合法数字时不显示
            self._clearPreview()

    def _clearPreview(self):
        self.lbl_orig_min.config(text="-")
        self.lbl_orig_max.config(text="-")
        self.lbl_real_min.config(text="-")
        self.lbl_real_max.config(text="-")

    def _executeFill(self):
        """根据当前 UI 面板上的数值生成 Excel 公式并填充"""
        if not self._destRng:
            messagebox.showwarning("提示", "请先选择需要填充的目标区域！", parent=self._root)
            return

        try:
            nominal = float(self.ent_nominal.get().strip())
            upperDev = float(self.ent_upper_dev.get().strip())
            lowerDev = float(self.ent_lower_dev.get().strip())

            uGap = float(self.ent_upper_gap.get().strip()) / 100.0
            lGap = float(self.ent_lower_gap.get().strip()) / 100.0
            
            # 兼容解析手动输入的步长（如 0.003 或 0.025）
            stepStr = self.cbo_step.get().strip()
            stepVal = float(stepStr)
            if stepVal <= 0:
                raise ValueError("步长必须大于0")
        except ValueError:
            messagebox.showerror("错误", "参数输入有误，请确保名义值、公差、收缩比例和步长均为有效的正数！", parent=self._root)
            return

        # 理论极限
        lMax = nominal + upperDev
        lMin = nominal + lowerDev
        if lMax <= lMin:
            messagebox.showerror("错误", f"上极限 ({lMax}) 必须大于下极限 ({lMin})，请检查规格参数！", parent=self._root)
            return

        # 控制极限
        bandwidth = lMax - lMin
        realMax = lMax - bandwidth * uGap
        realMin = lMin + bandwidth * lGap

        decimals = getDecimalsFromStep(stepVal)
        factor = 10 ** (decimals + 3)

        iMin = int(round(realMin * factor))
        iMax = int(round(realMax * factor))

        formulaStr = f"=MROUND(RANDBETWEEN({iMin}, {iMax})/{factor}, {stepVal})"
        numFormat = "0" if decimals == 0 else "0." + "0" * decimals

        # 批量应用到选定的目标区域
        self._destRng.Formula = formulaStr
        self._destRng.NumberFormatLocal = numFormat

        try:
            hwnd = self._app.Hwnd
            if hwnd:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass

        self._logger.info(f"填充成功: 区域 {self._destRng.Address} 已写入公式: {formulaStr}")

    def _closeWindow(self):
        self._stopListening()
        try:
            self._root.destroy()
        except Exception:
            pass

    def show(self):
        try:
            self._startListening()
            self._root.protocol("WM_DELETE_WINDOW", self._closeWindow)
            self._root.mainloop()
        except Exception as e:
            self._logger.error(f"插件运行异常: {e}")
            raise e
        finally:
            self._stopListening()


def run(appComHandle, logQueue):
    logger = logging.getLogger(pluginInfo['name'])
    win = ContinuousFillWindow(appComHandle, logger=logger)
    win.show()