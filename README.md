<p align="center">
  <img src="assets/app_icon.png" width="120" alt="Figure Studio Icon" />
</p>

<h1 align="center">科研绘图排版工具 (Figure Studio)</h1>

<p align="center">
  <b>纯 MATLAB 科研论文自由拼图 · Word 1:1 版心定宽 · 多图字体字号完全统一</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/MATLAB-R2020b%2B-0076A8?style=flat-square" alt="MATLAB" />
  <img src="https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-5856D6?style=flat-square" alt="Platform" />
  <img src="https://img.shields.io/badge/License-MIT-26AD6B?style=flat-square" alt="MIT License" />
</p>

---

## 为什么做这个工具？

写论文时用 MATLAB 导出多张 `.fig` 再放进 Word / PPT 拼图，最让人头疼的问题就是：**每张图原始窗口大小和字号都不一样，贴进 Word 一缩放，图里的字号全乱了**（有的太大、有的小到看不清，和论文正文字号完全不匹配）。

**Figure Studio（科研绘图排版工具）** 直接用纯 MATLAB 从根源解决这一痛点：
1. **自动读取论文 Word（`.docx`）版心宽度与正文字号**；
2. **在类 PPT 厘米标尺画布上自由拖拽排版**（支持任意摆放、对齐、紧凑铺满）；
3. **以 1:1 真实印刷厘米尺寸重绘每个 `.fig` 并统一全部字体/字号/线宽**；
4. **导出自带精确物理 DPI 元数据的高清 PNG / TIFF，或直接打开原生可编辑 `.fig` 用 MATLAB 属性检查器精修**——拖进 Word 默认就是 100% 原始宽度，图内 10 号字与 Word 正文 10 号字肉眼和印刷完全一致。

---

## 核心功能

- **📄 Word (`.docx`) 智能解析**
  - 一键导入论文 `.docx`，自动识别页面尺寸、左右页边距、单/双栏宽度以及默认正文字体与字号（如 `Times New Roman 10 pt`）。
  - 双栏论文支持一键切换 **通栏宽 (Full)** 与 **单栏宽 (Column)**，也支持手动输入任意厘米宽高。
- **🧩 类 PPT 自由排版与一键紧凑铺满**
  - 内置 `1×1 单图铺满`、`1×2`、`1×3`、`1×4`、`2×2`、`3+2 (上3下2)`、`2+3`、`2×3`、`1+2 (左1右2)` 等常用科研排版模板。
  - 支持自由添加/复制/删除图框、鼠标拖拽移动、Keynote 风格圆点 8 向缩放、网格与红色对齐参考线吸附、`Shift+点击` 多选批量对齐/等宽/等高/均分、方向键 `0.1 cm` 微调。
  - **独立列间距 / 行间距实时调节**与**「一键紧凑靠拢并放大铺满版心宽」**，告别图间大块留白；只放 1 张图时自动启用**单图模式**（忽略空框且不加多余 `(a)` 编号）。
- **🔤 1:1 真实物理字号统一（深度适配真实 `.fig`）**
  - 全局字号一键同步，也支持单独微调刻度、坐标轴标签、标题、图例、图内标注及子图编号 `(a)(b)...`。
  - 自动修复真实科研 `.fig` 常见坑：重置手动写死的标签位置防止与刻度重叠、自动修正自由定位图例防止冲出右边界、超长标签防裁切、完整支持 **`yyaxis` 双 Y 轴**、可选统一常规字重与线宽。
- **🛠️ 原生可编辑 `.fig` & MATLAB 属性检查器微调**
  - 支持一键打开 **等比大屏放大（所见即所得 WYSIWYG）的原生可编辑 MATLAB Figure**（包含双 Y 轴无损重建），可直接使用 MATLAB **属性检查器（Property Inspector）** 拖拽图例、修改曲线/文字，微调完后点击顶部菜单 **「科研微调与导出」** 即可按精确 1:1 Word DPI 导出。
- **🖼️ 高清印刷导出 & 排版方案存取**
  - 支持 `300 / 600 / 1200 DPI` 的 PNG 与无损 LZW TIFF（写入精确物理 DPI 元数据，拖进 Word 自动 100% 原始厘米宽度），并支持直接导出 `.fig` 与 `.json` 排版方案复用。
- **✨ Apple 磨砂玻璃风格 UI & 跨平台静默双击启动**
  - macOS / Windows 均提供独立双击启动器，**不弹出 MATLAB 主界面和终端黑框**，关闭窗口自动退出后台引擎。

---

## 快速开始

### 环境要求
- 已安装 **MATLAB**（推荐 `R2020b` 及以上版本，已在 `MATLAB R2026b` 下完成实测）。
- 无需任何第三方工具箱，纯内置 MATLAB 运行。

### 方式一：直接双击启动（推荐）
- **macOS**：直接双击项目目录下的 **`科研绘图排版工具.app`**（可拖到桌面或程序坞随时打开）。
- **Windows**：
  - 直接双击 **`启动科研绘图排版工具_Windows.vbs`**；
  - 或先双击一次 **`创建Windows桌面快捷方式.vbs`**，在桌面生成带磨砂玻璃图标的 **`Figure Studio`** 快捷方式后直接双击使用。

### 方式二：在 MATLAB 命令行启动
将 MATLAB 当前目录切换到本项目文件夹，运行：
```matlab
run('start_app.m')
```

---

## 项目结构

```text
Figure-Studio/
├── start_app.m                      # MATLAB 命令行一键启动脚本
├── FigureLayoutApp.m                # 主 GUI 应用 (uifigure，Apple 磨砂玻璃风格)
├── 科研绘图排版工具.app/             # macOS 独立无桌面双击启动器（含 AppIcon.icns）
├── 启动科研绘图排版工具_Windows.vbs  # Windows 静默无黑框双击启动脚本
├── 创建Windows桌面快捷方式.vbs       # Windows 一键创建带图标桌面快捷方式
├── assets/
│   ├── app_icon.png                 # 窗口图标 (256x256 透明圆角)
│   ├── app_icon_1024.png            # 高清母版图标 (1024x1024)
│   └── app_icon.ico                 # Windows 多分辨率图标
├── core/
│   ├── parseWordDocx.m              # Word (.docx) 版心与正文字号解析
│   ├── getLayoutPresets.m           # 常用科研排版模板 (1x1 ~ 3+2 等)
│   ├── styleFigure.m                # .fig 物理尺寸锁定、字体/线宽统一与防裁切
│   ├── renderSingleSlot.m           # 单子图 1:1 物理尺寸高清渲染
│   ├── renderComposite.m            # 多图精准拼合、单图智能裁剪与编号叠加
│   ├── buildEditableFigure.m        # 原生可编辑 .fig 组装（含双 Y 轴重建与 WYSIWYG 缩放）
│   └── saveImageWithDpi.m           # PNG/TIFF 写入精确物理 DPI 元数据
├── tests/
│   └── run_verification.m           # 端到端自动化验证脚本
├── LICENSE                          # MIT 开源协议
└── README.md
```

---

## 运行自动化测试

在 MATLAB 命令行中运行以下脚本，可自动生成测试 `.docx` 和 `.fig`，验证版心解析、多图统一字号拼图、DPI 元数据及 App 启动：

```matlab
run('tests/run_verification.m')
```

---

## 开源协议

本项目基于 [MIT License](LICENSE) 开源，欢迎科研工作者自由使用、修改与分享。
