classdef FigureLayoutApp < handle
%FIGURELAYOUTAPP 科研绘图自由排版与字体统一工具（纯 MATLAB App）
%
%   工作流：
%     1. 导入 Word (.docx) 自动识别版心宽度与正文字号（也可手动输入宽度）
%     2. 在类 PPT 厘米标尺画布上自由拖拽、缩放、对齐图框，或一键套用模板 (1x3, 1x4, 3+2 等)
%     3. 双击图框（或点击按钮）放入 MATLAB .fig 文件
%     4. 统一全部 .fig 的字体与字号（1:1 真实物理尺寸渲染），导出带精确 DPI 元数据的高清 PNG/TIFF

    properties (Access = public)
        UIFigure            matlab.ui.Figure
    end

    properties (Access = private)
        % 启动模式（独立双击启动时关闭窗口自动退出后台 MATLAB）
        StandaloneMode (1,1) logical = false

        % 数据状态
        CanvasWidthCm (1,1) double = 14.65   % 默认 A4 中文论文常见版心宽 (cm)
        CanvasHeightCm (1,1) double = 18.00  % 默认比宽度略长
        GapXCm (1,1) double = 0.10           % 水平列间距 (cm)
        GapYCm (1,1) double = 0.12           % 垂直行间距 (cm)
        SnapEnabled (1,1) logical = true
        SnapStepCm (1,1) double = 0.20

        WordInfo            struct = struct()
        Slots               struct = struct([])
        NextSlotId (1,1)    double = 1
        SelectedIds         double = []

        % 拖拽状态
        DragMode (1,:)      char = 'none'    % 'none' | 'move' | 'resize'
        DragHandle (1,:)    char = ''        % 'nw','n','ne','e','se','s','sw','w'
        DragSlotIdx (1,1)   double = 0
        DragStartPt (1,2)   double = [0, 0]
        DragOrigRect (1,4)  double = [0, 0, 0, 0]
        DragOrigMulti       double = []      % 多选移动时的初始坐标矩阵
        GuideLines          = gobjects(0)

        % UI 控件句柄
        CanvasAxes          matlab.ui.control.UIAxes
        StatusLabel         matlab.ui.control.Label

        % 左侧：Word & 画布
        WordSummaryLabel    matlab.ui.control.Label
        ColModeDropDown     matlab.ui.control.DropDown
        CanvasWField        matlab.ui.control.NumericEditField
        CanvasHField        matlab.ui.control.NumericEditField
        GapXField           matlab.ui.control.NumericEditField
        GapYField           matlab.ui.control.NumericEditField

        % 左侧：模板 & 吸附
        PresetDropDown      matlab.ui.control.DropDown
        SnapCheck           matlab.ui.control.CheckBox

        % 左侧：字体与外观
        FontNameDropDown    matlab.ui.control.DropDown
        GlobalFontField     matlab.ui.control.NumericEditField
        TickFontField       matlab.ui.control.NumericEditField
        LabelFontField      matlab.ui.control.NumericEditField
        TitleFontField      matlab.ui.control.NumericEditField
        LegendFontField     matlab.ui.control.NumericEditField
        TextFontField       matlab.ui.control.NumericEditField

        SubLabelCheck       matlab.ui.control.CheckBox
        SubLabelStyleDD     matlab.ui.control.DropDown
        SubLabelPosDD       matlab.ui.control.DropDown
        SubLabelSizeField   matlab.ui.control.NumericEditField
        SubLabelBoldCheck   matlab.ui.control.CheckBox

        TightLayoutCheck    matlab.ui.control.CheckBox
        ForceNormalCheck    matlab.ui.control.CheckBox
        AxisLWCheck         matlab.ui.control.CheckBox
        AxisLWField         matlab.ui.control.NumericEditField
        PlotLWCheck         matlab.ui.control.CheckBox
        PlotLWField         matlab.ui.control.NumericEditField

        % 左侧：导出
        FormatDropDown      matlab.ui.control.DropDown
        DpiDropDown         matlab.ui.control.DropDown
        AutoCropCheck       matlab.ui.control.CheckBox
        SaveSubCheck        matlab.ui.control.CheckBox

        % 右侧：属性检查器
        SlotListBox         matlab.ui.control.ListBox
        InspectorInfoLabel  matlab.ui.control.Label
        SlotXField          matlab.ui.control.NumericEditField
        SlotYField          matlab.ui.control.NumericEditField
        SlotWField          matlab.ui.control.NumericEditField
        SlotHField          matlab.ui.control.NumericEditField
        SlotLabelEdit       matlab.ui.control.EditField
        SlotFileLabel       matlab.ui.control.Label
        LoadFigBtn          matlab.ui.control.Button
        RefreshThumbBtn     matlab.ui.control.Button
        ClearFigBtn         matlab.ui.control.Button
        FullWidthBtn        matlab.ui.control.Button
    end

    methods (Access = public)
        function app = FigureLayoutApp(standaloneMode)
            if nargin >= 1 && ~isempty(standaloneMode)
                app.StandaloneMode = logical(standaloneMode);
            end

            % 确保 core 目录在 MATLAB 搜索路径中
            appDir = fileparts(mfilename('fullpath'));
            coreDir = fullfile(appDir, 'core');
            if isfolder(coreDir)
                addpath(coreDir);
            end

            app.buildUI();
            if app.StandaloneMode
                app.UIFigure.CloseRequestFcn = @(~,~) app.onStandaloneClose();
            end
            % 启动时默认加载 3+2 排版模板供直观上手
            app.applyPreset('3+2');
            app.setStatus('就绪：可先导入 Word (.docx) 确定版心宽度，再双击蓝框放入 .fig 文件。');
        end

        function delete(app)
            if ~isempty(app.UIFigure) && isvalid(app.UIFigure)
                delete(app.UIFigure);
            end
        end
    end

    methods (Access = private)
        function onStandaloneClose(app)
            if ~isempty(app.UIFigure) && isvalid(app.UIFigure)
                delete(app.UIFigure);
            end
            % 独立双击模式下，关闭 App 窗口时自动退出后台 MATLAB 进程
            exit();
        end
        %% ==================== UI 构建（Apple 磨砂玻璃风格） ====================
        function buildUI(app)
            % 跨平台无衬线现代 UI 字体（macOS 苹方 / Windows 微软雅黑）
            if ismac
                uiFont = 'PingFang SC';
            else
                uiFont = 'Microsoft YaHei';
            end

            % Apple 液态玻璃配色体系
            bgMist       = [0.925, 0.945, 0.975];
            glassCard    = [0.988, 0.992, 1.000];
            glassBorder  = [0.835, 0.865, 0.915];
            appleBlue    = [0.00, 0.48, 1.00];
            appleIndigo  = [0.35, 0.34, 0.84];
            appleGreen   = [0.16, 0.67, 0.43];
            textPrimary  = [0.13, 0.16, 0.22];
            textMuted    = [0.44, 0.48, 0.55];

            app.UIFigure = uifigure( ...
                'Name', 'Figure Studio — 科研绘图排版与字体统一工具', ...
                'Position', [70, 50, 1460, 880], ...
                'WindowStyle', 'normal', ...
                'Color', bgMist, ...
                'WindowButtonDownFcn',   @(~,~) app.onMouseDown(), ...
                'WindowButtonMotionFcn', @(~,~) app.onMouseMove(), ...
                'WindowButtonUpFcn',     @(~,~) app.onMouseUp(), ...
                'KeyPressFcn',           @(~,evt) app.onKeyPress(evt));

            % 加载自定义窗口图标（macOS & Windows 通用）
            iconPng = fullfile(fileparts(mfilename('fullpath')), 'assets', 'app_icon.png');
            if isfile(iconPng)
                try, app.UIFigure.Icon = iconPng; catch, end
            end

            % 顶层网格：顶部毛玻璃工具栏 | 中间主工作区 | 底部状态胶囊
            rootGrid = uigridlayout(app.UIFigure, [3, 1]);
            rootGrid.RowHeight = {50, '1x', 30};
            rootGrid.Padding = [12, 10, 12, 8];
            rootGrid.RowSpacing = 10;
            rootGrid.BackgroundColor = bgMist;

            % ---------- 0. 顶部 Apple 风格毛玻璃导航栏 ----------
            topBar = uipanel(rootGrid, 'BorderType', 'line', ...
                'BackgroundColor', glassCard, 'HighlightColor', glassBorder);
            tbGrid = uigridlayout(topBar, [1, 6]);
            tbGrid.ColumnWidth = {'1x', 130, 110, 155, 125, 4};
            tbGrid.Padding = [14, 8, 10, 8];
            tbGrid.ColumnSpacing = 8;
            tbGrid.BackgroundColor = glassCard;

            titleBox = uigridlayout(tbGrid, [1, 2]);
            titleBox.ColumnWidth = {'fit', '1x'};
            titleBox.Padding = [0, 0, 0, 0];
            titleBox.ColumnSpacing = 10;
            titleBox.BackgroundColor = glassCard;
            uilabel(titleBox, 'Text', 'Figure Studio', ...
                'FontName', uiFont, 'FontSize', 16, 'FontWeight', 'bold', ...
                'FontColor', textPrimary);
            uilabel(titleBox, 'Text', '科研绘图自由排版 · Word 1:1 物理字体统一', ...
                'FontName', uiFont, 'FontSize', 12, 'FontColor', textMuted);

            uibutton(tbGrid, 'push', 'Text', '导入 Word', ...
                'FontName', uiFont, 'FontSize', 12, 'FontWeight', 'bold', ...
                'BackgroundColor', [0.91, 0.95, 1.00], 'FontColor', appleBlue, ...
                'ButtonPushedFcn', @(~,~) app.onImportWord());
            uibutton(tbGrid, 'push', 'Text', '1:1 预览', ...
                'FontName', uiFont, 'FontSize', 12, 'FontWeight', 'bold', ...
                'BackgroundColor', [0.94, 0.96, 0.98], 'FontColor', textPrimary, ...
                'ButtonPushedFcn', @(~,~) app.onPreviewComposite());
            uibutton(tbGrid, 'push', 'Text', '属性检查器微调', ...
                'FontName', uiFont, 'FontSize', 12, 'FontWeight', 'bold', ...
                'BackgroundColor', appleIndigo, 'FontColor', 'w', ...
                'ButtonPushedFcn', @(~,~) app.onOpenEditableFigure());
            uibutton(tbGrid, 'push', 'Text', '导出高清大图', ...
                'FontName', uiFont, 'FontSize', 12, 'FontWeight', 'bold', ...
                'BackgroundColor', appleGreen, 'FontColor', 'w', ...
                'ButtonPushedFcn', @(~,~) app.onExportComposite());

            % 主区域三栏：左控制栏 | 中间自由画布 | 右属性栏
            mainGrid = uigridlayout(rootGrid, [1, 3]);
            mainGrid.ColumnWidth = {342, '1x', 252};
            mainGrid.Padding = [0, 0, 0, 0];
            mainGrid.ColumnSpacing = 10;
            mainGrid.BackgroundColor = bgMist;

            % ---------- 1. 左侧可滚动磨砂卡片栏 ----------
            leftScroll = uipanel(mainGrid, 'Scrollable', 'on', ...
                'BorderType', 'none', 'BackgroundColor', bgMist);
            leftGrid = uigridlayout(leftScroll, [4, 1]);
            leftGrid.RowHeight = {'fit', 'fit', 'fit', 'fit'};
            leftGrid.Padding = [0, 0, 8, 8];
            leftGrid.RowSpacing = 10;
            leftGrid.BackgroundColor = bgMist;

            app.buildWordPanel(leftGrid);
            app.buildLayoutPanel(leftGrid);
            app.buildStylePanel(leftGrid);
            app.buildExportPanel(leftGrid);

            % ---------- 2. 中间交互画布卡片 ----------
            centerPanel = uipanel(mainGrid, 'BorderType', 'line', ...
                'BackgroundColor', glassCard, 'HighlightColor', glassBorder);
            cGrid = uigridlayout(centerPanel, [2, 1]);
            cGrid.RowHeight = {24, '1x'};
            cGrid.Padding = [12, 8, 12, 10];
            cGrid.RowSpacing = 6;
            cGrid.BackgroundColor = glassCard;

            uilabel(cGrid, 'Text', '自由排版画布（cm） · 拖动移动 / 圆点缩放 / 双击放入 .fig / Shift 多选', ...
                'FontName', uiFont, 'FontSize', 11.5, 'FontColor', textMuted);

            app.CanvasAxes = uiaxes(cGrid);
            app.CanvasAxes.BackgroundColor = [0.92, 0.94, 0.965];
            app.CanvasAxes.XAxisLocation = 'top';
            app.CanvasAxes.YDir = 'reverse';
            app.CanvasAxes.Box = 'on';
            app.CanvasAxes.DataAspectRatio = [1, 1, 1];
            app.CanvasAxes.FontName = 'Helvetica Neue';
            app.CanvasAxes.FontSize = 10;
            app.CanvasAxes.XColor = [0.42, 0.46, 0.54];
            app.CanvasAxes.YColor = [0.42, 0.46, 0.54];
            app.CanvasAxes.XLabel.String = '宽度 (cm)';
            app.CanvasAxes.YLabel.String = '高度 (cm)';
            app.CanvasAxes.Toolbar.Visible = 'off';
            disableDefaultInteractivity(app.CanvasAxes);

            % ---------- 3. 右侧属性检查器卡片 ----------
            app.buildInspectorPanel(mainGrid);

            % ---------- 4. 底部状态胶囊 ----------
            statusPanel = uipanel(rootGrid, 'BorderType', 'line', ...
                'BackgroundColor', glassCard, 'HighlightColor', glassBorder);
            sGrid = uigridlayout(statusPanel, [1, 1]);
            sGrid.Padding = [12, 3, 12, 3];
            sGrid.BackgroundColor = glassCard;
            app.StatusLabel = uilabel(sGrid, 'Text', '正在初始化...', ...
                'FontName', uiFont, 'FontSize', 11.5, 'FontColor', [0.28, 0.33, 0.42]);
        end

        function buildWordPanel(app, parent)
            p = uipanel(parent, 'Title', '  Word 版心与画布', ...
                'FontWeight', 'bold', 'FontSize', 12, ...
                'ForegroundColor', [0.14, 0.18, 0.26], ...
                'BackgroundColor', [0.988, 0.992, 1.000], ...
                'HighlightColor',  [0.835, 0.865, 0.915]);
            g = uigridlayout(p, [5, 2]);
            g.RowHeight = {34, 'fit', 28, 28, 28};
            g.ColumnWidth = {'1x', '1x'};
            g.Padding = [10, 10, 10, 10];
            g.RowSpacing = 7;
            g.ColumnSpacing = 8;
            g.BackgroundColor = [0.988, 0.992, 1.000];

            btnWord = uibutton(g, 'push', 'Text', '导入 Word (.docx) 自动识别版心', ...
                'FontWeight', 'bold', 'BackgroundColor', [0.00, 0.48, 1.00], ...
                'FontColor', 'w', 'ButtonPushedFcn', @(~,~) app.onImportWord());
            btnWord.Layout.Row = 1;
            btnWord.Layout.Column = [1, 2];

            app.WordSummaryLabel = uilabel(g, ...
                'Text', '未导入 Word（当前默认 A4 单栏版心 14.65 cm）', ...
                'WordWrap', 'on', 'FontSize', 11, 'FontColor', [0.44, 0.48, 0.55]);
            app.WordSummaryLabel.Layout.Row = 2;
            app.WordSummaryLabel.Layout.Column = [1, 2];

            uilabel(g, 'Text', '栏宽模式:', 'FontColor', [0.24, 0.28, 0.35]);
            app.ColModeDropDown = uidropdown(g, ...
                'Items', {'通栏宽 (Full)', '单栏宽 (Column)'}, ...
                'ItemsData', {'full', 'column'}, ...
                'Value', 'full', 'Enable', 'off', ...
                'BackgroundColor', [0.96, 0.97, 0.99], ...
                'ValueChangedFcn', @(~,~) app.onColModeChanged());

            % 画布宽高输入行
            wGrid = uigridlayout(g, [1, 2]);
            wGrid.Layout.Row = 4; wGrid.Layout.Column = 1;
            wGrid.Padding = [0 0 0 0]; wGrid.ColumnWidth = {'fit', '1x'};
            wGrid.BackgroundColor = [0.988, 0.992, 1.000];
            uilabel(wGrid, 'Text', '宽(cm):', 'FontColor', [0.24, 0.28, 0.35]);
            app.CanvasWField = uieditfield(wGrid, 'numeric', ...
                'Value', app.CanvasWidthCm, 'Limits', [3, 60], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onCanvasSizeChanged());

            hGrid = uigridlayout(g, [1, 2]);
            hGrid.Layout.Row = 4; hGrid.Layout.Column = 2;
            hGrid.Padding = [0 0 0 0]; hGrid.ColumnWidth = {'fit', '1x'};
            hGrid.BackgroundColor = [0.988, 0.992, 1.000];
            uilabel(hGrid, 'Text', '高(cm):', 'FontColor', [0.24, 0.28, 0.35]);
            app.CanvasHField = uieditfield(hGrid, 'numeric', ...
                'Value', app.CanvasHeightCm, 'Limits', [3, 80], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onCanvasSizeChanged());

            % 列间距 / 行间距（修改后立即重排并放大铺满）
            gxGrid = uigridlayout(g, [1, 2]);
            gxGrid.Layout.Row = 5; gxGrid.Layout.Column = 1;
            gxGrid.Padding = [0 0 0 0]; gxGrid.ColumnWidth = {'fit', '1x'};
            gxGrid.BackgroundColor = [0.988, 0.992, 1.000];
            uilabel(gxGrid, 'Text', '列间距:', 'FontColor', [0.24, 0.28, 0.35]);
            app.GapXField = uieditfield(gxGrid, 'numeric', ...
                'Value', app.GapXCm, 'Limits', [0, 3], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onGapChanged());

            gyGrid = uigridlayout(g, [1, 2]);
            gyGrid.Layout.Row = 5; gyGrid.Layout.Column = 2;
            gyGrid.Padding = [0 0 0 0]; gyGrid.ColumnWidth = {'fit', '1x'};
            gyGrid.BackgroundColor = [0.988, 0.992, 1.000];
            uilabel(gyGrid, 'Text', '行间距:', 'FontColor', [0.24, 0.28, 0.35]);
            app.GapYField = uieditfield(gyGrid, 'numeric', ...
                'Value', app.GapYCm, 'Limits', [0, 3], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onGapChanged());
        end

        function buildLayoutPanel(app, parent)
            glassCard = [0.988, 0.992, 1.000];
            pillBg    = [0.945, 0.960, 0.982];
            pillTxt   = [0.18, 0.22, 0.30];

            p = uipanel(parent, 'Title', '  自由排版与对齐', ...
                'FontWeight', 'bold', 'FontSize', 12, ...
                'ForegroundColor', [0.14, 0.18, 0.26], ...
                'BackgroundColor', glassCard, ...
                'HighlightColor',  [0.835, 0.865, 0.915]);
            g = uigridlayout(p, [7, 4]);
            g.RowHeight = {28, 28, 26, 26, 30, 28, 24};
            g.ColumnWidth = {'1x', '1x', '1x', '1x'};
            g.Padding = [10, 10, 10, 10];
            g.RowSpacing = 6;
            g.ColumnSpacing = 6;
            g.BackgroundColor = glassCard;

            % 预设模板行
            app.PresetDropDown = uidropdown(g, ...
                'Items', {'1×1 单图铺满', '1×2 横排', '1×3 横排', '1×4 横排', '2×2 网格', ...
                          '3+2 (上3下2)', '2+3 (上2下3)', '2×3 网格', '1+2 (左1右2)'}, ...
                'ItemsData', {'1x1', '1x2', '1x3', '1x4', '2x2', '3+2', '2+3', '2x3', '1+2_LR'}, ...
                'Value', '3+2', 'BackgroundColor', [0.96, 0.97, 0.99]);
            app.PresetDropDown.Layout.Row = 1;
            app.PresetDropDown.Layout.Column = [1, 3];

            btnApply = uibutton(g, 'push', 'Text', '套用', ...
                'FontWeight', 'bold', 'BackgroundColor', [0.90, 0.95, 1.00], ...
                'FontColor', [0.00, 0.42, 0.90], ...
                'ButtonPushedFcn', @(~,~) app.applyPreset(app.PresetDropDown.Value));
            btnApply.Layout.Row = 1;
            btnApply.Layout.Column = 4;

            % 增删操作行
            uibutton(g, 'push', 'Text', '+ 添加', ...
                'BackgroundColor', [0.90, 0.95, 1.00], 'FontColor', [0.00, 0.42, 0.90], ...
                'ButtonPushedFcn', @(~,~) app.onAddSlot());
            uibutton(g, 'push', 'Text', '复制', ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onDuplicateSelected());
            uibutton(g, 'push', 'Text', '删除', ...
                'BackgroundColor', [0.99, 0.93, 0.93], 'FontColor', [0.80, 0.22, 0.22], ...
                'ButtonPushedFcn', @(~,~) app.onDeleteSelected());
            uibutton(g, 'push', 'Text', '清空', ...
                'BackgroundColor', [0.99, 0.93, 0.93], 'FontColor', [0.80, 0.22, 0.22], ...
                'ButtonPushedFcn', @(~,~) app.onClearAll());

            % 对齐第 1 排
            uibutton(g, 'push', 'Text', '左对齐', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('left'));
            uibutton(g, 'push', 'Text', '右对齐', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('right'));
            uibutton(g, 'push', 'Text', '顶对齐', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('top'));
            uibutton(g, 'push', 'Text', '底对齐', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('bottom'));

            % 对齐第 2 排
            uibutton(g, 'push', 'Text', '等宽', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('same_w'));
            uibutton(g, 'push', 'Text', '等高', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('same_h'));
            uibutton(g, 'push', 'Text', '水平均分', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('dist_h'));
            uibutton(g, 'push', 'Text', '垂直均分', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.alignSelected('dist_v'));

            % 一键紧凑铺满版心按钮
            btnCompact = uibutton(g, 'push', 'Text', '一键紧凑靠拢并放大铺满版心宽', ...
                'FontWeight', 'bold', 'BackgroundColor', [0.90, 0.95, 1.00], ...
                'FontColor', [0.00, 0.42, 0.90], ...
                'ButtonPushedFcn', @(~,~) app.onCompactFillLayout(true));
            btnCompact.Layout.Row = 5;
            btnCompact.Layout.Column = [1, 4];

            % 批量填图行
            btnBatch = uibutton(g, 'push', 'Text', '批量选择多个 .fig 按顺序填入图框', ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onBatchLoadFigs());
            btnBatch.Layout.Row = 6;
            btnBatch.Layout.Column = [1, 4];

            % 吸附开关
            app.SnapCheck = uicheckbox(g, 'Text', '拖拽吸附网格与对齐参考线 (0.2 cm)', ...
                'FontColor', [0.28, 0.33, 0.40], ...
                'Value', true, 'ValueChangedFcn', @(src,~) app.onSnapChanged(src.Value));
            app.SnapCheck.Layout.Row = 7;
            app.SnapCheck.Layout.Column = [1, 4];
        end

        function buildStylePanel(app, parent)
            glassCard = [0.988, 0.992, 1.000];
            lblCol    = [0.24, 0.28, 0.35];

            p = uipanel(parent, 'Title', '  统一字体与科研外观', ...
                'FontWeight', 'bold', 'FontSize', 12, ...
                'ForegroundColor', [0.14, 0.18, 0.26], ...
                'BackgroundColor', glassCard, ...
                'HighlightColor',  [0.835, 0.865, 0.915]);
            g = uigridlayout(p, [11, 4]);
            g.RowHeight = {28, 28, 26, 26, 26, 24, 26, 26, 24, 24, 24};
            g.ColumnWidth = {'fit', '1x', 'fit', '1x'};
            g.Padding = [10, 10, 10, 10];
            g.RowSpacing = 5;
            g.ColumnSpacing = 6;
            g.BackgroundColor = glassCard;

            % Row 1: 字体选择
            uilabel(g, 'Text', '字体:', 'FontColor', lblCol);
            commonFonts = {'Times New Roman', 'Arial', 'Helvetica', 'Calibri', ...
                           'SimSun', 'Songti SC', 'PingFang SC', 'Microsoft YaHei', 'STHeiti'};
            try
                sysFonts = listfonts();
                allFonts = unique([commonFonts(:); sysFonts(:)], 'stable');
            catch
                allFonts = commonFonts;
            end
            app.FontNameDropDown = uidropdown(g, 'Items', allFonts, ...
                'Value', 'Times New Roman', 'Editable', 'on', ...
                'BackgroundColor', [0.96, 0.97, 0.99]);
            app.FontNameDropDown.Layout.Row = 1;
            app.FontNameDropDown.Layout.Column = [2, 4];

            % Row 2: 全局统一字号 + 一键同步
            lblG = uilabel(g, 'Text', '全局字号:', 'FontWeight', 'bold', 'FontColor', [0.14, 0.18, 0.26]);
            lblG.Layout.Row = 2; lblG.Layout.Column = 1;
            app.GlobalFontField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f', ...
                'ValueChangedFcn', @(~,~) app.onGlobalFontChanged());
            app.GlobalFontField.Layout.Row = 2; app.GlobalFontField.Layout.Column = 2;

            btnSync = uibutton(g, 'push', 'Text', '同步铺满并刷新', ...
                'FontSize', 11, 'FontWeight', 'bold', ...
                'BackgroundColor', [0.90, 0.95, 1.00], 'FontColor', [0.00, 0.42, 0.90], ...
                'ButtonPushedFcn', @(~,~) app.onSyncAndRefreshThumbs());
            btnSync.Layout.Row = 2; btnSync.Layout.Column = [3, 4];

            % Row 3: 刻度 & 轴标签
            uilabel(g, 'Text', '刻度:', 'FontColor', lblCol);
            app.TickFontField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f');
            uilabel(g, 'Text', '轴标签:', 'FontColor', lblCol);
            app.LabelFontField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f');

            % Row 4: 标题 & 图例
            uilabel(g, 'Text', '标题:', 'FontColor', lblCol);
            app.TitleFontField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f');
            uilabel(g, 'Text', '图例:', 'FontColor', lblCol);
            app.LegendFontField = uieditfield(g, 'numeric', 'Value', 9.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f');

            % Row 5: 图内文本标注 & 紧凑边距
            uilabel(g, 'Text', '标注:', 'FontColor', lblCol);
            app.TextFontField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [4, 36], 'ValueDisplayFormat', '%.1f');
            app.TightLayoutCheck = uicheckbox(g, 'Text', '防裁切紧凑边距', 'Value', true, 'FontColor', lblCol);
            app.TightLayoutCheck.Layout.Row = 5;
            app.TightLayoutCheck.Layout.Column = [3, 4];

            % Row 6: 取消原图粗体（科研论文标准常规体）
            app.ForceNormalCheck = uicheckbox(g, 'Text', '统一为常规字重（取消原图过粗字体）', ...
                'Value', true, 'FontColor', lblCol);
            app.ForceNormalCheck.Layout.Row = 6;
            app.ForceNormalCheck.Layout.Column = [1, 4];

            % Row 7: 子图编号开关 & 样式
            app.SubLabelCheck = uicheckbox(g, 'Text', '子图编号', 'Value', true, 'FontColor', lblCol, ...
                'ValueChangedFcn', @(~,~) app.redrawCanvas());
            app.SubLabelCheck.Layout.Row = 7; app.SubLabelCheck.Layout.Column = [1, 2];

            app.SubLabelStyleDD = uidropdown(g, ...
                'Items', {'(a)', '(A)', 'a', 'A', 'a)'}, 'Value', '(a)', ...
                'BackgroundColor', [0.96, 0.97, 0.99], ...
                'ValueChangedFcn', @(~,~) app.redrawCanvas());
            app.SubLabelStyleDD.Layout.Row = 7; app.SubLabelStyleDD.Layout.Column = [3, 4];

            % Row 8: 编号位置 & 字号
            uilabel(g, 'Text', '位置:', 'FontColor', lblCol);
            app.SubLabelPosDD = uidropdown(g, ...
                'Items', {'左上独立行', '底部居中', '左上角内侧'}, ...
                'ItemsData', {'top-left-above', 'bottom-center', 'top-left-inside'}, ...
                'Value', 'top-left-above', 'BackgroundColor', [0.96, 0.97, 0.99]);
            uilabel(g, 'Text', '字号:', 'FontColor', lblCol);
            app.SubLabelSizeField = uieditfield(g, 'numeric', 'Value', 10.0, ...
                'Limits', [5, 36], 'ValueDisplayFormat', '%.1f');

            % Row 9: 编号加粗
            app.SubLabelBoldCheck = uicheckbox(g, 'Text', '子图编号加粗 (Bold)', 'Value', true, 'FontColor', lblCol);
            app.SubLabelBoldCheck.Layout.Row = 9; app.SubLabelBoldCheck.Layout.Column = [1, 4];

            % Row 10: 统一坐标轴线宽
            app.AxisLWCheck = uicheckbox(g, 'Text', '统一坐标轴线宽(pt):', 'Value', true, 'FontColor', lblCol);
            app.AxisLWCheck.Layout.Row = 10; app.AxisLWCheck.Layout.Column = [1, 3];
            app.AxisLWField = uieditfield(g, 'numeric', 'Value', 0.75, ...
                'Limits', [0.25, 4], 'ValueDisplayFormat', '%.2f');

            % Row 11: 统一曲线线宽
            app.PlotLWCheck = uicheckbox(g, 'Text', '统一数据曲线线宽(pt):', 'Value', true, 'FontColor', lblCol);
            app.PlotLWCheck.Layout.Row = 11; app.PlotLWCheck.Layout.Column = [1, 3];
            app.PlotLWField = uieditfield(g, 'numeric', 'Value', 1.20, ...
                'Limits', [0.25, 6], 'ValueDisplayFormat', '%.2f');
        end

        function buildExportPanel(app, parent)
            glassCard = [0.988, 0.992, 1.000];
            pillBg    = [0.945, 0.960, 0.982];
            pillTxt   = [0.18, 0.22, 0.30];

            p = uipanel(parent, 'Title', '  导出与方案存取', ...
                'FontWeight', 'bold', 'FontSize', 12, ...
                'ForegroundColor', [0.14, 0.18, 0.26], ...
                'BackgroundColor', glassCard, ...
                'HighlightColor',  [0.835, 0.865, 0.915]);
            g = uigridlayout(p, [6, 2]);
            g.RowHeight = {28, 24, 24, 34, 34, 28};
            g.ColumnWidth = {'1x', '1x'};
            g.Padding = [10, 10, 10, 10];
            g.RowSpacing = 6;
            g.ColumnSpacing = 8;
            g.BackgroundColor = glassCard;

            app.FormatDropDown = uidropdown(g, ...
                'Items', {'PNG 高清位图 (.png)', 'TIFF 无损印刷 (.tif)', 'MATLAB 可编辑图 (.fig)'}, ...
                'ItemsData', {'.png', '.tif', '.fig'}, 'Value', '.png', ...
                'BackgroundColor', [0.96, 0.97, 0.99]);
            app.DpiDropDown = uidropdown(g, ...
                'Items', {'300 DPI (期刊标准)', '600 DPI (超清推荐)', '1200 DPI (极致印刷)'}, ...
                'ItemsData', [300, 600, 1200], 'Value', 600, ...
                'BackgroundColor', [0.96, 0.97, 0.99]);

            app.AutoCropCheck = uicheckbox(g, 'Text', '导出时自动裁剪底部多余空白', ...
                'Value', true, 'FontColor', [0.24, 0.28, 0.35]);
            app.AutoCropCheck.Layout.Row = 2; app.AutoCropCheck.Layout.Column = [1, 2];

            app.SaveSubCheck = uicheckbox(g, 'Text', '同时保存各子图单独高清文件', ...
                'Value', false, 'FontColor', [0.24, 0.28, 0.35]);
            app.SaveSubCheck.Layout.Row = 3; app.SaveSubCheck.Layout.Column = [1, 2];

            uibutton(g, 'push', 'Text', '1:1 渲染预览', ...
                'FontWeight', 'bold', 'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onPreviewComposite());
            uibutton(g, 'push', 'Text', '导出高清大图', ...
                'FontWeight', 'bold', 'BackgroundColor', [0.16, 0.67, 0.43], ...
                'FontColor', 'w', 'ButtonPushedFcn', @(~,~) app.onExportComposite());

            btnEditFig = uibutton(g, 'push', 'Text', '在 MATLAB 中打开微调（属性检查器）', ...
                'FontWeight', 'bold', 'BackgroundColor', [0.35, 0.34, 0.84], ...
                'FontColor', 'w', 'ButtonPushedFcn', @(~,~) app.onOpenEditableFigure());
            btnEditFig.Layout.Row = 5; btnEditFig.Layout.Column = [1, 2];

            uibutton(g, 'push', 'Text', '保存方案 (.json)', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onSaveProject());
            uibutton(g, 'push', 'Text', '加载方案 (.json)', 'FontSize', 11, ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onLoadProject());
        end

        function buildInspectorPanel(app, parent)
            glassCard = [0.988, 0.992, 1.000];
            pillBg    = [0.945, 0.960, 0.982];
            pillTxt   = [0.18, 0.22, 0.30];

            p = uipanel(parent, 'Title', '  图框属性与列表', ...
                'FontWeight', 'bold', 'FontSize', 12, ...
                'ForegroundColor', [0.14, 0.18, 0.26], ...
                'BackgroundColor', glassCard, ...
                'HighlightColor',  [0.835, 0.865, 0.915]);
            g = uigridlayout(p, [11, 2]);
            g.RowHeight = {20, 195, 22, 28, 28, 28, 36, 32, 28, 28, '1x'};
            g.ColumnWidth = {'1x', '1x'};
            g.Padding = [10, 10, 10, 10];
            g.RowSpacing = 6;
            g.ColumnSpacing = 6;
            g.BackgroundColor = glassCard;

            lblList = uilabel(g, 'Text', '所有图框（点击选中）', 'FontSize', 11.5, ...
                'FontWeight', 'bold', 'FontColor', [0.32, 0.36, 0.44]);
            lblList.Layout.Row = 1; lblList.Layout.Column = [1, 2];

            app.SlotListBox = uilistbox(g, 'Items', {}, 'Multiselect', 'on', ...
                'BackgroundColor', [0.965, 0.975, 0.992], ...
                'ValueChangedFcn', @(~,~) app.onListBoxSelectionChanged());
            app.SlotListBox.Layout.Row = 2; app.SlotListBox.Layout.Column = [1, 2];

            app.InspectorInfoLabel = uilabel(g, 'Text', '未选中图框', ...
                'FontWeight', 'bold', 'FontColor', [0.00, 0.45, 0.95]);
            app.InspectorInfoLabel.Layout.Row = 3; app.InspectorInfoLabel.Layout.Column = [1, 2];

            % X / Y
            xG = uigridlayout(g, [1, 2]); xG.Padding = [0 0 0 0]; xG.ColumnWidth = {26, '1x'}; xG.BackgroundColor = glassCard;
            uilabel(xG, 'Text', 'X:', 'FontColor', [0.24, 0.28, 0.35]);
            app.SlotXField = uieditfield(xG, 'numeric', 'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onInspectorGeomChanged());

            yG = uigridlayout(g, [1, 2]); yG.Padding = [0 0 0 0]; yG.ColumnWidth = {26, '1x'}; yG.BackgroundColor = glassCard;
            uilabel(yG, 'Text', 'Y:', 'FontColor', [0.24, 0.28, 0.35]);
            app.SlotYField = uieditfield(yG, 'numeric', 'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onInspectorGeomChanged());

            % W / H
            wG = uigridlayout(g, [1, 2]); wG.Padding = [0 0 0 0]; wG.ColumnWidth = {26, '1x'}; wG.BackgroundColor = glassCard;
            uilabel(wG, 'Text', '宽:', 'FontColor', [0.24, 0.28, 0.35]);
            app.SlotWField = uieditfield(wG, 'numeric', 'Limits', [1.0, 60], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onInspectorGeomChanged());

            hG = uigridlayout(g, [1, 2]); hG.Padding = [0 0 0 0]; hG.ColumnWidth = {26, '1x'}; hG.BackgroundColor = glassCard;
            uilabel(hG, 'Text', '高:', 'FontColor', [0.24, 0.28, 0.35]);
            app.SlotHField = uieditfield(hG, 'numeric', 'Limits', [1.0, 60], ...
                'ValueDisplayFormat', '%.2f', ...
                'ValueChangedFcn', @(~,~) app.onInspectorGeomChanged());

            % 自定义子图编号
            lblCustom = uilabel(g, 'Text', '编号覆写:', 'FontColor', [0.24, 0.28, 0.35]);
            lblCustom.Layout.Row = 6; lblCustom.Layout.Column = 1;
            app.SlotLabelEdit = uieditfield(g, 'text', 'Placeholder', '留空自动编号', ...
                'ValueChangedFcn', @(~,~) app.onSlotLabelChanged());
            app.SlotLabelEdit.Layout.Row = 6; app.SlotLabelEdit.Layout.Column = 2;

            % 文件信息
            app.SlotFileLabel = uilabel(g, 'Text', '尚未载入 .fig 文件', ...
                'WordWrap', 'on', 'FontSize', 11, 'FontColor', [0.44, 0.48, 0.55]);
            app.SlotFileLabel.Layout.Row = 7; app.SlotFileLabel.Layout.Column = [1, 2];

            app.LoadFigBtn = uibutton(g, 'push', 'Text', '载入 .fig / 图片', ...
                'BackgroundColor', [0.00, 0.48, 1.00], 'FontColor', 'w', ...
                'FontWeight', 'bold', 'ButtonPushedFcn', @(~,~) app.onLoadSingleFig());
            app.LoadFigBtn.Layout.Row = 8; app.LoadFigBtn.Layout.Column = [1, 2];

            app.RefreshThumbBtn = uibutton(g, 'push', 'Text', '刷新预览', ...
                'BackgroundColor', pillBg, 'FontColor', pillTxt, ...
                'ButtonPushedFcn', @(~,~) app.onRefreshSelectedThumb());
            app.ClearFigBtn = uibutton(g, 'push', 'Text', '清除图片', ...
                'BackgroundColor', [0.99, 0.93, 0.93], 'FontColor', [0.80, 0.22, 0.22], ...
                'ButtonPushedFcn', @(~,~) app.onClearSelectedFig());

            app.FullWidthBtn = uibutton(g, 'push', 'Text', '横向铺满版心宽', ...
                'BackgroundColor', [0.90, 0.95, 1.00], 'FontColor', [0.00, 0.42, 0.90], ...
                'ButtonPushedFcn', @(~,~) app.onStretchFullWidth());
            app.FullWidthBtn.Layout.Row = 10; app.FullWidthBtn.Layout.Column = [1, 2];
        end

        %% ==================== Word 导入与画布尺寸 ====================
        function onImportWord(app)
            [file, path] = uigetfile({'*.docx', 'Word 文档 (*.docx)'}, '选择论文 Word 文档');
            if isequal(file, 0), return; end
            docxPath = fullfile(path, file);

            try
                info = parseWordDocx(docxPath);
                app.WordInfo = info;
                app.WordSummaryLabel.Text = info.summary;

                % 自动设置画布宽度（以及高度至少略大于宽度）
                app.CanvasWidthCm = info.fullTextWidthCm;
                if app.CanvasHeightCm < app.CanvasWidthCm
                    app.CanvasHeightCm = round(app.CanvasWidthCm * 1.25, 2);
                end
                app.CanvasWField.Value = app.CanvasWidthCm;
                app.CanvasHField.Value = app.CanvasHeightCm;

                % 分栏切换开关
                if info.numColumns > 1
                    app.ColModeDropDown.Enable = 'on';
                    app.ColModeDropDown.Value = 'full';
                else
                    app.ColModeDropDown.Enable = 'off';
                    app.ColModeDropDown.Value = 'full';
                end

                % 自动同步 Word 识别出的正文字号与字体
                app.GlobalFontField.Value = info.bodyFontSizePt;
                app.onGlobalFontChanged();

                if ~isempty(info.bodyFontName)
                    if ~any(strcmp(app.FontNameDropDown.Items, info.bodyFontName))
                        app.FontNameDropDown.Items = [{info.bodyFontName}, app.FontNameDropDown.Items];
                    end
                    app.FontNameDropDown.Value = info.bodyFontName;
                end

                % 按新宽度等比缩放现有图框，避免超出画布
                app.fitSlotsToCanvasWidth();
                app.redrawCanvas();
                app.syncInspector();
                app.setStatus(sprintf('已解析 Word：%s', info.summary));
            catch ME
                uialert(app.UIFigure, ME.message, 'Word 解析失败');
            end
        end

        function onColModeChanged(app)
            if ~isfield(app.WordInfo, 'fullTextWidthCm'), return; end
            if strcmp(app.ColModeDropDown.Value, 'column')
                app.CanvasWidthCm = app.WordInfo.columnWidthCm;
            else
                app.CanvasWidthCm = app.WordInfo.fullTextWidthCm;
            end
            app.CanvasWField.Value = app.CanvasWidthCm;
            app.fitSlotsToCanvasWidth();
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus(sprintf('已切换画布宽度为 %.2f cm', app.CanvasWidthCm));
        end

        function onCanvasSizeChanged(app)
            oldW = app.CanvasWidthCm;
            app.CanvasWidthCm  = app.CanvasWField.Value;
            app.CanvasHeightCm = app.CanvasHField.Value;
            if oldW > 0 && abs(app.CanvasWidthCm - oldW) > 1e-3 && ~isempty(app.Slots)
                % 若有图框超出新宽度，按比例适配
                maxRight = max(arrayfun(@(s) s.x + s.w, app.Slots));
                if maxRight > app.CanvasWidthCm + 0.05
                    app.fitSlotsToCanvasWidth();
                end
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function fitSlotsToCanvasWidth(app)
            if isempty(app.Slots), return; end
            maxRight = max(arrayfun(@(s) s.x + s.w, app.Slots));
            if maxRight <= 0, return; end
            scale = app.CanvasWidthCm / maxRight;
            for i = 1:numel(app.Slots)
                app.Slots(i).x = round(app.Slots(i).x * scale, 2);
                app.Slots(i).w = max(1.2, round(app.Slots(i).w * scale, 2));
            end
        end

        %% ==================== 排版模板 & 自由框操作 ====================
        function applyPreset(app, presetName)
            % 保留已绑定的 .fig 路径，尽量无缝迁移到新模板
            oldFiles = {};
            if ~isempty(app.Slots)
                oldFiles = {app.Slots.filePath};
                oldFiles = oldFiles(~cellfun(@isempty, oldFiles));
            end

            newSlots = getLayoutPresets(presetName, app.CanvasWidthCm, [app.GapXCm, app.GapYCm], 0.0);
            for i = 1:numel(newSlots)
                newSlots(i).id = i;
                if i <= numel(oldFiles)
                    newSlots(i).filePath = oldFiles{i};
                end
            end
            app.Slots = newSlots;
            app.NextSlotId = numel(newSlots) + 1;

            % 若模板总高度超过当前画布高度，自动拉长画布高度
            if ~isempty(app.Slots)
                maxBottom = max(arrayfun(@(s) s.y + s.h, app.Slots));
                if maxBottom + 1.0 > app.CanvasHeightCm
                    app.CanvasHeightCm = round(maxBottom + 2.0, 2);
                    app.CanvasHField.Value = app.CanvasHeightCm;
                end
                app.SelectedIds = app.Slots(1).id;
            else
                app.SelectedIds = [];
            end

            % 为迁移过来的 .fig 重新生成缩略图
            for i = 1:numel(app.Slots)
                if ~isempty(app.Slots(i).filePath) && isfile(app.Slots(i).filePath)
                    app.updateSlotThumbnail(i);
                end
            end

            app.redrawCanvas();
            app.syncInspector();
            app.setStatus(sprintf('已应用排版模板 [%s]，共 %d 个图框。', presetName, numel(app.Slots)));
        end

        function onGapChanged(app)
            app.GapXCm = app.GapXField.Value;
            app.GapYCm = app.GapYField.Value;
            % 修改列间距或行间距后，立即紧凑重排当前画布上的所有图框并铺满版心宽
            app.onCompactFillLayout(true);
        end

        function onCompactFillLayout(app, refreshThumbs)
            if nargin < 2, refreshThumbs = true; end
            if isempty(app.Slots), return; end

            n = numel(app.Slots);
            % 按垂直中心 y 聚成行（纵坐标差距在半个行高以内视为同一行）
            yCenters = arrayfun(@(s) s.y + s.h / 2, app.Slots);
            avgH = mean(arrayfun(@(s) s.h, app.Slots));
            rowTol = max(0.8, avgH * 0.48);

            [~, sortByY] = sort(yCenters);
            rows = {};
            curRow = sortByY(1);
            curRefY = yCenters(sortByY(1));
            for k = 2:n
                idx = sortByY(k);
                if abs(yCenters(idx) - curRefY) <= rowTol
                    curRow(end+1) = idx; %#ok<AGROW>
                else
                    rows{end+1} = curRow; %#ok<AGROW>
                    curRow = idx;
                    curRefY = yCenters(idx);
                end
            end
            rows{end+1} = curRow;

            % 逐行从 y=0 开始紧密排列，并将每行宽度按原比例放大至 100% 铺满版心宽
            curY = 0.0;
            for r = 1:numel(rows)
                rIdxs = rows{r};
                xs = arrayfun(@(s) s.x, app.Slots(rIdxs));
                [~, ordX] = sort(xs);
                rIdxs = rIdxs(ordX);

                m = numel(rIdxs);
                origWs = arrayfun(@(s) max(0.5, s.w), app.Slots(rIdxs));
                origHs = arrayfun(@(s) max(0.5, s.h), app.Slots(rIdxs));
                availW = max(1.0, app.CanvasWidthCm - (m - 1) * app.GapXCm);
                scaleW = availW / sum(origWs);

                curX = 0.0;
                rowH = mean(origHs); % 同行统一高度，整齐美观
                for c = 1:m
                    idx = rIdxs(c);
                    newW = origWs(c) * scaleW;
                    app.Slots(idx).x = round(curX, 2);
                    app.Slots(idx).y = round(curY, 2);
                    app.Slots(idx).w = round(newW, 2);
                    app.Slots(idx).h = round(rowH, 2);
                    curX = curX + newW + app.GapXCm;
                end
                % 最后一列精确贴齐右边界，消除舍入误差
                lastIdx = rIdxs(end);
                app.Slots(lastIdx).w = round(max(1.0, app.CanvasWidthCm - app.Slots(lastIdx).x), 2);

                curY = curY + rowH + app.GapYCm;
            end

            if refreshThumbs
                for i = 1:n
                    if ~isempty(app.Slots(i).filePath) && isfile(app.Slots(i).filePath)
                        app.updateSlotThumbnail(i);
                    end
                end
            end
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus(sprintf('已按列间距 %.2f cm / 行间距 %.2f cm 紧凑靠拢并铺满版心宽 (%.2f cm)。', ...
                app.GapXCm, app.GapYCm, app.CanvasWidthCm));
        end

        function onAddSlot(app)
            w = round(min(app.CanvasWidthCm * 0.45, 6.0), 2);
            h = round(w * 0.75, 2);
            offset = mod(numel(app.Slots), 6) * 0.6;
            x = round(min(app.CanvasWidthCm - w, 0.4 + offset), 2);
            y = round(min(app.CanvasHeightCm - h, 0.4 + offset), 2);

            s = struct('id', app.NextSlotId, 'x', max(0, x), 'y', max(0, y), ...
                'w', w, 'h', h, 'filePath', '', 'label', '', 'thumbImg', []);
            app.NextSlotId = app.NextSlotId + 1;
            if isempty(app.Slots)
                app.Slots = s;
            else
                app.Slots(end+1) = s;
            end
            app.SelectedIds = s.id;
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus('已添加新图框，可自由拖动/缩放或双击载入 .fig。');
        end

        function onDuplicateSelected(app)
            idxs = app.getSelectedIndices();
            if isempty(idxs), return; end
            newSel = [];
            for k = 1:numel(idxs)
                src = app.Slots(idxs(k));
                s = src;
                s.id = app.NextSlotId;
                app.NextSlotId = app.NextSlotId + 1;
                s.x = min(app.CanvasWidthCm - s.w, round(src.x + 0.5, 2));
                s.y = min(app.CanvasHeightCm - s.h, round(src.y + 0.5, 2));
                s.label = '';
                app.Slots(end+1) = s; %#ok<AGROW>
                newSel(end+1) = s.id; %#ok<AGROW>
            end
            app.SelectedIds = newSel;
            app.redrawCanvas();
            app.syncInspector();
        end

        function onDeleteSelected(app)
            idxs = app.getSelectedIndices();
            if isempty(idxs), return; end
            app.Slots(idxs) = [];
            app.SelectedIds = [];
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus('已删除选中的图框。');
        end

        function onClearAll(app)
            app.Slots = struct([]);
            app.SelectedIds = [];
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus('已清空画布。');
        end

        function alignSelected(app, mode)
            idxs = app.getSelectedIndices();
            if isempty(idxs), return; end

            if numel(idxs) == 1
                % 单选时相对画布边界对齐
                i = idxs(1);
                switch mode
                    case 'left',   app.Slots(i).x = 0;
                    case 'right',  app.Slots(i).x = round(app.CanvasWidthCm - app.Slots(i).w, 2);
                    case 'top',    app.Slots(i).y = 0;
                    case 'bottom', app.Slots(i).y = round(app.CanvasHeightCm - app.Slots(i).h, 2);
                end
            else
                xs = arrayfun(@(s) s.x, app.Slots(idxs));
                ys = arrayfun(@(s) s.y, app.Slots(idxs));
                ws = arrayfun(@(s) s.w, app.Slots(idxs));
                hs = arrayfun(@(s) s.h, app.Slots(idxs));

                switch mode
                    case 'left'
                        v = min(xs);
                        for k = idxs, app.Slots(k).x = round(v, 2); end
                    case 'right'
                        v = max(xs + ws);
                        for k = idxs, app.Slots(k).x = round(v - app.Slots(k).w, 2); end
                    case 'top'
                        v = min(ys);
                        for k = idxs, app.Slots(k).y = round(v, 2); end
                    case 'bottom'
                        v = max(ys + hs);
                        for k = idxs, app.Slots(k).y = round(v - app.Slots(k).h, 2); end
                    case 'same_w'
                        v = app.Slots(idxs(1)).w;
                        for k = idxs, app.Slots(k).w = round(v, 2); end
                    case 'same_h'
                        v = app.Slots(idxs(1)).h;
                        for k = idxs, app.Slots(k).h = round(v, 2); end
                    case 'dist_h'
                        if numel(idxs) >= 3
                            [~, ord] = sort(xs);
                            sortedIdx = idxs(ord);
                            leftEdge = app.Slots(sortedIdx(1)).x;
                            rightEdge = app.Slots(sortedIdx(end)).x + app.Slots(sortedIdx(end)).w;
                            totalW = sum(arrayfun(@(s) s.w, app.Slots(sortedIdx)));
                            gap = (rightEdge - leftEdge - totalW) / (numel(sortedIdx) - 1);
                            curX = leftEdge;
                            for m = 1:numel(sortedIdx)
                                app.Slots(sortedIdx(m)).x = round(curX, 2);
                                curX = curX + app.Slots(sortedIdx(m)).w + gap;
                            end
                        end
                    case 'dist_v'
                        if numel(idxs) >= 3
                            [~, ord] = sort(ys);
                            sortedIdx = idxs(ord);
                            topEdge = app.Slots(sortedIdx(1)).y;
                            botEdge = app.Slots(sortedIdx(end)).y + app.Slots(sortedIdx(end)).h;
                            totalH = sum(arrayfun(@(s) s.h, app.Slots(sortedIdx)));
                            gap = (botEdge - topEdge - totalH) / (numel(sortedIdx) - 1);
                            curY = topEdge;
                            for m = 1:numel(sortedIdx)
                                app.Slots(sortedIdx(m)).y = round(curY, 2);
                                curY = curY + app.Slots(sortedIdx(m)).h + gap;
                            end
                        end
                end
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function onSnapChanged(app, val)
            app.SnapEnabled = logical(val);
        end

        %% ==================== 字体样式控制 ====================
        function onGlobalFontChanged(app)
            sz = app.GlobalFontField.Value;
            app.TickFontField.Value     = sz;
            app.LabelFontField.Value    = sz;
            app.TitleFontField.Value    = sz;
            app.LegendFontField.Value   = max(6, round(sz - 1.0, 1));
            app.TextFontField.Value     = sz;
            app.SubLabelSizeField.Value = sz;
        end

        function onSyncAndRefreshThumbs(app)
            app.onGlobalFontChanged();
            % 刷新前自动紧凑靠拢并放大铺满版心宽，避免间距过宽
            app.onCompactFillLayout(false);
            app.refreshAllThumbnails();
        end

        function cfg = getStyleConfig(app)
            cfg = struct( ...
                'fontName',            app.FontNameDropDown.Value, ...
                'tickFontSize',        app.TickFontField.Value, ...
                'labelFontSize',       app.LabelFontField.Value, ...
                'titleFontSize',       app.TitleFontField.Value, ...
                'legendFontSize',      app.LegendFontField.Value, ...
                'textFontSize',        app.TextFontField.Value, ...
                'forceNormalWeight',   app.ForceNormalCheck.Value, ...
                'enableAxisLineWidth', app.AxisLWCheck.Value, ...
                'axisLineWidth',       app.AxisLWField.Value, ...
                'enablePlotLineWidth', app.PlotLWCheck.Value, ...
                'plotLineWidth',       app.PlotLWField.Value, ...
                'tightLayout',         app.TightLayoutCheck.Value ...
            );
        end

        function lcfg = getLabelConfig(app)
            lcfg = struct( ...
                'enabled',  app.SubLabelCheck.Value, ...
                'style',    app.SubLabelStyleDD.Value, ...
                'position', app.SubLabelPosDD.Value, ...
                'fontSize', app.SubLabelSizeField.Value, ...
                'bold',     app.SubLabelBoldCheck.Value ...
            );
        end

        %% ==================== 载入 .fig 与缩略图 ====================
        function onLoadSingleFig(app)
            idxs = app.getSelectedIndices();
            if isempty(idxs)
                uialert(app.UIFigure, '请先在画布中点击选中一个图框。', '提示');
                return;
            end
            app.loadFigIntoSlot(idxs(1));
        end

        function loadFigIntoSlot(app, slotIdx)
            [file, path] = uigetfile( ...
                {'*.fig', 'MATLAB 图形文件 (*.fig)'; ...
                 '*.png;*.jpg;*.jpeg;*.tif;*.tiff', '普通图片 (*.png, *.jpg, *.tif)'}, ...
                sprintf('为图框 #%d 选择 .fig 或图片', slotIdx));
            if isequal(file, 0), return; end

            app.Slots(slotIdx).filePath = fullfile(path, file);
            app.setStatus(sprintf('正在渲染预览：%s ...', file));
            drawnow;
            app.updateSlotThumbnail(slotIdx);
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus(sprintf('已载入：%s（字体已按 %.1f pt 预览）', ...
                file, app.GlobalFontField.Value));
        end

        function onBatchLoadFigs(app)
            if isempty(app.Slots)
                uialert(app.UIFigure, '请先添加图框或套用一个排版模板。', '提示');
                return;
            end
            [files, path] = uigetfile( ...
                {'*.fig', 'MATLAB 图形文件 (*.fig)'; ...
                 '*.png;*.jpg;*.jpeg;*.tif;*.tiff', '普通图片'}, ...
                '选择多个 .fig 文件（按顺序填入图框）', 'MultiSelect', 'on');
            if isequal(files, 0), return; end
            if ischar(files), files = {files}; end

            nFill = min(numel(files), numel(app.Slots));
            dlg = uiprogressdlg(app.UIFigure, 'Title', '批量载入 .fig', ...
                'Message', '正在渲染缩略图...', 'Value', 0);
            cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>

            for i = 1:nFill
                dlg.Value = (i - 0.5) / nFill;
                dlg.Message = sprintf('正在载入 (%d/%d): %s', i, nFill, files{i});
                app.Slots(i).filePath = fullfile(path, files{i});
                app.updateSlotThumbnail(i);
            end
            dlg.Value = 1;
            app.redrawCanvas();
            app.syncInspector();
            app.setStatus(sprintf('已批量载入 %d 个图形文件。', nFill));
        end

        function updateSlotThumbnail(app, slotIdx)
            if slotIdx < 1 || slotIdx > numel(app.Slots), return; end
            fp = app.Slots(slotIdx).filePath;
            if isempty(fp) || ~isfile(fp)
                app.Slots(slotIdx).thumbImg = [];
                return;
            end
            try
                styleCfg = app.getStyleConfig();
                % 缩略图使用 110 DPI，兼顾速度与真实字号比例预览
                app.Slots(slotIdx).thumbImg = renderSingleSlot( ...
                    fp, app.Slots(slotIdx).w, app.Slots(slotIdx).h, 110, styleCfg);
            catch ME
                app.Slots(slotIdx).thumbImg = [];
                app.setStatus(sprintf('缩略图渲染失败: %s', ME.message));
            end
        end

        function refreshAllThumbnails(app)
            hasFiles = find(arrayfun(@(s) ~isempty(s.filePath) && isfile(s.filePath), app.Slots));
            if isempty(hasFiles)
                app.redrawCanvas();
                return;
            end
            dlg = uiprogressdlg(app.UIFigure, 'Title', '刷新预览', ...
                'Message', '正在按最新字体和尺寸重新渲染预览...', 'Value', 0);
            cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>

            for k = 1:numel(hasFiles)
                idx = hasFiles(k);
                dlg.Value = (k - 0.3) / numel(hasFiles);
                app.updateSlotThumbnail(idx);
            end
            dlg.Value = 1;
            app.redrawCanvas();
            app.setStatus('已按最新字号与物理尺寸刷新所有图框预览。');
        end

        function onRefreshSelectedThumb(app)
            idxs = app.getSelectedIndices();
            for k = 1:numel(idxs)
                app.updateSlotThumbnail(idxs(k));
            end
            app.redrawCanvas();
        end

        function onClearSelectedFig(app)
            idxs = app.getSelectedIndices();
            for k = 1:numel(idxs)
                app.Slots(idxs(k)).filePath = '';
                app.Slots(idxs(k)).thumbImg = [];
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function onStretchFullWidth(app)
            idxs = app.getSelectedIndices();
            if isempty(idxs), return; end
            for k = 1:numel(idxs)
                app.Slots(idxs(k)).x = 0;
                app.Slots(idxs(k)).w = app.CanvasWidthCm;
                if ~isempty(app.Slots(idxs(k)).filePath)
                    app.updateSlotThumbnail(idxs(k));
                end
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        %% ==================== 属性检查器联动 ====================
        function syncInspector(app)
            % 更新列表
            n = numel(app.Slots);
            items = cell(1, n);
            itemsData = zeros(1, n);
            autoLbls = app.computePreviewAutoLabels();
            for i = 1:n
                s = app.Slots(i);
                lbl = strtrim(s.label);
                if isempty(lbl), lbl = autoLbls{i}; end
                if ~isempty(s.filePath)
                    [~, fn, ext] = fileparts(s.filePath);
                    fDesc = [fn ext];
                else
                    fDesc = '[空]';
                end
                items{i} = sprintf('%s  %.2f×%.2fcm  %s', lbl, s.w, s.h, fDesc);
                itemsData(i) = s.id;
            end
            app.SlotListBox.Items = items;
            app.SlotListBox.ItemsData = itemsData;
            validSel = intersect(app.SelectedIds, itemsData, 'stable');
            app.SelectedIds = validSel;
            if ~isempty(validSel)
                app.SlotListBox.Value = validSel;
            else
                app.SlotListBox.Value = {};
            end

            % 更新选中项参数框
            idxs = app.getSelectedIndices();
            if isempty(idxs)
                app.InspectorInfoLabel.Text = '未选中图框';
                app.setInspectorEnabled('off');
                app.SlotFileLabel.Text = '点击画布中的图框查看或修改属性';
            elseif numel(idxs) > 1
                app.InspectorInfoLabel.Text = sprintf('已多选 %d 个图框（可使用对齐工具）', numel(idxs));
                app.setInspectorEnabled('off');
                app.SlotFileLabel.Text = '多选状态下可使用左侧对齐/等宽/等高工具';
            else
                s = app.Slots(idxs(1));
                app.InspectorInfoLabel.Text = sprintf('图框 #%d（单位：cm）', idxs(1));
                app.setInspectorEnabled('on');
                app.SlotXField.Value = s.x;
                app.SlotYField.Value = s.y;
                app.SlotWField.Value = s.w;
                app.SlotHField.Value = s.h;
                app.SlotLabelEdit.Value = s.label;
                if isempty(s.filePath)
                    app.SlotFileLabel.Text = '尚未载入 .fig 文件（双击图框或点下方按钮）';
                else
                    [~, fn, ext] = fileparts(s.filePath);
                    app.SlotFileLabel.Text = sprintf('已绑定: %s%s', fn, ext);
                end
            end
        end

        function setInspectorEnabled(app, state)
            app.SlotXField.Enable = state;
            app.SlotYField.Enable = state;
            app.SlotWField.Enable = state;
            app.SlotHField.Enable = state;
            app.SlotLabelEdit.Enable = state;
            app.LoadFigBtn.Enable = state;
            app.RefreshThumbBtn.Enable = state;
            app.ClearFigBtn.Enable = state;
            app.FullWidthBtn.Enable = state;
        end

        function onListBoxSelectionChanged(app)
            val = app.SlotListBox.Value;
            if isempty(val)
                app.SelectedIds = [];
            else
                app.SelectedIds = double(val(:)');
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function onInspectorGeomChanged(app)
            idxs = app.getSelectedIndices();
            if numel(idxs) ~= 1, return; end
            i = idxs(1);
            app.Slots(i).x = max(0, app.SlotXField.Value);
            app.Slots(i).y = max(0, app.SlotYField.Value);
            app.Slots(i).w = max(1.0, app.SlotWField.Value);
            app.Slots(i).h = max(1.0, app.SlotHField.Value);
            if ~isempty(app.Slots(i).filePath)
                app.updateSlotThumbnail(i);
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function onSlotLabelChanged(app)
            idxs = app.getSelectedIndices();
            if numel(idxs) ~= 1, return; end
            app.Slots(idxs(1)).label = app.SlotLabelEdit.Value;
            app.redrawCanvas();
            app.syncInspector();
        end

        %% ==================== 画布绘制与鼠标交互 ====================
        function redrawCanvas(app)
            ax = app.CanvasAxes;
            cla(ax);
            hold(ax, 'on');

            W = app.CanvasWidthCm;
            H = app.CanvasHeightCm;
            pad = max(0.8, W * 0.06);
            ax.XLim = [-pad, W + pad];
            ax.YLim = [-pad, H + pad];
            ax.XTick = 0:2:ceil(W);
            ax.YTick = 0:2:ceil(H);

            % 柔和多层纸张投影 + 纯白画布纸面（Apple Keynote / Freeform 风格）
            patch(ax, [0.22, W+0.22, W+0.22, 0.22], [0.24, 0.24, H+0.24, H+0.24], ...
                [0.65, 0.70, 0.78], 'FaceAlpha', 0.22, 'EdgeColor', 'none', 'PickableParts', 'none');
            patch(ax, [0.10, W+0.10, W+0.10, 0.10], [0.12, 0.12, H+0.12, H+0.12], ...
                [0.65, 0.70, 0.78], 'FaceAlpha', 0.18, 'EdgeColor', 'none', 'PickableParts', 'none');
            rectangle(ax, 'Position', [0, 0, W, H], ...
                'FaceColor', [1, 1, 1], 'EdgeColor', [0.78, 0.82, 0.88], 'LineWidth', 1.1);

            % 1 cm 极细辅助网格
            for gx = 1:floor(W - 0.01)
                plot(ax, [gx, gx], [0, H], '-', 'Color', [0.945, 0.955, 0.975], 'LineWidth', 0.5);
            end
            for gy = 1:floor(H - 0.01)
                plot(ax, [0, W], [gy, gy], '-', 'Color', [0.945, 0.955, 0.975], 'LineWidth', 0.5);
            end

            % 绘制每个图框
            autoLbls = app.computePreviewAutoLabels();
            for i = 1:numel(app.Slots)
                s = app.Slots(i);
                isSel = ismember(s.id, app.SelectedIds);

                if ~isempty(s.thumbImg)
                    image(ax, 'XData', [s.x, s.x + s.w], 'YData', [s.y, s.y + s.h], ...
                        'CData', s.thumbImg);
                    edgeCol = [0.78, 0.82, 0.88];
                    lw = 0.8;
                    if isSel
                        edgeCol = [0.00, 0.48, 1.00];
                        lw = 2.0;
                    end
                    rectangle(ax, 'Position', [s.x, s.y, s.w, s.h], ...
                        'FaceColor', 'none', 'EdgeColor', edgeCol, 'LineWidth', lw);
                else
                    % 未放 .fig 时：Apple 半透明磨砂蓝玻璃卡片
                    px = [s.x, s.x + s.w, s.x + s.w, s.x];
                    py = [s.y, s.y, s.y + s.h, s.y + s.h];
                    if isSel
                        patch(ax, px, py, [0.00, 0.48, 1.00], 'FaceAlpha', 0.22, ...
                            'EdgeColor', [0.00, 0.48, 1.00], 'LineWidth', 2.0, 'PickableParts', 'none');
                    else
                        patch(ax, px, py, [0.16, 0.52, 0.96], 'FaceAlpha', 0.13, ...
                            'EdgeColor', [0.35, 0.62, 0.95], 'LineWidth', 1.1, 'PickableParts', 'none');
                    end

                    % 居中文字：清爽深蓝灰文字
                    lblStr = strtrim(s.label);
                    if isempty(lblStr), lblStr = autoLbls{i}; end
                    infoTxt = sprintf('%s\n%.2f × %.2f cm\n双击放入 .fig', lblStr, s.w, s.h);
                    text(ax, s.x + s.w/2, s.y + s.h/2, infoTxt, ...
                        'Color', [0.08, 0.32, 0.68], 'FontWeight', 'bold', 'FontSize', 10, ...
                        'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', ...
                        'Interpreter', 'none', 'PickableParts', 'none');
                end

                % 若已有缩略图且启用子图编号，在角上显示清爽编号角标
                if ~isempty(s.thumbImg) && app.SubLabelCheck.Value
                    lblStr = strtrim(s.label);
                    if isempty(lblStr), lblStr = autoLbls{i}; end
                    text(ax, s.x + 0.14, s.y + 0.14, lblStr, ...
                        'Color', [0.12, 0.16, 0.24], 'BackgroundColor', [0.98, 0.99, 1.0], ...
                        'Margin', 1.5, 'FontWeight', 'bold', 'FontSize', 9, ...
                        'HorizontalAlignment', 'left', 'VerticalAlignment', 'top', ...
                        'Interpreter', 'none', 'PickableParts', 'none');
                end

                % 选中状态：绘制 Keynote 风格白色圆点 + Apple 蓝边框 8 向手柄
                if isSel
                    pts = app.getHandlePoints(s);
                    plot(ax, pts(:,1), pts(:,2), 'o', ...
                        'MarkerSize', 7.5, ...
                        'MarkerFaceColor', [1, 1, 1], ...
                        'MarkerEdgeColor', [0.00, 0.48, 1.00], ...
                        'LineWidth', 1.6, 'PickableParts', 'none');
                end
            end

            hold(ax, 'off');
        end

        function onMouseDown(app)
            if ~app.isPointerInsideCanvas(), return; end
            pt = app.CanvasAxes.CurrentPoint(1, 1:2);
            selType = app.UIFigure.SelectionType; % 'normal' | 'extend'(Shift) | 'open'(双击)

            % 1. 先检查是否点中已选中图框的 8 个缩放手柄
            selIdxs = app.getSelectedIndices();
            hitTol = max(0.30, app.CanvasWidthCm * 0.022);
            if numel(selIdxs) == 1 && ~strcmp(selType, 'open')
                s = app.Slots(selIdxs(1));
                [hName, hit] = app.hitTestHandles(pt, s, hitTol);
                if hit
                    app.DragMode = 'resize';
                    app.DragHandle = hName;
                    app.DragSlotIdx = selIdxs(1);
                    app.DragStartPt = pt;
                    app.DragOrigRect = [s.x, s.y, s.w, s.h];
                    return;
                end
            end

            % 2. 检查是否点中某个图框内部（从顶层向下遍历）
            hitIdx = 0;
            for i = numel(app.Slots):-1:1
                s = app.Slots(i);
                if pt(1) >= s.x && pt(1) <= s.x + s.w && pt(2) >= s.y && pt(2) <= s.y + s.h
                    hitIdx = i;
                    break;
                end
            end

            if hitIdx == 0
                % 点击空白处取消选中
                app.SelectedIds = [];
                app.DragMode = 'none';
                app.redrawCanvas();
                app.syncInspector();
                return;
            end

            hitId = app.Slots(hitIdx).id;

            % 双击：直接载入 .fig
            if strcmp(selType, 'open')
                app.SelectedIds = hitId;
                app.DragMode = 'none';
                app.redrawCanvas();
                app.syncInspector();
                app.loadFigIntoSlot(hitIdx);
                return;
            end

            % Shift+点击：切换多选
            if strcmp(selType, 'extend')
                if ismember(hitId, app.SelectedIds)
                    app.SelectedIds(app.SelectedIds == hitId) = [];
                else
                    app.SelectedIds(end+1) = hitId;
                end
                app.DragMode = 'none';
                app.redrawCanvas();
                app.syncInspector();
                return;
            end

            % 普通单击并准备拖动
            if ~ismember(hitId, app.SelectedIds)
                app.SelectedIds = hitId;
            end
            app.DragMode = 'move';
            app.DragSlotIdx = hitIdx;
            app.DragStartPt = pt;
            app.DragOrigRect = [app.Slots(hitIdx).x, app.Slots(hitIdx).y, ...
                                app.Slots(hitIdx).w, app.Slots(hitIdx).h];

            curSelIdxs = app.getSelectedIndices();
            app.DragOrigMulti = zeros(numel(curSelIdxs), 2);
            for k = 1:numel(curSelIdxs)
                app.DragOrigMulti(k, :) = [app.Slots(curSelIdxs(k)).x, app.Slots(curSelIdxs(k)).y];
            end

            app.redrawCanvas();
            app.syncInspector();
        end

        function onMouseMove(app)
            if strcmp(app.DragMode, 'none'), return; end
            pt = app.CanvasAxes.CurrentPoint(1, 1:2);
            dx = pt(1) - app.DragStartPt(1);
            dy = pt(2) - app.DragStartPt(2);

            app.clearGuideLines();
            guidesX = [];
            guidesY = [];

            if strcmp(app.DragMode, 'move')
                selIdxs = app.getSelectedIndices();
                if numel(selIdxs) == 1
                    idx = selIdxs(1);
                    newX = app.DragOrigRect(1) + dx;
                    newY = app.DragOrigRect(2) + dy;
                    w = app.DragOrigRect(3);
                    h = app.DragOrigRect(4);

                    if app.SnapEnabled
                        [newX, gx] = app.snapValueX(newX, w, idx);
                        [newY, gy] = app.snapValueY(newY, h, idx);
                        guidesX = gx; guidesY = gy;
                    end
                    app.Slots(idx).x = round(max(0, min(app.CanvasWidthCm - w, newX)), 2);
                    app.Slots(idx).y = round(max(0, min(app.CanvasHeightCm - h, newY)), 2);
                else
                    for k = 1:numel(selIdxs)
                        idx = selIdxs(k);
                        nx = app.DragOrigMulti(k, 1) + dx;
                        ny = app.DragOrigMulti(k, 2) + dy;
                        if app.SnapEnabled
                            nx = round(nx / app.SnapStepCm) * app.SnapStepCm;
                            ny = round(ny / app.SnapStepCm) * app.SnapStepCm;
                        end
                        app.Slots(idx).x = round(max(0, min(app.CanvasWidthCm - app.Slots(idx).w, nx)), 2);
                        app.Slots(idx).y = round(max(0, min(app.CanvasHeightCm - app.Slots(idx).h, ny)), 2);
                    end
                end
            elseif strcmp(app.DragMode, 'resize')
                idx = app.DragSlotIdx;
                x1 = app.DragOrigRect(1);
                y1 = app.DragOrigRect(2);
                x2 = x1 + app.DragOrigRect(3);
                y2 = y1 + app.DragOrigRect(4);
                minSz = 1.5;

                hName = app.DragHandle;
                if contains(hName, 'w')
                    x1 = min(x2 - minSz, max(0, x1 + dx));
                    if app.SnapEnabled, x1 = round(x1 / app.SnapStepCm) * app.SnapStepCm; end
                end
                if contains(hName, 'e')
                    x2 = max(x1 + minSz, min(app.CanvasWidthCm, x2 + dx));
                    if app.SnapEnabled, x2 = round(x2 / app.SnapStepCm) * app.SnapStepCm; end
                end
                if contains(hName, 'n')
                    y1 = min(y2 - minSz, max(0, y1 + dy));
                    if app.SnapEnabled, y1 = round(y1 / app.SnapStepCm) * app.SnapStepCm; end
                end
                if contains(hName, 's')
                    y2 = max(y1 + minSz, min(app.CanvasHeightCm, y2 + dy));
                    if app.SnapEnabled, y2 = round(y2 / app.SnapStepCm) * app.SnapStepCm; end
                end

                app.Slots(idx).x = round(x1, 2);
                app.Slots(idx).y = round(y1, 2);
                app.Slots(idx).w = round(max(minSz, x2 - x1), 2);
                app.Slots(idx).h = round(max(minSz, y2 - y1), 2);
            end

            app.redrawCanvas();
            app.drawGuideLines(guidesX, guidesY);
        end

        function onMouseUp(app)
            if strcmp(app.DragMode, 'none'), return; end
            wasResize = strcmp(app.DragMode, 'resize');
            idx = app.DragSlotIdx;

            app.DragMode = 'none';
            app.DragHandle = '';
            app.clearGuideLines();

            % 缩放结束后，如果该图框已绑定 .fig，自动按新尺寸刷新缩略图（实时看新尺寸下的字号比例）
            if wasResize && idx >= 1 && idx <= numel(app.Slots) && ~isempty(app.Slots(idx).filePath)
                app.setStatus(sprintf('尺寸已调整为 %.2f × %.2f cm，正在按 1:1 字号刷新预览...', ...
                    app.Slots(idx).w, app.Slots(idx).h));
                drawnow;
                app.updateSlotThumbnail(idx);
            end
            app.redrawCanvas();
            app.syncInspector();
        end

        function onKeyPress(app, evt)
            idxs = app.getSelectedIndices();
            if isempty(idxs), return; end
            step = 0.1;
            moved = false;
            switch evt.Key
                case {'delete', 'backspace'}
                    app.onDeleteSelected();
                    return;
                case 'leftarrow'
                    for k = idxs, app.Slots(k).x = round(max(0, app.Slots(k).x - step), 2); end
                    moved = true;
                case 'rightarrow'
                    for k = idxs, app.Slots(k).x = round(min(app.CanvasWidthCm - app.Slots(k).w, app.Slots(k).x + step), 2); end
                    moved = true;
                case 'uparrow'
                    for k = idxs, app.Slots(k).y = round(max(0, app.Slots(k).y - step), 2); end
                    moved = true;
                case 'downarrow'
                    for k = idxs, app.Slots(k).y = round(min(app.CanvasHeightCm - app.Slots(k).h, app.Slots(k).y + step), 2); end
                    moved = true;
            end
            if moved
                app.redrawCanvas();
                app.syncInspector();
            end
        end

        %% ==================== 预览、导出与方案存取 ====================
        function onPreviewComposite(app)
            if isempty(app.Slots)
                uialert(app.UIFigure, '画布中还没有图框。', '提示');
                return;
            end
            dlg = uiprogressdlg(app.UIFigure, 'Title', '正在生成 1:1 渲染预览', ...
                'Message', '正在调用 MATLAB 统一字号并合成预览...', 'Value', 0.1);
            cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>

            try
                opts = struct('autoCropHeight', app.AutoCropCheck.Value, ...
                    'saveIndividuals', false, ...
                    'progressFcn', @(i,n,msg) app.updateProgress(dlg, i, n, msg));

                [imgRgb, fW, fH] = renderComposite(app.Slots, ...
                    app.CanvasWidthCm, app.CanvasHeightCm, 200, ...
                    app.getStyleConfig(), app.getLabelConfig(), opts);

                % 使用独立 uifigure + uiimage 显示预览，避免停靠到 MATLAB Desktop 唤醒主窗口
                aspect = fH / max(0.1, fW);
                winW = 920;
                winH = min(820, max(360, round(winW * aspect) + 70));
                prevFig = uifigure( ...
                    'Name', sprintf('1:1 排版预览 (最终印刷尺寸 %.2f × %.2f cm)', fW, fH), ...
                    'Position', [180, 120, winW, winH], ...
                    'WindowStyle', 'normal', ...
                    'Color', [0.94, 0.95, 0.96]);
                pg = uigridlayout(prevFig, [2, 1]);
                pg.RowHeight = {28, '1x'};
                pg.Padding = [12, 10, 12, 12];
                uilabel(pg, 'Text', sprintf('最终印刷尺寸: %.2f × %.2f cm  |  全局字号 %.1f pt (%s)', ...
                    fW, fH, app.GlobalFontField.Value, app.FontNameDropDown.Value), ...
                    'FontWeight', 'bold', 'HorizontalAlignment', 'center');
                uiimage(pg, 'ImageSource', imgRgb, 'ScaleMethod', 'fit');
                app.setStatus(sprintf('预览已生成：最终尺寸 %.2f × %.2f cm', fW, fH));
            catch ME
                uialert(app.UIFigure, ME.message, '预览生成失败');
            end
        end

        function onOpenEditableFigure(app)
            if isempty(app.Slots)
                uialert(app.UIFigure, '画布中还没有图框。', '提示');
                return;
            end
            dlg = uiprogressdlg(app.UIFigure, 'Title', '正在构建可编辑 MATLAB 图形', ...
                'Message', '正在按 1:1 物理尺寸与统一字体组装原生可编辑 Figure ...', 'Value', 0.3);
            cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>

            try
                opts = struct('autoCropHeight', app.AutoCropCheck.Value, 'visible', true);
                [masterFig, fW, fH] = buildEditableFigure(app.Slots, ...
                    app.CanvasWidthCm, app.CanvasHeightCm, ...
                    app.getStyleConfig(), app.getLabelConfig(), opts);
                dlg.Value = 1.0;

                % 开启绘图编辑模式，方便直接点击图例/曲线/坐标轴，并可从顶部「科研微调与导出」打开属性检查器
                try, plotedit(masterFig, 'on'); catch, end
                app.setStatus(sprintf('已打开原生可编辑图 (%.2f × %.2f cm)：可在顶部「科研微调与导出」打开属性检查器或直接导出给 Word。', fW, fH));
            catch ME
                uialert(app.UIFigure, ME.message, '打开可编辑图形失败');
            end
        end

        function onExportComposite(app)
            if isempty(app.Slots)
                uialert(app.UIFigure, '画布中还没有图框。', '提示');
                return;
            end
            ext = app.FormatDropDown.Value;
            dpi = app.DpiDropDown.Value;

            % 直接导出原生可编辑 .fig 分支
            if strcmpi(ext, '.fig')
                defaultName = sprintf('Composite_Figure_%.1fcm.fig', app.CanvasWidthCm);
                [file, path] = uiputfile({'*.fig', 'MATLAB 可编辑图形 (*.fig)'}, ...
                    '保存可编辑 .fig 文件', defaultName);
                if isequal(file, 0), return; end
                outPath = fullfile(path, file);

                dlg = uiprogressdlg(app.UIFigure, 'Title', '正在生成可编辑 .fig', ...
                    'Message', '正在按 1:1 物理尺寸组装原生 MATLAB Figure ...', 'Value', 0.3);
                cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>
                try
                    opts = struct('autoCropHeight', app.AutoCropCheck.Value, 'visible', false);
                    [masterFig, fW, fH] = buildEditableFigure(app.Slots, ...
                        app.CanvasWidthCm, app.CanvasHeightCm, ...
                        app.getStyleConfig(), app.getLabelConfig(), opts);
                    cleanupFig = onCleanup(@() close(masterFig)); %#ok<NASGU>
                    savefig(masterFig, outPath);
                    dlg.Value = 1.0;
                    uialert(app.UIFigure, sprintf('已保存可编辑 .fig 文件：\n%s\n物理尺寸：%.2f × %.2f cm', ...
                        outPath, fW, fH), '导出成功', 'Icon', 'success');
                    app.setStatus(sprintf('已导出可编辑 .fig：%s (%.2f×%.2f cm)', file, fW, fH));
                catch ME
                    uialert(app.UIFigure, ME.message, '导出 .fig 失败');
                end
                return;
            end

            defaultName = sprintf('Composite_Figure_%.1fcm_%ddpi%s', app.CanvasWidthCm, dpi, ext);
            [file, path] = uiputfile({['*' ext], sprintf('高清图像 (*%s)', ext)}, ...
                '保存最终高清科研大图', defaultName);
            if isequal(file, 0), return; end
            outPath = fullfile(path, file);

            dlg = uiprogressdlg(app.UIFigure, 'Title', '正在导出高清科研大图', ...
                'Message', '正在以 1:1 物理尺寸重新渲染每个 .fig ...', 'Value', 0.05);
            cleanupDlg = onCleanup(@() close(dlg)); %#ok<NASGU>

            try
                [outDir, outBase, ~] = fileparts(outPath);
                indDir = '';
                if app.SaveSubCheck.Value
                    indDir = fullfile(outDir, [outBase '_subplots']);
                    if ~isfolder(indDir), mkdir(indDir); end
                end

                opts = struct( ...
                    'autoCropHeight',  app.AutoCropCheck.Value, ...
                    'saveIndividuals', app.SaveSubCheck.Value, ...
                    'individualDir',   indDir, ...
                    'individualExt',   ext, ...
                    'progressFcn',     @(i,n,msg) app.updateProgress(dlg, i, n, msg));

                [imgRgb, fW, fH] = renderComposite(app.Slots, ...
                    app.CanvasWidthCm, app.CanvasHeightCm, dpi, ...
                    app.getStyleConfig(), app.getLabelConfig(), opts);

                dlg.Message = '正在写入精确 DPI 元数据并保存文件...';
                dlg.Value = 0.95;
                saveImageWithDpi(imgRgb, outPath, dpi);
                dlg.Value = 1.0;

                msg = sprintf(['导出成功！\n\n文件：%s\n物理尺寸：%.2f × %.2f cm\n分辨率：%d DPI\n\n' ...
                    '提示：直接将此图片拖入 Word，Word 会自动按 100%% 原始宽度 (%.2f cm) 显示，' ...
                    '图内字体将与 Word 正文字号完全一致！'], ...
                    outPath, fW, fH, dpi, fW);
                uialert(app.UIFigure, msg, '导出成功', 'Icon', 'success');
                app.setStatus(sprintf('已导出：%s (%.2f×%.2f cm, %d DPI)', file, fW, fH, dpi));
            catch ME
                uialert(app.UIFigure, ME.message, '导出失败');
            end
        end

        function onSaveProject(app)
            [file, path] = uiputfile({'*.json', '排版方案文件 (*.json)'}, ...
                '保存排版方案', 'layout_project.json');
            if isequal(file, 0), return; end

            proj = struct();
            proj.canvasWidthCm  = app.CanvasWidthCm;
            proj.canvasHeightCm = app.CanvasHeightCm;
            proj.gapXCm         = app.GapXCm;
            proj.gapYCm         = app.GapYCm;
            proj.styleCfg       = app.getStyleConfig();
            proj.labelCfg       = app.getLabelConfig();
            proj.slots          = arrayfun(@(s) rmfield(s, 'thumbImg'), app.Slots);

            txt = jsonencode(proj, 'PrettyPrint', true);
            fid = fopen(fullfile(path, file), 'w', 'n', 'UTF-8');
            fwrite(fid, txt, 'char');
            fclose(fid);
            app.setStatus(sprintf('排版方案已保存：%s', file));
        end

        function onLoadProject(app)
            [file, path] = uigetfile({'*.json', '排版方案文件 (*.json)'}, '加载排版方案');
            if isequal(file, 0), return; end

            try
                txt = fileread(fullfile(path, file));
                proj = jsondecode(txt);

                app.CanvasWidthCm  = proj.canvasWidthCm;
                app.CanvasHeightCm = proj.canvasHeightCm;
                app.CanvasWField.Value = app.CanvasWidthCm;
                app.CanvasHField.Value = app.CanvasHeightCm;
                if isfield(proj, 'gapXCm')
                    app.GapXCm = proj.gapXCm;
                    app.GapXField.Value = app.GapXCm;
                elseif isfield(proj, 'gapCm')
                    app.GapXCm = proj.gapCm;
                    app.GapXField.Value = app.GapXCm;
                end
                if isfield(proj, 'gapYCm')
                    app.GapYCm = proj.gapYCm;
                    app.GapYField.Value = app.GapYCm;
                elseif isfield(proj, 'gapCm')
                    app.GapYCm = proj.gapCm;
                    app.GapYField.Value = app.GapYCm;
                end

                % 恢复字体配置
                if isfield(proj, 'styleCfg')
                    sc = proj.styleCfg;
                    if ~any(strcmp(app.FontNameDropDown.Items, sc.fontName))
                        app.FontNameDropDown.Items = [{sc.fontName}, app.FontNameDropDown.Items];
                    end
                    app.FontNameDropDown.Value = sc.fontName;
                    app.GlobalFontField.Value  = sc.labelFontSize;
                    app.TickFontField.Value    = sc.tickFontSize;
                    app.LabelFontField.Value   = sc.labelFontSize;
                    app.TitleFontField.Value   = sc.titleFontSize;
                    app.LegendFontField.Value  = sc.legendFontSize;
                    app.TextFontField.Value    = sc.textFontSize;
                    app.TightLayoutCheck.Value = logical(sc.tightLayout);
                end

                % 恢复图框
                loaded = proj.slots;
                app.Slots = struct([]);
                for i = 1:numel(loaded)
                    s = loaded(i);
                    s.thumbImg = [];
                    if isempty(app.Slots)
                        app.Slots = s;
                    else
                        app.Slots(end+1) = s;
                    end
                end
                app.NextSlotId = numel(app.Slots) + 1;
                app.SelectedIds = [];
                app.refreshAllThumbnails();
                app.syncInspector();
                app.setStatus(sprintf('已加载排版方案：%s', file));
            catch ME
                uialert(app.UIFigure, ME.message, '加载方案失败');
            end
        end

        %% ==================== 辅助工具函数 ====================
        function idxs = getSelectedIndices(app)
            if isempty(app.Slots) || isempty(app.SelectedIds)
                idxs = [];
                return;
            end
            allIds = arrayfun(@(s) s.id, app.Slots);
            idxs = find(ismember(allIds, app.SelectedIds));
        end

        function tf = isPointerInsideCanvas(app)
            pt = app.CanvasAxes.CurrentPoint(1, 1:2);
            xl = app.CanvasAxes.XLim;
            yl = app.CanvasAxes.YLim;
            tf = pt(1) >= xl(1) && pt(1) <= xl(2) && pt(2) >= yl(1) && pt(2) <= yl(2);
        end

        function pts = getHandlePoints(~, s)
            x1 = s.x; xm = s.x + s.w/2; x2 = s.x + s.w;
            y1 = s.y; ym = s.y + s.h/2; y2 = s.y + s.h;
            pts = [x1, y1; xm, y1; x2, y1; x2, ym; x2, y2; xm, y2; x1, y2; x1, ym];
        end

        function [hName, hit] = hitTestHandles(app, pt, s, tol)
            names = {'nw','n','ne','e','se','s','sw','w'};
            pts = app.getHandlePoints(s);
            dists = hypot(pts(:,1) - pt(1), pts(:,2) - pt(2));
            [minD, idx] = min(dists);
            if minD <= tol
                hName = names{idx};
                hit = true;
            else
                hName = '';
                hit = false;
            end
        end

        function [snappedX, guideX] = snapValueX(app, x, w, selfIdx)
            tol = 0.25;
            guideX = [];
            targets = [0, app.CanvasWidthCm];
            for i = 1:numel(app.Slots)
                if i == selfIdx, continue; end
                targets = [targets, app.Slots(i).x, app.Slots(i).x + app.Slots(i).w]; %#ok<AGROW>
            end
            % 检查左边缘吸附
            [dL, iL] = min(abs(targets - x));
            % 检查右边缘吸附
            [dR, iR] = min(abs(targets - (x + w)));
            if dL <= tol && dL <= dR
                snappedX = targets(iL);
                guideX = targets(iL);
            elseif dR <= tol
                snappedX = targets(iR) - w;
                guideX = targets(iR);
            else
                snappedX = round(x / app.SnapStepCm) * app.SnapStepCm;
            end
        end

        function [snappedY, guideY] = snapValueY(app, y, h, selfIdx)
            tol = 0.25;
            guideY = [];
            targets = [0, app.CanvasHeightCm];
            for i = 1:numel(app.Slots)
                if i == selfIdx, continue; end
                targets = [targets, app.Slots(i).y, app.Slots(i).y + app.Slots(i).h]; %#ok<AGROW>
            end
            [dT, iT] = min(abs(targets - y));
            [dB, iB] = min(abs(targets - (y + h)));
            if dT <= tol && dT <= dB
                snappedY = targets(iT);
                guideY = targets(iT);
            elseif dB <= tol
                snappedY = targets(iB) - h;
                guideY = targets(iB);
            else
                snappedY = round(y / app.SnapStepCm) * app.SnapStepCm;
            end
        end

        function drawGuideLines(app, gx, gy)
            ax = app.CanvasAxes;
            hold(ax, 'on');
            app.GuideLines = gobjects(0);
            for k = 1:numel(gx)
                app.GuideLines(end+1) = plot(ax, [gx(k), gx(k)], [0, app.CanvasHeightCm], ...
                    '--', 'Color', [0.95, 0.20, 0.20], 'LineWidth', 1.0);
            end
            for k = 1:numel(gy)
                app.GuideLines(end+1) = plot(ax, [0, app.CanvasWidthCm], [gy(k), gy(k)], ...
                    '--', 'Color', [0.95, 0.20, 0.20], 'LineWidth', 1.0);
            end
            hold(ax, 'off');
        end

        function clearGuideLines(app)
            if ~isempty(app.GuideLines)
                delete(app.GuideLines(isgraphics(app.GuideLines)));
                app.GuideLines = gobjects(0);
            end
        end

        function labels = computePreviewAutoLabels(app)
            n = numel(app.Slots);
            labels = cell(1, n);
            if n == 0, return; end
            ys = arrayfun(@(s) s.y, app.Slots);
            xs = arrayfun(@(s) s.x, app.Slots);
            [~, order] = sortrows([round(ys(:) / 0.6), xs(:)]);
            st = app.SubLabelStyleDD.Value;
            for rank = 1:n
                idx = order(rank);
                cL = char('a' + mod(rank - 1, 26));
                cU = char('A' + mod(rank - 1, 26));
                switch st
                    case '(a)', labels{idx} = sprintf('(%s)', cL);
                    case '(A)', labels{idx} = sprintf('(%s)', cU);
                    case 'a',   labels{idx} = cL;
                    case 'A',   labels{idx} = cU;
                    case 'a)',  labels{idx} = sprintf('%s)', cL);
                    otherwise,  labels{idx} = sprintf('(%s)', cL);
                end
            end
        end

        function updateProgress(~, dlg, i, n, msg)
            if isvalid(dlg)
                dlg.Value = min(0.92, (i - 0.3) / max(1, n));
                dlg.Message = msg;
            end
        end

        function setStatus(app, msg)
            if ~isempty(app.StatusLabel) && isvalid(app.StatusLabel)
                app.StatusLabel.Text = msg;
            end
        end
    end
end
