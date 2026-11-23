function styleFigure(fig, widthCm, heightCm, styleCfg)
%STYLEFIGURE 将 MATLAB Figure 调整为精确物理尺寸 (cm) 并统一内部全部文字字体/字号
%
%   styleFigure(fig, widthCm, heightCm, styleCfg)
%
%   针对真实科研 .fig 的常见痛点做了深度适配：
%     - 强制 WindowStyle='normal' 避免唤醒/停靠 MATLAB 主界面
%     - 重置手动固定的 XLabel/YLabel 位置 (PositionMode='auto')，杜绝标签与刻度重叠
%     - 自动修正 Location='none' 的自由拖放图例，防止图例冲出右/上边界
%     - 完整支持 yyaxis 双 Y 轴、超长标签防裁切、可选统一常规字重与线宽

    if nargin < 4 || isempty(styleCfg)
        styleCfg = defaultStyleConfig();
    end

    % 1. 锁定 Figure 的精确物理打印尺寸（cm），并禁止停靠到 MATLAB Desktop
    trySet(fig, 'WindowStyle', 'normal');
    trySet(fig, 'Visible', 'off');
    set(fig, 'Color', 'w', 'InvertHardcopy', 'off');
    set(fig, 'Units', 'centimeters');
    set(fig, 'Position', [2, 2, max(1.0, widthCm), max(1.0, heightCm)]);
    set(fig, 'PaperUnits', 'centimeters');
    set(fig, 'PaperSize', [max(1.0, widthCm), max(1.0, heightCm)]);
    set(fig, 'PaperPositionMode', 'manual');
    set(fig, 'PaperPosition', [0, 0, max(1.0, widthCm), max(1.0, heightCm)]);

    fontName   = styleCfg.fontName;
    tickSz     = styleCfg.tickFontSize;
    labelSz    = styleCfg.labelFontSize;
    titleSz    = styleCfg.titleFontSize;
    legendSz   = styleCfg.legendFontSize;
    textSz     = styleCfg.textFontSize;
    forceNormalWeight = ~isfield(styleCfg, 'forceNormalWeight') || styleCfg.forceNormalWeight;

    % 2. 统一所有坐标轴（Axes / PolarAxes / GeoAxes）
    allAxes = [ ...
        findall(fig, 'Type', 'axes'); ...
        findall(fig, 'Type', 'polaraxes'); ...
        findall(fig, 'Type', 'geoaxes') ...
    ];

    specialTextHandles = gobjects(0);

    for k = 1:numel(allAxes)
        ax = allAxes(k);
        trySet(ax, 'FontName', fontName);
        trySet(ax, 'FontSize', tickSz);
        if forceNormalWeight
            trySet(ax, 'FontWeight', 'normal');
        end
        trySet(ax, 'LabelFontSizeMultiplier', 1.0);
        trySet(ax, 'TitleFontSizeMultiplier', 1.0);

        if isfield(styleCfg, 'enableAxisLineWidth') && styleCfg.enableAxisLineWidth
            trySet(ax, 'LineWidth', styleCfg.axisLineWidth);
        end

        % X 轴标签：重置手动固定位置，防止缩小后与刻度重叠
        if isprop(ax, 'XLabel') && isgraphics(ax.XLabel)
            resetLabelPosition(ax.XLabel);
            trySet(ax.XLabel, 'FontName', fontName);
            trySet(ax.XLabel, 'FontSize', labelSz);
            if forceNormalWeight, trySet(ax.XLabel, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = ax.XLabel; %#ok<AGROW>
        end

        % Y 轴（完整兼容单 Y 轴和 yyaxis 双 Y 轴）
        if isprop(ax, 'YAxis')
            for yIdx = 1:numel(ax.YAxis)
                try
                    ax.YAxis(yIdx).FontName = fontName;
                    ax.YAxis(yIdx).FontSize = tickSz;
                    if forceNormalWeight
                        ax.YAxis(yIdx).FontWeight = 'normal';
                    end
                    if isgraphics(ax.YAxis(yIdx).Label)
                        resetLabelPosition(ax.YAxis(yIdx).Label);
                        ax.YAxis(yIdx).Label.FontName = fontName;
                        ax.YAxis(yIdx).Label.FontSize = labelSz;
                        if forceNormalWeight
                            ax.YAxis(yIdx).Label.FontWeight = 'normal';
                        end
                        specialTextHandles(end+1) = ax.YAxis(yIdx).Label; %#ok<AGROW>
                    end
                catch
                end
            end
        elseif isprop(ax, 'YLabel') && isgraphics(ax.YLabel)
            resetLabelPosition(ax.YLabel);
            trySet(ax.YLabel, 'FontName', fontName);
            trySet(ax.YLabel, 'FontSize', labelSz);
            if forceNormalWeight, trySet(ax.YLabel, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = ax.YLabel; %#ok<AGROW>
        end

        if isprop(ax, 'ZLabel') && isgraphics(ax.ZLabel)
            resetLabelPosition(ax.ZLabel);
            trySet(ax.ZLabel, 'FontName', fontName);
            trySet(ax.ZLabel, 'FontSize', labelSz);
            if forceNormalWeight, trySet(ax.ZLabel, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = ax.ZLabel; %#ok<AGROW>
        end

        if isprop(ax, 'Title') && isgraphics(ax.Title)
            resetLabelPosition(ax.Title);
            trySet(ax.Title, 'FontName', fontName);
            trySet(ax.Title, 'FontSize', titleSz);
            if forceNormalWeight, trySet(ax.Title, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = ax.Title; %#ok<AGROW>
        end
        if isprop(ax, 'Subtitle') && isgraphics(ax.Subtitle)
            trySet(ax.Subtitle, 'FontName', fontName);
            trySet(ax.Subtitle, 'FontSize', labelSz);
            if forceNormalWeight, trySet(ax.Subtitle, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = ax.Subtitle; %#ok<AGROW>
        end
    end

    % 3. 统一所有图例 (Legend)，并修正自由定位导致冲出边界的问题
    allLegends = findall(fig, 'Type', 'Legend');
    for k = 1:numel(allLegends)
        lgd = allLegends(k);
        trySet(lgd, 'FontName', fontName);
        trySet(lgd, 'FontSize', legendSz);
        if forceNormalWeight
            trySet(lgd, 'FontWeight', 'normal');
        end
        if isfield(styleCfg, 'enableAxisLineWidth') && styleCfg.enableAxisLineWidth
            trySet(lgd, 'LineWidth', max(0.5, styleCfg.axisLineWidth * 0.85));
        end
        % 紧凑图例线段长度，避免小图里图例太宽
        trySet(lgd, 'ItemTokenSize', [14, 8]);
        if isprop(lgd, 'Title') && isgraphics(lgd.Title)
            trySet(lgd.Title, 'FontName', fontName);
            trySet(lgd.Title, 'FontSize', legendSz);
            if forceNormalWeight, trySet(lgd.Title, 'FontWeight', 'normal'); end
        end

        % 若原图图例是手动拖动定位 (Location='none')，自动映射到最近的内角，防止出界
        try
            if strcmpi(lgd.Location, 'none')
                pos = lgd.Position;
                cx = pos(1) + pos(3) / 2;
                cy = pos(2) + pos(4) / 2;
                if cy >= 0.5 && cx >= 0.5
                    lgd.Location = 'northeast';
                elseif cy >= 0.5 && cx < 0.5
                    lgd.Location = 'northwest';
                elseif cy < 0.5 && cx >= 0.5
                    lgd.Location = 'southeast';
                else
                    lgd.Location = 'southwest';
                end
            end
        catch
        end
    end

    % 4. 统一所有颜色条 (ColorBar)
    allColorbars = findall(fig, 'Type', 'ColorBar');
    for k = 1:numel(allColorbars)
        cb = allColorbars(k);
        trySet(cb, 'FontName', fontName);
        trySet(cb, 'FontSize', tickSz);
        if forceNormalWeight, trySet(cb, 'FontWeight', 'normal'); end
        if isprop(cb, 'Label') && isgraphics(cb.Label)
            trySet(cb.Label, 'FontName', fontName);
            trySet(cb.Label, 'FontSize', labelSz);
            if forceNormalWeight, trySet(cb.Label, 'FontWeight', 'normal'); end
            specialTextHandles(end+1) = cb.Label; %#ok<AGROW>
        end
    end

    % 5. 统一 TiledChartLayout 标题与标签
    allTiled = findall(fig, 'Type', 'TiledChartLayout');
    for k = 1:numel(allTiled)
        tl = allTiled(k);
        trySet(tl.Title,  'FontName', fontName);
        trySet(tl.Title,  'FontSize', titleSz);
        trySet(tl.XLabel, 'FontName', fontName);
        trySet(tl.XLabel, 'FontSize', labelSz);
        trySet(tl.YLabel, 'FontName', fontName);
        trySet(tl.YLabel, 'FontSize', labelSz);
        if isfield(styleCfg, 'tightLayout') && styleCfg.tightLayout
            trySet(tl, 'Padding', 'compact');
            trySet(tl, 'TileSpacing', 'compact');
        end
    end

    % 6. 统一图内所有独立文本标注 (Text & Annotation)
    allTexts = findall(fig, 'Type', 'Text');
    for k = 1:numel(allTexts)
        t = allTexts(k);
        if ~any(t == specialTextHandles)
            trySet(t, 'FontName', fontName);
            trySet(t, 'FontSize', textSz);
            if forceNormalWeight, trySet(t, 'FontWeight', 'normal'); end
        end
    end

    annotTypes = {'textboxshape', 'textarrowshape'};
    for at = 1:numel(annotTypes)
        annots = findall(fig, 'Type', annotTypes{at});
        for k = 1:numel(annots)
            trySet(annots(k), 'FontName', fontName);
            trySet(annots(k), 'FontSize', textSz);
        end
    end

    % 7. 可选：统一数据曲线线宽与小图标记点尺寸
    if isfield(styleCfg, 'enablePlotLineWidth') && styleCfg.enablePlotLineWidth
        allLines = findall(fig, 'Type', 'Line');
        for k = 1:numel(allLines)
            trySet(allLines(k), 'LineWidth', styleCfg.plotLineWidth);
            % 小图框下适当约束过大的 MarkerSize，避免数据点挤成一团
            try
                if widthCm < 7.0 && ~strcmpi(allLines(k).Marker, 'none') && allLines(k).MarkerSize > 4.5
                    allLines(k).MarkerSize = 4.5;
                end
            catch
            end
        end
    end

    % 8. 紧凑布局与防裁切：使用 OuterPosition 让 MATLAB 自动把全部标签包在物理框内
    if isfield(styleCfg, 'tightLayout') && styleCfg.tightLayout
        stdAxes = findall(fig, 'Type', 'axes');
        if numel(stdAxes) == 1 && isempty(allTiled)
            ax = stdAxes(1);
            try
                ax.Units = 'normalized';
                ax.PositionConstraint = 'outerposition';
                ax.OuterPosition = [0.005, 0.005, 0.99, 0.99];
                drawnow;

                % 二次保护：检查超长 XLabel / YLabel 是否仍超出物理画布边界，必要时微调布局
                preventLabelOverflow(ax, widthCm, heightCm);
            catch
            end
        end
    end

    drawnow;
end

function resetLabelPosition(lblObj)
    try
        if isprop(lblObj, 'PositionMode') && strcmpi(lblObj.PositionMode, 'manual')
            lblObj.Units = 'data';
            lblObj.PositionMode = 'auto';
        end
    catch
    end
end

function preventLabelOverflow(ax, widthCm, heightCm)
    % 利用 TightInset 精确计算最小边距，将坐标系内框 (innerposition) 最大化铺满图框，
    % 同时保护超长 X/Y 标签不被右/上边界裁切
    try
        ax.Units = 'centimeters';
        ti  = ax.TightInset;     % [l, b, r, t] in cm

        padL = 0.02; padB = 0.02; padR = 0.04; padT = 0.03;
        reqL = max(0.02, ti(1) + padL);
        reqB = max(0.02, ti(2) + padB);
        reqR = max(0.04, ti(3) + padR);
        reqT = max(0.03, ti(4) + padT);

        % 若无标题，进一步压缩顶部多余留白
        if isprop(ax, 'Title') && isgraphics(ax.Title) && isempty(strtrim(string(ax.Title.String)))
            reqT = min(reqT, max(0.08, ti(4) + 0.01));
        end

        % 检查 XLabel 实际物理宽度是否会超出右边界
        if isgraphics(ax.XLabel) && ~isempty(strtrim(string(ax.XLabel.String)))
            oldU = ax.XLabel.Units;
            ax.XLabel.Units = 'centimeters';
            ext = ax.XLabel.Extent; % [x, y, w, h]
            ax.XLabel.Units = oldU;
            lblW = ext(3);
            availW = max(1.0, widthCm - reqL - reqR);
            if lblW > widthCm - 0.16
                scale = (widthCm - 0.16) / lblW;
                ax.XLabel.FontSize = max(7.0, round(ax.XLabel.FontSize * scale * 10) / 10);
            elseif (reqL + availW/2 + lblW/2) > (widthCm - 0.08)
                extraR = (reqL + availW/2 + lblW/2) - (widthCm - 0.08);
                reqR = reqR + extraR;
            end
        end

        % 检查 YLabel 实际物理高度是否会超出上边界
        if isprop(ax, 'YAxis') && numel(ax.YAxis) >= 1 && isgraphics(ax.YAxis(1).Label)
            yLbl = ax.YAxis(1).Label;
            if ~isempty(strtrim(string(yLbl.String)))
                oldU = yLbl.Units;
                yLbl.Units = 'centimeters';
                ext = yLbl.Extent;
                yLbl.Units = oldU;
                lblH = ext(4);
                availH = max(0.8, heightCm - reqB - reqT);
                if lblH > heightCm - 0.16
                    scale = (heightCm - 0.16) / lblH;
                    yLbl.FontSize = max(7.0, round(yLbl.FontSize * scale * 10) / 10);
                elseif (reqB + availH/2 + lblH/2) > (heightCm - 0.08)
                    extraT = (reqB + availH/2 + lblH/2) - (heightCm - 0.08);
                    reqT = reqT + extraT;
                end
            end
        end

        % 切换为 innerposition 后精确铺满，彻底消除 MATLAB OuterPosition 的隐性白边
        newW = max(0.8, widthCm - reqL - reqR);
        newH = max(0.8, heightCm - reqB - reqT);
        ax.PositionConstraint = 'innerposition';
        ax.Position = [reqL, reqB, newW, newH];
        ax.Units = 'normalized';
    catch
        try, ax.Units = 'normalized'; catch, end
    end
end

function trySet(obj, prop, val)
    try
        if isprop(obj, prop)
            set(obj, prop, val);
        end
    catch
    end
end

function cfg = defaultStyleConfig()
    cfg = struct( ...
        'fontName',            'Times New Roman', ...
        'tickFontSize',        10, ...
        'labelFontSize',       10, ...
        'titleFontSize',       10, ...
        'legendFontSize',      9, ...
        'textFontSize',        10, ...
        'forceNormalWeight',   true, ...
        'enableAxisLineWidth', true, ...
        'axisLineWidth',       0.75, ...
        'enablePlotLineWidth', true, ...
        'plotLineWidth',       1.2, ...
        'tightLayout',         true ...
    );
end
