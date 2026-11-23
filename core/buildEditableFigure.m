function [masterFig, finalWidthCm, finalHeightCm] = buildEditableFigure(slots, canvasWidthCm, canvasHeightCm, styleCfg, labelCfg, options)
%BUILDEDITABLEFIGURE 生成完全可在 MATLAB 中用属性检查器/绘图工具微调的原生 Figure (.fig)
%
%   [masterFig, finalWidthCm, finalHeightCm] = buildEditableFigure( ...
%       slots, canvasWidthCm, canvasHeightCm, styleCfg, labelCfg, options)
%
%   与位图拼合不同，本函数通过 copyobj 将所有已统一字体和物理尺寸的坐标系、曲线、
%   双 Y 轴、图例、ColorBar 完整复制到同一张精确物理尺寸 (cm) 的 MATLAB Figure 中，
%   支持直接使用 MATLAB 属性检查器 (inspect)、图例拖拽微调及另存为 .fig。

    if nargin < 5 || isempty(labelCfg)
        labelCfg = struct('enabled', true, 'style', '(a)', 'position', 'top-left-above', ...
                          'fontSize', styleCfg.labelFontSize, 'bold', true);
    end
    if nargin < 6 || isempty(options)
        options = struct('autoCropHeight', true, 'visible', true);
    end

    visStr = 'on';
    if isfield(options, 'visible') && ~options.visible
        visStr = 'off';
    end

    % 筛选已载入文件的图框
    hasFileMask = arrayfun(@(s) ~isempty(s.filePath) && isfile(s.filePath), slots);
    loadedSlots = slots(hasFileMask);
    if isempty(loadedSlots)
        error('FigureTool:NoLoadedSlots', '当前画布中还没有载入任何 .fig 或图片文件。');
    end

    isSingleMode = (numel(loadedSlots) == 1);

    if isSingleMode
        finalWidthCm  = loadedSlots(1).w;
        finalHeightCm = loadedSlots(1).h;
        activeSlots   = loadedSlots;
        activeSlots(1).x = 0;
        activeSlots(1).y = 0;
    else
        activeSlots  = loadedSlots;
        finalWidthCm = canvasWidthCm;
        if isfield(options, 'autoCropHeight') && options.autoCropHeight
            maxBottom = max(arrayfun(@(s) s.y + s.h, activeSlots));
            finalHeightCm = max(1.0, min(canvasHeightCm, maxBottom + 0.08));
        else
            finalHeightCm = canvasHeightCm;
        end
    end

    % 临时防止新建 Figure 停靠进 MATLAB Desktop
    prevDock = get(0, 'DefaultFigureWindowStyle');
    set(0, 'DefaultFigureWindowStyle', 'normal');
    cleanupDock = onCleanup(@() set(0, 'DefaultFigureWindowStyle', prevDock)); %#ok<NASGU>

    % 计算屏幕等比缩放倍数 zoomScale：
    % 仅在可见微调模式下，将「窗口物理尺寸 + 字体大小 + 线宽」同步按 zoomScale 等比放大，
    % 实现 100% 所见即所得 (WYSIWYG) 大屏微调；导出给 Word 时再按 1/zoomScale 还原精确 1:1 尺寸与字号！
    zoomScale = 1.0;
    if strcmp(visStr, 'on')
        scr = get(0, 'ScreenSize'); % [1, 1, scrW, scrH]
        dpiScreen = get(0, 'ScreenPixelsPerInch');
        if isempty(dpiScreen) || dpiScreen <= 0, dpiScreen = 96; end
        baseWpx = finalWidthCm  / 2.54 * dpiScreen;
        baseHpx = finalHeightCm / 2.54 * dpiScreen;
        maxWpx  = min(1240, round(scr(3) * 0.80));
        maxHpx  = min(780,  round(scr(4) * 0.76));
        zoomScale = min(maxWpx / max(1, baseWpx), maxHpx / max(1, baseHpx));
        zoomScale = max(1.0, min(2.4, zoomScale));
    end

    dispWcm = finalWidthCm  * zoomScale;
    dispHcm = finalHeightCm * zoomScale;

    % 按 zoomScale 同步放大样式与编号字号，保证屏幕视觉比例与最终印刷 100% 一致
    effStyleCfg = styleCfg;
    effStyleCfg.tickFontSize   = styleCfg.tickFontSize   * zoomScale;
    effStyleCfg.labelFontSize  = styleCfg.labelFontSize  * zoomScale;
    effStyleCfg.titleFontSize  = styleCfg.titleFontSize  * zoomScale;
    effStyleCfg.legendFontSize = styleCfg.legendFontSize * zoomScale;
    effStyleCfg.textFontSize   = styleCfg.textFontSize   * zoomScale;
    if isfield(styleCfg, 'axisLineWidth')
        effStyleCfg.axisLineWidth = styleCfg.axisLineWidth * zoomScale;
    end
    if isfield(styleCfg, 'plotLineWidth')
        effStyleCfg.plotLineWidth = styleCfg.plotLineWidth * zoomScale;
    end

    effLabelCfg = labelCfg;
    effLabelCfg.fontSize = labelCfg.fontSize * zoomScale;

    if strcmp(visStr, 'on')
        scr = get(0, 'ScreenSize');
        dpiScreen = get(0, 'ScreenPixelsPerInch');
        if isempty(dpiScreen) || dpiScreen <= 0, dpiScreen = 96; end
        dispWpx = round(dispWcm / 2.54 * dpiScreen);
        dispHpx = round(dispHcm / 2.54 * dpiScreen);
        posX = max(40, round((scr(3) - dispWpx) / 2));
        posY = max(60, round((scr(4) - dispHpx) / 2));
        figUnits = 'pixels';
        figPos   = [posX, posY, dispWpx, dispHpx];
    else
        figUnits = 'centimeters';
        figPos   = [3, 3, dispWcm, dispHcm];
    end

    masterFig = figure( ...
        'Visible', visStr, ...
        'WindowStyle', 'normal', ...
        'Name', sprintf('可编辑科研图 (印刷 %.2f×%.2f cm | 屏幕 %.1f× 等比放大微调)', ...
            finalWidthCm, finalHeightCm, zoomScale), ...
        'NumberTitle', 'off', ...
        'Color', 'w', ...
        'InvertHardcopy', 'off', ...
        'Units', figUnits, ...
        'Position', figPos, ...
        'PaperUnits', 'centimeters', ...
        'PaperSize', [dispWcm, dispHcm], ...
        'PaperPositionMode', 'manual', ...
        'PaperPosition', [0, 0, dispWcm, dispHcm]);

    autoLabels = computeAutoLabels(activeSlots, labelCfg.style);

    for i = 1:numel(activeSlots)
        s = activeSlots(i);

        % 计算子图编号独立条带（单图模式不预留）
        stripCm  = 0;
        plotY_cm = s.y;
        plotH_cm = s.h;
        if ~isSingleMode && labelCfg.enabled
            stripCm = max(0.38, labelCfg.fontSize * 0.045);
            if strcmp(labelCfg.position, 'top-left-above') && s.h > stripCm + 1.2
                plotY_cm = s.y + stripCm;
                plotH_cm = s.h - stripCm;
            elseif strcmp(labelCfg.position, 'bottom-center') && s.h > stripCm + 1.2
                plotH_cm = s.h - stripCm;
            end
        end

        % 转换为 MATLAB 左下角坐标系 (cm)
        slotBottomCm = finalHeightCm - (plotY_cm + plotH_cm);

        [~, ~, ext] = fileparts(s.filePath);
        if strcmpi(ext, '.fig')
            srcFig = openfig(s.filePath, 'invisible');
            try, set(srcFig, 'WindowStyle', 'normal', 'HandleVisibility', 'off'); catch, end
            cleanupSrc = onCleanup(@() safeClose(srcFig));

            styleFigure(srcFig, s.w * zoomScale, plotH_cm * zoomScale, effStyleCfg);

            % 收集源图中的所有坐标轴、图例与颜色条（同一次 copyobj 保留联动关系）
            srcAxes = [ ...
                findall(srcFig, 'Type', 'axes'); ...
                findall(srcFig, 'Type', 'polaraxes'); ...
                findall(srcFig, 'Type', 'geoaxes') ...
            ];
            srcLegends   = findall(srcFig, 'Type', 'Legend');
            srcColorbars = findall(srcFig, 'Type', 'ColorBar');

            % 记录源坐标轴在 srcFig 中的精确厘米位置
            origAxPosCm = zeros(numel(srcAxes), 4);
            for k = 1:numel(srcAxes)
                srcAxes(k).Units = 'centimeters';
                origAxPosCm(k, :) = srcAxes(k).Position;
            end

            % 逐个复制坐标轴：直接换算为 normalized 比例坐标放置，在大屏编辑窗口下等比放大且导出 1:1 精确
            for k = 1:numel(srcAxes)
                ax = srcAxes(k);
                p  = origAxPosCm(k, :);
                dstPosNorm = [ ...
                    (p(1) + s.x * zoomScale) / dispWcm, ...
                    (p(2) + slotBottomCm * zoomScale) / dispHcm, ...
                    p(3) / dispWcm, ...
                    p(4) / dispHcm ...
                ];

                isDualY = isprop(ax, 'YAxis') && numel(ax.YAxis) > 1;
                if isDualY
                    newAx = copyYyaxisAxes(ax, masterFig, dstPosNorm);
                    % 若该双 Y 轴带图例，重建可编辑图例
                    if isprop(ax, 'Legend') && ~isempty(ax.Legend) && isgraphics(ax.Legend)
                        recreateLegend(ax.Legend, newAx);
                    end
                else
                    % 找出属于当前 ax 的 Legend 与 ColorBar 一起 copyobj，保留自动联动
                    axObjs = ax;
                    for m = 1:numel(srcLegends)
                        try
                            if isequal(srcLegends(m).Axes, ax)
                                axObjs(end+1) = srcLegends(m); %#ok<AGROW>
                            end
                        catch
                        end
                    end
                    for m = 1:numel(srcColorbars)
                        try
                            if isequal(srcColorbars(m).Axes, ax)
                                axObjs(end+1) = srcColorbars(m); %#ok<AGROW>
                            end
                        catch
                        end
                    end
                    copiedObjs = copyobj(axObjs, masterFig);
                    newAx = copiedObjs(1);
                    newAx.Units = 'normalized';
                    try, newAx.PositionConstraint = 'innerposition'; catch, end
                    newAx.Position = dstPosNorm;
                end
            end

            clear cleanupSrc;
        else
            % 普通图片作为可编辑 Image Axes 插入
            rawImg = imread(s.filePath);
            axImg = axes(masterFig, 'Units', 'normalized', ...
                'Position', [s.x / finalWidthCm, slotBottomCm / finalHeightCm, ...
                             s.w / finalWidthCm, plotH_cm / finalHeightCm]);
            image(axImg, rawImg);
            axis(axImg, 'image', 'off');
        end

        % 添加可拖拽、可编辑的子图编号 (annotation textbox)
        shouldDrawLabel = false;
        lblStr = strtrim(s.label);
        if isSingleMode
            if labelCfg.enabled && ~isempty(lblStr)
                shouldDrawLabel = true;
            end
        else
            if labelCfg.enabled
                if isempty(lblStr), lblStr = autoLabels{i}; end
                shouldDrawLabel = ~isempty(lblStr);
            end
        end

        if shouldDrawLabel
            fw = 'normal';
            if labelCfg.bold, fw = 'bold'; end
            boxW = min(s.w, max(1.2, numel(lblStr) * labelCfg.fontSize * 0.04 + 0.4));
            boxH = max(0.38, labelCfg.fontSize * 0.055);
            horizAlign = 'left';

            if isSingleMode || strcmp(labelCfg.position, 'top-left-inside')
                lx = s.x + 0.06;
                ly = finalHeightCm - s.y - boxH - 0.04;
            elseif strcmp(labelCfg.position, 'bottom-center')
                lx = s.x + (s.w - boxW) / 2;
                ly = finalHeightCm - (s.y + s.h);
                horizAlign = 'center';
            else % top-left-above
                lx = s.x + 0.04;
                ly = finalHeightCm - s.y - stripCm;
            end

            nx = max(0, min(0.95, lx / finalWidthCm));
            ny = max(0, min(0.95, ly / finalHeightCm));
            nw = max(0.02, min(1 - nx, boxW / finalWidthCm));
            nh = max(0.02, min(1 - ny, boxH / finalHeightCm));

            annotation(masterFig, 'textbox', [nx, ny, nw, nh], ...
                'String', lblStr, ...
                'FontName', styleCfg.fontName, ...
                'FontSize', effLabelCfg.fontSize, ...
                'FontWeight', fw, ...
                'Color', [0, 0, 0], ...
                'EdgeColor', 'none', ...
                'BackgroundColor', 'none', ...
                'Margin', 0, ...
                'HorizontalAlignment', horizAlign, ...
                'VerticalAlignment', 'middle', ...
                'Interpreter', 'none');
        end
    end

    % 为可编辑窗口挂载便捷菜单（属性检查器 & 微调后一键按 Word 1:1 DPI 导出）
    if strcmp(visStr, 'on')
        addEditableToolsMenu(masterFig, finalWidthCm, finalHeightCm, zoomScale);
    end

    drawnow;
end

function addEditableToolsMenu(fig, wCm, hCm, zoomScale)
    try
        mRoot = uimenu(fig, 'Text', '科研微调与导出');
        uimenu(mRoot, 'Text', '打开属性检查器 (Property Inspector)', ...
            'MenuSelectedFcn', @(~,~) openInspector(fig));
        uimenu(mRoot, 'Text', '开启/关闭绘图编辑模式（拖拽图例/坐标轴/标注）', ...
            'MenuSelectedFcn', @(~,~) plotedit(fig));
        uimenu(mRoot, 'Text', '保存为可编辑 .fig 文件...', 'Separator', 'on', ...
            'MenuSelectedFcn', @(~,~) saveEditableFig(fig));
        uimenu(mRoot, 'Text', '微调后导出 300 DPI PNG (Word 1:1 尺寸)...', 'Separator', 'on', ...
            'MenuSelectedFcn', @(~,~) exportEditedBitmap(fig, wCm, hCm, zoomScale, 300, '.png'));
        uimenu(mRoot, 'Text', '微调后导出 600 DPI PNG (Word 1:1 尺寸)...', ...
            'MenuSelectedFcn', @(~,~) exportEditedBitmap(fig, wCm, hCm, zoomScale, 600, '.png'));
        uimenu(mRoot, 'Text', '微调后导出 600 DPI TIFF (Word 1:1 尺寸)...', ...
            'MenuSelectedFcn', @(~,~) exportEditedBitmap(fig, wCm, hCm, zoomScale, 600, '.tif'));
    catch
    end
end

function openInspector(fig)
    try
        plotedit(fig, 'on');
        inspect(fig);
    catch
        try, propertyeditor(fig, 'on'); catch, end
    end
end

function saveEditableFig(fig)
    [f, p] = uiputfile({'*.fig', 'MATLAB 图形 (*.fig)'}, '保存可编辑 .fig', 'Edited_Composite.fig');
    if isequal(f, 0), return; end
    savefig(fig, fullfile(p, f));
    msgbox(sprintf('已保存可编辑 .fig：\n%s', fullfile(p, f)), '保存成功');
end

function exportEditedBitmap(fig, wCm, hCm, zoomScale, dpi, ext)
    [f, p] = uiputfile({['*' ext], sprintf('高清图像 (*%s)', ext)}, ...
        sprintf('导出 %d DPI 高清图', dpi), ...
        sprintf('Edited_Figure_%.1fcm_%ddpi%s', wCm, dpi, ext));
    if isequal(f, 0), return; end
    outPath = fullfile(p, f);

    % 暂时关闭编辑选中框；按 PaperSize = [wCm*zoomScale, hCm*zoomScale] 与 dpi/zoomScale 打印，
    % 输出像素与字体大小与 1:1 物理尺寸完全等价，且 100% 保留你在屏幕上的微调比例！
    try, plotedit(fig, 'off'); catch, end
    set(fig, 'PaperUnits', 'centimeters', 'PaperSize', [wCm * zoomScale, hCm * zoomScale], ...
        'PaperPositionMode', 'manual', 'PaperPosition', [0, 0, wCm * zoomScale, hCm * zoomScale], ...
        'Color', 'w', 'InvertHardcopy', 'off');
    drawnow;

    tmpPng = [tempname '.png'];
    cleanupTmp = onCleanup(@() safeDelete(tmpPng)); %#ok<NASGU>
    printDpi = max(72, round(dpi / zoomScale));
    has3DOrRaster = ~isempty(findall(fig, 'Type', 'Surface')) || ...
                    ~isempty(findall(fig, 'Type', 'Image'));
    if ~has3DOrRaster
        try
            print(fig, tmpPng, '-dpng', '-painters', sprintf('-r%d', printDpi));
        catch
            print(fig, tmpPng, '-dpng', sprintf('-r%d', printDpi));
        end
    else
        print(fig, tmpPng, '-dpng', sprintf('-r%d', printDpi));
    end
    img = imread(tmpPng);
    if ndims(img) == 2
        img = cat(3, img, img, img);
    elseif size(img, 3) > 3
        img = img(:, :, 1:3);
    end
    saveImageWithDpi(img, outPath, dpi);
    msgbox(sprintf('已按 Word 1:1 尺寸 (%.2f × %.2f cm, %d DPI) 导出：\n%s', ...
        wCm, hCm, dpi, outPath), '导出成功');
end

function safeDelete(f)
    if isfile(f)
        try, delete(f); catch, end
    end
end

function labels = computeAutoLabels(slots, styleStr)
    n = numel(slots);
    ys = arrayfun(@(s) s.y, slots);
    xs = arrayfun(@(s) s.x, slots);
    [~, order] = sortrows([round(ys(:) / 0.6), xs(:)]);
    labels = cell(1, n);
    for rank = 1:n
        idx = order(rank);
        cL = char('a' + mod(rank - 1, 26));
        cU = char('A' + mod(rank - 1, 26));
        switch styleStr
            case '(a)', labels{idx} = sprintf('(%s)', cL);
            case '(A)', labels{idx} = sprintf('(%s)', cU);
            case 'a',   labels{idx} = cL;
            case 'A',   labels{idx} = cU;
            case 'a)',  labels{idx} = sprintf('%s)', cL);
            otherwise,  labels{idx} = sprintf('(%s)', cL);
        end
    end
end

function newAx = copyYyaxisAxes(srcAx, masterFig, dstPosNorm)
    % 无损重建 yyaxis 双 Y 轴坐标系到 masterFig（100% 保留曲线、刻度、颜色与可编辑性）
    newAx = axes(masterFig, 'Units', 'normalized', ...
        'PositionConstraint', 'innerposition', 'Position', dstPosNorm);

    % 1. 分别复制左、右 Y 轴的所有绘图子对象及轴属性
    sides = {'left', 'right'};
    for sIdx = 1:2
        side = sides{sIdx};
        yyaxis(srcAx, side);
        yyaxis(newAx, side);
        hold(newAx, 'on');

        ch = allchild(srcAx);
        if ~isempty(ch)
            copyobj(flipud(ch), newAx);
        end

        % 同步该侧 Y 轴属性
        try, newAx.YAxis(sIdx).Color     = srcAx.YAxis(sIdx).Color;     catch, end
        try, newAx.YAxis(sIdx).Limits    = srcAx.YAxis(sIdx).Limits;    catch, end
        try, newAx.YAxis(sIdx).Scale     = srcAx.YAxis(sIdx).Scale;     catch, end
        try, newAx.YAxis(sIdx).Direction = srcAx.YAxis(sIdx).Direction; catch, end
        try, newAx.YAxis(sIdx).FontName  = srcAx.YAxis(sIdx).FontName;  catch, end
        try, newAx.YAxis(sIdx).FontSize  = srcAx.YAxis(sIdx).FontSize;  catch, end
        try, newAx.YAxis(sIdx).FontWeight= srcAx.YAxis(sIdx).FontWeight;catch, end
        try
            if strcmp(srcAx.YAxis(sIdx).TickValuesMode, 'manual')
                newAx.YAxis(sIdx).TickValues = srcAx.YAxis(sIdx).TickValues;
            end
            if strcmp(srcAx.YAxis(sIdx).TickLabelsMode, 'manual')
                newAx.YAxis(sIdx).TickLabels = srcAx.YAxis(sIdx).TickLabels;
            end
        catch
        end
        try
            srcLbl = srcAx.YAxis(sIdx).Label;
            dstLbl = newAx.YAxis(sIdx).Label;
            dstLbl.String     = srcLbl.String;
            dstLbl.FontName   = srcLbl.FontName;
            dstLbl.FontSize   = srcLbl.FontSize;
            dstLbl.FontWeight = srcLbl.FontWeight;
            dstLbl.Color      = srcLbl.Color;
            dstLbl.Interpreter= srcLbl.Interpreter;
        catch
        end
    end

    % 2. 同步共享的 X 轴、网格、边框与字体属性
    sharedProps = {'FontName','FontSize','FontWeight','LineWidth','Box', ...
        'XLim','XScale','XDir','XGrid','YGrid','XMinorGrid','YMinorGrid', ...
        'GridLineStyle','GridAlpha','TickDir','Color','XColor'};
    for k = 1:numel(sharedProps)
        prop = sharedProps{k};
        try, newAx.(prop) = srcAx.(prop); catch, end
    end
    try
        if strcmp(srcAx.XTickMode, 'manual'), newAx.XTick = srcAx.XTick; end
        if strcmp(srcAx.XTickLabelMode, 'manual'), newAx.XTickLabel = srcAx.XTickLabel; end
    catch
    end
    try
        newAx.XLabel.String     = srcAx.XLabel.String;
        newAx.XLabel.FontName   = srcAx.XLabel.FontName;
        newAx.XLabel.FontSize   = srcAx.XLabel.FontSize;
        newAx.XLabel.FontWeight = srcAx.XLabel.FontWeight;
        newAx.XLabel.Color      = srcAx.XLabel.Color;
        newAx.XLabel.Interpreter= srcAx.XLabel.Interpreter;
    catch
    end
    try
        newAx.Title.String      = srcAx.Title.String;
        newAx.Title.FontName    = srcAx.Title.FontName;
        newAx.Title.FontSize    = srcAx.Title.FontSize;
        newAx.Title.FontWeight  = srcAx.Title.FontWeight;
        newAx.Title.Interpreter = srcAx.Title.Interpreter;
    catch
    end

    newAx.Units = 'normalized';
    newAx.PositionConstraint = 'innerposition';
    newAx.Position = dstPosNorm;
end

function recreateLegend(srcLgd, newAx)
    try
        lgd = legend(newAx, srcLgd.String, 'Location', srcLgd.Location);
        lgd.FontName   = srcLgd.FontName;
        lgd.FontSize   = srcLgd.FontSize;
        lgd.FontWeight = srcLgd.FontWeight;
        lgd.Box        = srcLgd.Box;
        lgd.LineWidth  = srcLgd.LineWidth;
        try, lgd.ItemTokenSize = srcLgd.ItemTokenSize; catch, end
    catch
    end
end

function safeClose(fig)
    if ~isempty(fig) && isgraphics(fig)
        try, close(fig); catch, end
    end
end
