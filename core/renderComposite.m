function [canvasRgb, finalWidthCm, finalHeightCm] = renderComposite(slots, canvasWidthCm, canvasHeightCm, dpi, styleCfg, labelCfg, options)
%RENDERCOMPOSITE 将所有排版图框按精确厘米坐标与 DPI 拼合为最终高清科研大图
%
%   [canvasRgb, finalWidthCm, finalHeightCm] = renderComposite( ...
%       slots, canvasWidthCm, canvasHeightCm, dpi, styleCfg, labelCfg, options)

    if nargin < 6 || isempty(labelCfg)
        labelCfg = struct('enabled', true, 'style', '(a)', 'position', 'top-left-above', ...
                          'fontSize', styleCfg.labelFontSize, 'bold', true);
    end
    if nargin < 7 || isempty(options)
        options = struct('autoCropHeight', true, 'saveIndividuals', false, ...
                         'individualDir', '', 'individualExt', '.png', 'progressFcn', []);
    end

    if isempty(slots)
        finalWidthCm = canvasWidthCm;
        finalHeightCm = canvasHeightCm;
        canvasRgb = uint8(255 * ones(max(32, round(finalHeightCm/2.54*dpi)), ...
                                     max(32, round(finalWidthCm/2.54*dpi)), 3));
        return;
    end

    % 仅筛选已实际载入文件的图框，忽略空白占位框
    hasFileMask = arrayfun(@(s) ~isempty(s.filePath) && isfile(s.filePath), slots);
    loadedSlots = slots(hasFileMask);

    % 单图快速路径：如果只放了 1 张图，直接只渲染这一张（精确按该图框的 w × h cm 输出，不留空框白边、不加多余的 (a)）
    if numel(loadedSlots) == 1
        s = loadedSlots(1);
        if isfield(options, 'progressFcn') && ~isempty(options.progressFcn)
            try
                options.progressFcn(1, 1, '正在以 1:1 物理尺寸渲染单图...');
            catch
            end
        end
        finalWidthCm  = s.w;
        finalHeightCm = s.h;
        canvasRgb = renderSingleSlot(s.filePath, s.w, s.h, dpi, styleCfg);

        % 若用户明确手动填写了编号覆写，才叠加编号；默认单图不加 (a)
        customLbl = strtrim(s.label);
        if labelCfg.enabled && ~isempty(customLbl)
            lblPatch = renderLabelPatch(customLbl, styleCfg.fontName, ...
                labelCfg.fontSize, labelCfg.bold, dpi);
            padPx = round(0.06 / 2.54 * dpi);
            canvasRgb = pastePatch(canvasRgb, lblPatch, 1 + padPx, 1 + padPx, true);
        end

        if isfield(options, 'saveIndividuals') && options.saveIndividuals && ...
           isfield(options, 'individualDir') && ~isempty(options.individualDir)
            [~, baseName, ~] = fileparts(s.filePath);
            subFile = fullfile(options.individualDir, sprintf('sub_01_%s%s', baseName, options.individualExt));
            saveImageWithDpi(canvasRgb, subFile, dpi);
        end
        return;
    end

    % 多图模式：只处理已载入文件的图框（若全部为空则保留原空白画布）
    if ~isempty(loadedSlots)
        activeSlots = loadedSlots;
    else
        activeSlots = slots;
    end

    % 计算最终画布高度（若开启自动紧凑裁剪高度，则贴合最底部已载入图框）
    finalWidthCm = canvasWidthCm;
    if isfield(options, 'autoCropHeight') && options.autoCropHeight
        maxBottom = max(arrayfun(@(s) s.y + s.h, activeSlots));
        extraBottom = 0.08;
        finalHeightCm = max(1.0, min(canvasHeightCm, maxBottom + extraBottom));
    else
        finalHeightCm = canvasHeightCm;
    end

    W_px = max(32, round(finalWidthCm / 2.54 * dpi));
    H_px = max(32, round(finalHeightCm / 2.54 * dpi));
    canvasRgb = uint8(255 * ones(H_px, W_px, 3));

    % 按阅读顺序（先上后下、同行先左后右）生成默认自动编号
    autoLabels = computeAutoLabels(activeSlots, labelCfg.style);

    nSlots = numel(activeSlots);
    for i = 1:nSlots
        s = activeSlots(i);
        if isfield(options, 'progressFcn') && ~isempty(options.progressFcn)
            try
                options.progressFcn(i, nSlots, sprintf('正在渲染子图 %d/%d ...', i, nSlots));
            catch
            end
        end

        % 根据子图编号位置预留独立文字条带，彻底避免 (a)(b)(c) 与 Y 轴标签重叠
        stripCm = 0;
        plotY_cm = s.y;
        plotH_cm = s.h;
        if labelCfg.enabled
            stripCm = max(0.38, labelCfg.fontSize * 0.045);
            if strcmp(labelCfg.position, 'top-left-above') && s.h > stripCm + 1.2
                plotY_cm = s.y + stripCm;
                plotH_cm = s.h - stripCm;
            elseif strcmp(labelCfg.position, 'bottom-center') && s.h > stripCm + 1.2
                plotH_cm = s.h - stripCm;
            end
        end

        % 1. 渲染该图框（1:1 物理尺寸）
        if ~isempty(s.filePath) && isfile(s.filePath)
            subImg = renderSingleSlot(s.filePath, s.w, plotH_cm, dpi, styleCfg);

            % 可选：保存单独子图文件
            if isfield(options, 'saveIndividuals') && options.saveIndividuals && ...
               isfield(options, 'individualDir') && ~isempty(options.individualDir)
                [~, baseName, ~] = fileparts(s.filePath);
                subFile = fullfile(options.individualDir, ...
                    sprintf('sub_%02d_%s%s', i, baseName, options.individualExt));
                saveImageWithDpi(subImg, subFile, dpi);
            end

            % 将子图精准粘贴到画布对应像素位置
            x0 = round(s.x / 2.54 * dpi) + 1;
            y0 = round(plotY_cm / 2.54 * dpi) + 1;
            canvasRgb = pastePatch(canvasRgb, subImg, x0, y0, false);
        end

        % 2. 叠加子图编号 (a), (b)...
        if labelCfg.enabled
            lblStr = strtrim(s.label);
            if isempty(lblStr)
                lblStr = autoLabels{i};
            end
            if ~isempty(lblStr)
                lblPatch = renderLabelPatch(lblStr, styleCfg.fontName, ...
                    labelCfg.fontSize, labelCfg.bold, dpi);
                [lh, lw, ~] = size(lblPatch);
                x0 = round(s.x / 2.54 * dpi) + 1;
                y0 = round(s.y / 2.54 * dpi) + 1;
                sw = round(s.w / 2.54 * dpi);
                sh = round(s.h / 2.54 * dpi);
                stripPx = round(stripCm / 2.54 * dpi);
                padPx = round(0.06 / 2.54 * dpi);

                switch labelCfg.position
                    case 'top-left-above'
                        lx = x0 + padPx;
                        ly = y0 + max(0, round((stripPx - lh) / 2));
                    case 'bottom-center'
                        lx = x0 + round((sw - lw) / 2);
                        ly = y0 + sh - stripPx + max(0, round((stripPx - lh) / 2));
                    case 'top-left-inside'
                        lx = x0 + padPx;
                        ly = y0 + padPx;
                    otherwise
                        lx = x0 + padPx;
                        ly = y0 + max(0, round((stripPx - lh) / 2));
                end
                canvasRgb = pastePatch(canvasRgb, lblPatch, lx, ly, true);
            end
        end
    end
end

function labels = computeAutoLabels(slots, styleStr)
    n = numel(slots);
    % 按 y (容差 0.6 cm 视为同一行) 再按 x 排序
    ys = arrayfun(@(s) s.y, slots);
    xs = arrayfun(@(s) s.x, slots);
    rowKey = round(ys / 0.6);
    [~, order] = sortrows([rowKey(:), xs(:)]);

    labels = cell(1, n);
    for rank = 1:n
        idx = order(rank);
        chLower = char('a' + mod(rank - 1, 26));
        chUpper = char('A' + mod(rank - 1, 26));
        switch styleStr
            case '(a)'
                labels{idx} = sprintf('(%s)', chLower);
            case '(A)'
                labels{idx} = sprintf('(%s)', chUpper);
            case 'a'
                labels{idx} = chLower;
            case 'A'
                labels{idx} = chUpper;
            case 'a)'
                labels{idx} = sprintf('%s)', chLower);
            otherwise
                labels{idx} = sprintf('(%s)', chLower);
        end
    end
end

function patchRgb = renderLabelPatch(lblStr, fontName, fontSizePt, isBold, dpi)
    % 用不可见 MATLAB Figure 以精确物理字号渲染文字贴片
    patchW = max(1.2, numel(lblStr) * fontSizePt * 0.038 + 0.5); % cm
    patchH = max(0.6, fontSizePt * 0.065 + 0.25);                % cm

    fig = figure('Visible', 'off', 'WindowStyle', 'normal', 'HandleVisibility', 'off', ...
        'Color', 'w', 'InvertHardcopy', 'off', ...
        'Units', 'centimeters', 'Position', [2, 2, patchW, patchH], ...
        'PaperUnits', 'centimeters', 'PaperSize', [patchW, patchH], ...
        'PaperPositionMode', 'manual', 'PaperPosition', [0, 0, patchW, patchH]);
    cleanupFig = onCleanup(@() safeClose(fig)); %#ok<NASGU>

    ax = axes(fig, 'Units', 'normalized', 'Position', [0, 0, 1, 1], ...
        'Visible', 'off', 'XLim', [0, 1], 'YLim', [0, 1]);
    fw = 'normal';
    if isBold, fw = 'bold'; end
    text(ax, 0.5, 0.5, lblStr, ...
        'FontName', fontName, ...
        'FontSize', fontSizePt, ...
        'FontWeight', fw, ...
        'Color', [0, 0, 0], ...
        'HorizontalAlignment', 'center', ...
        'VerticalAlignment', 'middle', ...
        'Interpreter', 'none');
    drawnow;

    tmpPng = [tempname '.png'];
    cleanupTmp = onCleanup(@() safeDelete(tmpPng)); %#ok<NASGU>
    print(fig, tmpPng, '-dpng', sprintf('-r%d', round(dpi)));
    raw = imread(tmpPng);
    if ndims(raw) == 2
        raw = cat(3, raw, raw, raw);
    elseif size(raw, 3) > 3
        raw = raw(:, :, 1:3);
    end

    % 裁剪文字周围空白，保留极小 2px 边距
    nonWhite = any(raw < 248, 3);
    [rows, cols] = find(nonWhite);
    if isempty(rows)
        patchRgb = raw;
        return;
    end
    r1 = max(1, min(rows) - 2);
    r2 = min(size(raw, 1), max(rows) + 2);
    c1 = max(1, min(cols) - 2);
    c2 = min(size(raw, 2), max(cols) + 2);
    patchRgb = raw(r1:r2, c1:c2, :);
end

function canvas = pastePatch(canvas, patch, x0, y0, transparentWhite)
    [ch, cw, ~] = size(canvas);
    [ph, pw, ~] = size(patch);

    dstX1 = max(1, x0);
    dstY1 = max(1, y0);
    dstX2 = min(cw, x0 + pw - 1);
    dstY2 = min(ch, y0 + ph - 1);

    if dstX2 < dstX1 || dstY2 < dstY1
        return;
    end

    srcX1 = dstX1 - x0 + 1;
    srcY1 = dstY1 - y0 + 1;
    srcX2 = srcX1 + (dstX2 - dstX1);
    srcY2 = srcY1 + (dstY2 - dstY1);

    subPatch = patch(srcY1:srcY2, srcX1:srcX2, :);
    if transparentWhite
        bg = canvas(dstY1:dstY2, dstX1:dstX2, :);
        mask = any(subPatch < 248, 3);
        for c = 1:3
            chDst = bg(:, :, c);
            chSrc = subPatch(:, :, c);
            chDst(mask) = chSrc(mask);
            bg(:, :, c) = chDst;
        end
        canvas(dstY1:dstY2, dstX1:dstX2, :) = bg;
    else
        canvas(dstY1:dstY2, dstX1:dstX2, :) = subPatch;
    end
end

function safeClose(fig)
    if ~isempty(fig) && isgraphics(fig)
        try, close(fig); catch, end
    end
end

function safeDelete(f)
    if isfile(f)
        try, delete(f); catch, end
    end
end
