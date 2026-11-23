function imgRgb = renderSingleSlot(filePath, widthCm, heightCm, dpi, styleCfg)
%RENDERSINGLESLOT 按精确物理尺寸 (cm) 和 DPI 渲染单个 .fig（或普通图片）为 RGB 图像
%
%   imgRgb = renderSingleSlot(filePath, widthCm, heightCm, dpi, styleCfg)

    targetW = max(16, round(widthCm / 2.54 * dpi));
    targetH = max(16, round(heightCm / 2.54 * dpi));

    if isempty(filePath) || ~isfile(filePath)
        imgRgb = uint8(255 * ones(targetH, targetW, 3));
        return;
    end

    [~, ~, ext] = fileparts(filePath);
    if strcmpi(ext, '.fig')
        % 临时阻止图形默认停靠到 MATLAB Desktop，防止后台渲染唤醒最小化的 MATLAB 主窗口
        prevDock = get(0, 'DefaultFigureWindowStyle');
        set(0, 'DefaultFigureWindowStyle', 'normal');
        cleanupDock = onCleanup(@() set(0, 'DefaultFigureWindowStyle', prevDock)); %#ok<NASGU>

        fig = openfig(filePath, 'invisible');
        try, set(fig, 'WindowStyle', 'normal', 'HandleVisibility', 'off'); catch, end
        cleanupFig = onCleanup(@() safeClose(fig)); %#ok<NASGU>

        styleFigure(fig, widthCm, heightCm, styleCfg);

        tmpPng = [tempname '.png'];
        cleanupTmp = onCleanup(@() safeDelete(tmpPng)); %#ok<NASGU>

        % 使用 print 严格按照 PaperPosition (cm) 与 -rDPI 输出 1:1 物理尺寸位图
        print(fig, tmpPng, '-dpng', sprintf('-r%d', round(dpi)));

        rawImg = imread(tmpPng);
        rawImg = ensureRgbUint8(rawImg);

        % 对齐到精确目标像素尺寸（仅做极小 ±1~2px 裁剪或留白，绝不拉伸变形）
        imgRgb = fitExactWithoutStretch(rawImg, targetH, targetW);
    else
        % 普通图片：保持原始长宽比居中放入目标框
        rawImg = imread(filePath);
        rawImg = ensureRgbUint8(rawImg);
        imgRgb = letterboxImage(rawImg, targetH, targetW);
    end
end

function rgb = ensureRgbUint8(img)
    if isa(img, 'uint16')
        img = uint8(double(img) / 257);
    elseif isfloat(img)
        if max(img(:)) <= 1.0
            img = uint8(round(img * 255));
        else
            img = uint8(img);
        end
    elseif islogical(img)
        img = uint8(img) * 255;
    end

    if ndims(img) == 2
        rgb = cat(3, img, img, img);
    elseif size(img, 3) >= 3
        rgb = img(:, :, 1:3);
    else
        rgb = cat(3, img(:,:,1), img(:,:,1), img(:,:,1));
    end
end

function out = fitExactWithoutStretch(src, targetH, targetW)
    [sh, sw, ~] = size(src);
    if sh == targetH && sw == targetW
        out = src;
        return;
    end
    out = uint8(255 * ones(targetH, targetW, 3));
    copyH = min(sh, targetH);
    copyW = min(sw, targetW);
    % 居中放置
    dstY = floor((targetH - copyH) / 2) + 1;
    dstX = floor((targetW - copyW) / 2) + 1;
    srcY = floor((sh - copyH) / 2) + 1;
    srcX = floor((sw - copyW) / 2) + 1;
    out(dstY:dstY+copyH-1, dstX:dstX+copyW-1, :) = ...
        src(srcY:srcY+copyH-1, srcX:srcX+copyW-1, :);
end

function out = letterboxImage(src, targetH, targetW)
    [sh, sw, ~] = size(src);
    scale = min(targetW / sw, targetH / sh);
    newW = max(1, round(sw * scale));
    newH = max(1, round(sh * scale));

    if exist('imresize', 'file') == 2
        resized = imresize(src, [newH, newW], 'bicubic');
    else
        % 纯 MATLAB 内置插值回退（不依赖 Image Processing Toolbox）
        yGrid = linspace(1, sh, newH);
        xGrid = linspace(1, sw, newW);
        [X, Y] = meshgrid(xGrid, yGrid);
        resized = zeros(newH, newW, 3, 'uint8');
        for c = 1:3
            resized(:,:,c) = uint8(round(interp2(double(src(:,:,c)), X, Y, 'linear')));
        end
    end

    out = uint8(255 * ones(targetH, targetW, 3));
    offY = floor((targetH - newH) / 2) + 1;
    offX = floor((targetW - newW) / 2) + 1;
    out(offY:offY+newH-1, offX:offX+newW-1, :) = resized;
end

function safeClose(fig)
    if ~isempty(fig) && isgraphics(fig)
        try
            close(fig);
        catch
        end
    end
end

function safeDelete(f)
    if isfile(f)
        try
            delete(f);
        catch
        end
    end
end
