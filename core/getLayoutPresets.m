function slots = getLayoutPresets(presetName, canvasWidthCm, gapCm, marginCm)
%GETLAYOUTPRESETS 生成常用科研排版预设模板（坐标单位：cm，原点在画布左上角）
%
%   slots = getLayoutPresets(presetName, canvasWidthCm, gapCm, marginCm)
%   gapCm 可为标量（水平垂直同间距）或 2 元素向量 [gapX, gapY]

    if nargin < 3 || isempty(gapCm),    gapCm = [0.10, 0.12]; end
    if nargin < 4 || isempty(marginCm), marginCm = 0.0; end

    if numel(gapCm) == 1
        gapX = gapCm;
        gapY = gapCm;
    else
        gapX = gapCm(1);
        gapY = gapCm(2);
    end

    usableW = max(2.0, canvasWidthCm - 2 * marginCm);
    slots = [];

    switch presetName
        case '1x1'
            % 单图铺满（双栏论文单栏图 / 单图精准定宽）
            slots = makeRows([1], [0.72], usableW, gapX, gapY, marginCm);
        case '1x2'
            slots = makeRows([2], [0.76], usableW, gapX, gapY, marginCm);
        case '1x3'
            slots = makeRows([3], [0.82], usableW, gapX, gapY, marginCm);
        case '1x4'
            slots = makeRows([4], [0.88], usableW, gapX, gapY, marginCm);
        case '2x2'
            slots = makeRows([2, 2], [0.72, 0.72], usableW, gapX, gapY, marginCm);
        case '3+2'
            % 示例图 5：上 3 下 2，左右齐平填满版心宽度
            slots = makeRows([3, 2], [0.82, 0.56], usableW, gapX, gapY, marginCm);
        case '2+3'
            % 上 2 下 3
            slots = makeRows([2, 3], [0.56, 0.82], usableW, gapX, gapY, marginCm);
        case '2x3'
            slots = makeRows([3, 3], [0.82, 0.82], usableW, gapX, gapY, marginCm);
        case '1+2_LR'
            % 左 1 大图 + 右 2 叠放小图
            wLeft  = (usableW - gapX) * 0.56;
            wRight = usableW - gapX - wLeft;
            hRight = wRight * 0.70;
            hLeft  = 2 * hRight + gapY;
            slots = [ ...
                createSlot(1, marginCm, marginCm, wLeft, hLeft), ...
                createSlot(2, marginCm + wLeft + gapX, marginCm, wRight, hRight), ...
                createSlot(3, marginCm + wLeft + gapX, marginCm + hRight + gapY, wRight, hRight) ...
            ];
        otherwise
            slots = makeRows([3], [0.82], usableW, gapX, gapY, marginCm);
    end
end

function slots = makeRows(colCounts, aspectRatios, usableW, gapX, gapY, marginCm)
    slots = [];
    curY = marginCm;
    idx = 0;
    for r = 1:numel(colCounts)
        nCols = colCounts(r);
        boxW = (usableW - (nCols - 1) * gapX) / nCols;
        boxH = max(2.0, boxW * aspectRatios(r));
        for c = 1:nCols
            idx = idx + 1;
            bx = marginCm + (c - 1) * (boxW + gapX);
            s = createSlot(idx, bx, curY, boxW, boxH);
            if isempty(slots)
                slots = s;
            else
                slots(end+1) = s; %#ok<AGROW>
            end
        end
        curY = curY + boxH + gapY;
    end
end

function s = createSlot(id, x, y, w, h)
    s = struct( ...
        'id',       id, ...
        'x',        round(x, 2), ...
        'y',        round(y, 2), ...
        'w',        round(w, 2), ...
        'h',        round(h, 2), ...
        'filePath', '', ...
        'label',    '', ...
        'thumbImg', [] ...
    );
end
