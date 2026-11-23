function info = parseWordDocx(docxPath)
%PARSEWORDDOCX 解析 Word (.docx) 文件的页面版心宽度、分栏设置和默认正文字体字号
%
%   info = parseWordDocx(docxPath)
%
%   返回结构体 info 包含：
%     .fileName          - 文件名
%     .pageWidthCm       - 页面总宽度 (cm)
%     .pageHeightCm      - 页面总高度 (cm)
%     .marginLeftCm      - 左页边距 (cm)
%     .marginRightCm     - 右页边距 (cm)
%     .marginTopCm       - 上页边距 (cm)
%     .marginBottomCm    - 下页边距 (cm)
%     .numColumns        - 栏数 (1=单栏, 2=双栏, ...)
%     .columnSpaceCm     - 栏间距 (cm)
%     .fullTextWidthCm   - 通栏版心宽度 (cm) = 页面宽 - 左边距 - 右边距
%     .columnWidthCm     - 单栏宽度 (cm)
%     .fullTextHeightCm  - 版心高度 (cm)
%     .bodyFontSizePt    - 识别到的正文字号 (pt)
%     .bodyFontName      - 识别到的英文字体名
%     .eastAsiaFontName  - 识别到的中文字体名
%     .summary           - 人类可读的解析摘要

    arguments
        docxPath (1,:) char
    end

    if ~isfile(docxPath)
        error('FigureTool:FileNotFound', '找不到 Word 文件: %s', docxPath);
    end

    [~, name, ext] = fileparts(docxPath);
    if ~strcmpi(ext, '.docx')
        error('FigureTool:InvalidFormat', '仅支持 .docx 格式的 Word 文档（当前为 %s）', ext);
    end

    % 创建临时解压目录，退出时自动清理
    tmpDir = tempname;
    mkdir(tmpDir);
    cleanupObj = onCleanup(@() safeRmdir(tmpDir)); %#ok<NASGU>

    try
        unzip(docxPath, tmpDir);
    catch ME
        error('FigureTool:UnzipFailed', '无法解压 .docx 文件，请确认文件未加密或损坏。\n详情: %s', ME.message);
    end

    % 默认值（A4 标准页面）
    TWIPS_TO_CM = 2.54 / 1440.0;
    pageWidthCm      = 21.0;
    pageHeightCm     = 29.7;
    marginLeftCm     = 3.175;  % 1.25 in
    marginRightCm    = 3.175;
    marginTopCm      = 2.54;   % 1.0 in
    marginBottomCm   = 2.54;
    numColumns       = 1;
    columnSpaceCm    = 0.75;
    bodyFontSizePt   = 10.0;
    bodyFontName     = 'Times New Roman';
    eastAsiaFontName = '';

    % 1. 解析 word/document.xml 中的页面设置 (sectPr)
    docXmlPath = fullfile(tmpDir, 'word', 'document.xml');
    if isfile(docXmlPath)
        docXml = fileread(docXmlPath);

        % 提取所有 sectPr，优先使用文档末尾的主节设置
        sectMatches = regexp(docXml, '<w:sectPr\b[^>]*>.*?</w:sectPr>', 'match');
        if ~isempty(sectMatches)
            sectXml = sectMatches{end};

            % 页面尺寸 w:pgSz
            pgSzTag = regexp(sectXml, '<w:pgSz\b[^>]*/?>', 'match', 'once');
            if ~isempty(pgSzTag)
                wVal = extractXmlNumAttr(pgSzTag, 'w:w');
                hVal = extractXmlNumAttr(pgSzTag, 'w:h');
                if ~isnan(wVal) && wVal > 0
                    pageWidthCm = wVal * TWIPS_TO_CM;
                end
                if ~isnan(hVal) && hVal > 0
                    pageHeightCm = hVal * TWIPS_TO_CM;
                end
            end

            % 页边距 w:pgMar
            pgMarTag = regexp(sectXml, '<w:pgMar\b[^>]*/?>', 'match', 'once');
            if ~isempty(pgMarTag)
                lVal = extractXmlNumAttr(pgMarTag, 'w:left');
                rVal = extractXmlNumAttr(pgMarTag, 'w:right');
                tVal = extractXmlNumAttr(pgMarTag, 'w:top');
                bVal = extractXmlNumAttr(pgMarTag, 'w:bottom');
                if ~isnan(lVal), marginLeftCm   = lVal * TWIPS_TO_CM; end
                if ~isnan(rVal), marginRightCm  = rVal * TWIPS_TO_CM; end
                if ~isnan(tVal), marginTopCm    = tVal * TWIPS_TO_CM; end
                if ~isnan(bVal), marginBottomCm = bVal * TWIPS_TO_CM; end
            end

            % 分栏设置 w:cols
            colsTag = regexp(sectXml, '<w:cols\b[^>]*/?>', 'match', 'once');
            if ~isempty(colsTag)
                nVal = extractXmlNumAttr(colsTag, 'w:num');
                sVal = extractXmlNumAttr(colsTag, 'w:space');
                if ~isnan(nVal) && nVal >= 1
                    numColumns = round(nVal);
                end
                if ~isnan(sVal) && sVal >= 0
                    columnSpaceCm = sVal * TWIPS_TO_CM;
                end
            end
        end
    end

    % 2. 解析 word/styles.xml 中的正文字体与字号
    stylesXmlPath = fullfile(tmpDir, 'word', 'styles.xml');
    foundSize = false;
    foundFont = false;
    if isfile(stylesXmlPath)
        stylesXml = fileread(stylesXmlPath);

        % 优先查找 Normal 样式 (w:styleId="Normal" 或默认段落样式)
        normalBlock = regexp(stylesXml, ...
            '<w:style\b[^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>.*?</w:style>', ...
            'match', 'once');
        if isempty(normalBlock)
            normalBlock = regexp(stylesXml, ...
                '<w:style\b[^>]*w:styleId="Normal"[^>]*>.*?</w:style>', ...
                'match', 'once');
        end

        if ~isempty(normalBlock)
            [szPt, fAscii, fEA] = extractFontFromBlock(normalBlock);
            if ~isnan(szPt)
                bodyFontSizePt = szPt;
                foundSize = true;
            end
            if ~isempty(fAscii)
                bodyFontName = fAscii;
                foundFont = true;
            end
            if ~isempty(fEA)
                eastAsiaFontName = fEA;
            end
        end

        % 其次查找 docDefaults
        if ~foundSize || ~foundFont
            defaultsBlock = regexp(stylesXml, ...
                '<w:docDefaults\b[^>]*>.*?</w:docDefaults>', 'match', 'once');
            if ~isempty(defaultsBlock)
                [szPt, fAscii, fEA] = extractFontFromBlock(defaultsBlock);
                if ~foundSize && ~isnan(szPt)
                    bodyFontSizePt = szPt;
                    foundSize = true;
                end
                if ~foundFont && ~isempty(fAscii)
                    bodyFontName = fAscii;
                    foundFont = true;
                end
                if isempty(eastAsiaFontName) && ~isempty(fEA)
                    eastAsiaFontName = fEA;
                end
            end
        end
    end

    % 如果 styles.xml 中未找到字号，统计 document.xml 中最常出现的正文字号
    if ~foundSize && isfile(docXmlPath)
        szMatches = regexp(docXml, '<w:sz\b[^>]*w:val="(\d+)"', 'tokens');
        if ~isempty(szMatches)
            vals = cellfun(@(c) str2double(c{1}) / 2.0, szMatches);
            vals = vals(vals >= 7 & vals <= 16); % 过滤标题/上标等极端字号
            if ~isempty(vals)
                bodyFontSizePt = mode(vals);
            end
        end
    end

    % 3. 计算版心宽度与单栏宽度
    fullTextWidthCm  = max(2.0, pageWidthCm - marginLeftCm - marginRightCm);
    fullTextHeightCm = max(2.0, pageHeightCm - marginTopCm - marginBottomCm);
    if numColumns > 1
        columnWidthCm = max(2.0, (fullTextWidthCm - (numColumns - 1) * columnSpaceCm) / numColumns);
    else
        columnWidthCm = fullTextWidthCm;
    end

    % 规范化常见字体名称
    bodyFontName = normalizeFontName(bodyFontName);

    info = struct();
    info.fileName         = [name ext];
    info.pageWidthCm      = round(pageWidthCm, 2);
    info.pageHeightCm     = round(pageHeightCm, 2);
    info.marginLeftCm     = round(marginLeftCm, 2);
    info.marginRightCm    = round(marginRightCm, 2);
    info.marginTopCm      = round(marginTopCm, 2);
    info.marginBottomCm   = round(marginBottomCm, 2);
    info.numColumns       = numColumns;
    info.columnSpaceCm    = round(columnSpaceCm, 2);
    info.fullTextWidthCm  = round(fullTextWidthCm, 2);
    info.columnWidthCm    = round(columnWidthCm, 2);
    info.fullTextHeightCm = round(fullTextHeightCm, 2);
    info.bodyFontSizePt   = round(bodyFontSizePt, 1);
    info.bodyFontName     = bodyFontName;
    info.eastAsiaFontName = eastAsiaFontName;

    if numColumns > 1
        colDesc = sprintf('%d 栏（通栏 %.2f cm / 单栏 %.2f cm）', ...
            numColumns, info.fullTextWidthCm, info.columnWidthCm);
    else
        colDesc = sprintf('单栏（版心宽 %.2f cm）', info.fullTextWidthCm);
    end

    info.summary = sprintf('%s | 页面 %.2f×%.2f cm | %s | 正文字号 %.1f pt (%s)', ...
        info.fileName, info.pageWidthCm, info.pageHeightCm, ...
        colDesc, info.bodyFontSizePt, info.bodyFontName);
end

function val = extractXmlNumAttr(tagStr, attrName)
    val = NaN;
    pattern = [attrName '="(-?\d+(?:\.\d+)?)"'];
    tok = regexp(tagStr, pattern, 'tokens', 'once');
    if ~isempty(tok)
        val = str2double(tok{1});
    end
end

function [szPt, fAscii, fEA] = extractFontFromBlock(xmlBlock)
    szPt = NaN;
    fAscii = '';
    fEA = '';

    szTok = regexp(xmlBlock, '<w:sz\b[^>]*w:val="(\d+)"', 'tokens', 'once');
    if ~isempty(szTok)
        szPt = str2double(szTok{1}) / 2.0;
    end

    rFontsTag = regexp(xmlBlock, '<w:rFonts\b[^>]*/?>', 'match', 'once');
    if ~isempty(rFontsTag)
        tokA = regexp(rFontsTag, 'w:ascii="([^"]+)"', 'tokens', 'once');
        if ~isempty(tokA)
            fAscii = tokA{1};
        end
        tokEA = regexp(rFontsTag, 'w:eastAsia="([^"]+)"', 'tokens', 'once');
        if ~isempty(tokEA)
            fEA = tokEA{1};
        end
    end
end

function fontName = normalizeFontName(rawName)
    fontName = strtrim(rawName);
    if isempty(fontName)
        fontName = 'Times New Roman';
        return;
    end
    % 检查当前系统可用字体，尽量匹配
    try
        avail = listfonts();
        if any(strcmpi(avail, fontName))
            idx = find(strcmpi(avail, fontName), 1);
            fontName = avail{idx};
        elseif contains(lower(fontName), 'times')
            fontName = 'Times New Roman';
        elseif contains(lower(fontName), 'arial')
            fontName = 'Arial';
        elseif contains(lower(fontName), 'helvetica')
            fontName = 'Helvetica';
        end
    catch
        % 忽略字体列举错误
    end
end

function safeRmdir(d)
    if isfolder(d)
        try
            rmdir(d, 's');
        catch
        end
    end
end
