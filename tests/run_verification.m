function run_verification()
%RUN_VERIFICATION 端到端自动化测试（解析 .docx + 统一 .fig 字体 + 3+2 拼图导出 + DPI 校验 + App 启动测试）

    rootDir = fileparts(fileparts(mfilename('fullpath')));
    addpath(rootDir);
    addpath(fullfile(rootDir, 'core'));

    outDir = fullfile(rootDir, 'tests', 'output');
    if ~isfolder(outDir), mkdir(outDir); end

    fprintf('=== 1. 生成测试 Word (.docx) 并验证解析 ===\n');
    sampleDocx = createSampleDocx(outDir);
    info = parseWordDocx(sampleDocx);
    fprintf('Word 解析摘要: %s\n', info.summary);
    assert(abs(info.fullTextWidthCm - 14.65) < 0.05, ...
        '期望版心宽约 14.65 cm，实际 %.2f cm', info.fullTextWidthCm);
    assert(abs(info.bodyFontSizePt - 10.0) < 0.1, ...
        '期望正文字号 10.0 pt，实际 %.1f pt', info.bodyFontSizePt);
    fprintf('[PASS] Word .docx 版心与正文字号解析通过！\n\n');

    fprintf('=== 2. 生成 5 个原始尺寸和字号完全不同的测试 .fig ===\n');
    figPaths = createSampleFigs(outDir);
    fprintf('[PASS] 已生成 5 个测试 .fig 文件。\n\n');

    fprintf('=== 3. 应用 3+2 排版模板并统一为 %.1f pt (%s) 渲染拼图 ===\n', ...
        info.bodyFontSizePt, info.bodyFontName);
    slots = getLayoutPresets('3+2', info.fullTextWidthCm, 0.30, 0.0);
    assert(numel(slots) == 5, '3+2 模板应有 5 个图框');
    for i = 1:5
        slots(i).filePath = figPaths{i};
    end

    styleCfg = struct( ...
        'fontName',            'Times New Roman', ...
        'tickFontSize',        10.0, ...
        'labelFontSize',       10.0, ...
        'titleFontSize',       10.0, ...
        'legendFontSize',      9.0, ...
        'textFontSize',        10.0, ...
        'enableAxisLineWidth', true, ...
        'axisLineWidth',       0.75, ...
        'enablePlotLineWidth', false, ...
        'plotLineWidth',       1.2, ...
        'tightLayout',         true ...
    );
    labelCfg = struct( ...
        'enabled',  true, ...
        'style',    '(a)', ...
        'position', 'top-left-inside', ...
        'fontSize', 10.0, ...
        'bold',     true ...
    );
    opts = struct( ...
        'autoCropHeight',  true, ...
        'saveIndividuals', false ...
    );

    dpi = 300;
    [canvasRgb, finalW, finalH] = renderComposite(slots, ...
        info.fullTextWidthCm, 20.0, dpi, styleCfg, labelCfg, opts);

    expectedWpx = round(finalW / 2.54 * dpi);
    expectedHpx = round(finalH / 2.54 * dpi);
    assert(size(canvasRgb, 2) == expectedWpx && size(canvasRgb, 1) == expectedHpx, ...
        '输出像素尺寸不匹配！');

    pngOut = fullfile(outDir, 'verify_composite_3plus2_300dpi.png');
    tifOut = fullfile(outDir, 'verify_composite_3plus2_300dpi.tif');
    saveImageWithDpi(canvasRgb, pngOut, dpi);
    saveImageWithDpi(canvasRgb, tifOut, dpi);

    % 校验 PNG & TIFF 元数据中的物理分辨率
    pngInfo = imfinfo(pngOut);
    tifInfo = imfinfo(tifOut);
    pngDpi = pngInfo.XResolution * 0.0254;
    fprintf('PNG 尺寸: %dx%d px, 元数据 DPI = %.1f (物理宽 %.2f cm)\n', ...
        pngInfo.Width, pngInfo.Height, pngDpi, pngInfo.Width / pngDpi * 2.54);
    fprintf('TIFF 尺寸: %dx%d px, 元数据 DPI = %.1f (物理宽 %.2f cm)\n', ...
        tifInfo.Width, tifInfo.Height, tifInfo.XResolution, tifInfo.Width / tifInfo.XResolution * 2.54);
    assert(abs(pngDpi - dpi) < 1.0, 'PNG DPI 元数据不准确');
    assert(abs(tifInfo.XResolution - dpi) < 1.0, 'TIFF DPI 元数据不准确');
    fprintf('[PASS] 3+2 拼图渲染与 PNG/TIFF DPI 元数据校验通过！\n\n');

    fprintf('=== 4. 验证 FigureLayoutApp 实例化与界面构建 ===\n');
    app = FigureLayoutApp();
    drawnow;
    assert(isvalid(app.UIFigure), 'App 窗口创建失败');
    delete(app);
    fprintf('[PASS] FigureLayoutApp 启动与销毁测试通过！\n');
    fprintf('\n=== 全部自动化测试通过 ===\n');
end

function docxPath = createSampleDocx(outDir)
    % 构造标准 A4 (21.0 cm) + 左右页边距 3.175 cm -> 版心宽 14.65 cm, 正文 10 pt Times New Roman
    tmpDir = tempname;
    mkdir(tmpDir);
    mkdir(fullfile(tmpDir, 'word'));

    % 21.0 cm = 11906 twips, 29.7 cm = 16838 twips, 3.175 cm = 1800 twips
    docXml = [ ...
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ...
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' ...
        '<w:body><w:p><w:r><w:t>Sample Paper</w:t></w:r></w:p>' ...
        '<w:sectPr>' ...
        '<w:pgSz w:w="11906" w:h="16838"/>' ...
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>' ...
        '<w:cols w:num="1" w:space="425"/>' ...
        '</w:sectPr></w:body></w:document>'];

    stylesXml = [ ...
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ...
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' ...
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' ...
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun"/><w:sz w:val="20"/></w:rPr>' ...
        '</w:style></w:styles>'];

    fid = fopen(fullfile(tmpDir, 'word', 'document.xml'), 'w');
    fwrite(fid, docXml, 'char'); fclose(fid);
    fid = fopen(fullfile(tmpDir, 'word', 'styles.xml'), 'w');
    fwrite(fid, stylesXml, 'char'); fclose(fid);

    zipPath = fullfile(outDir, 'sample_paper.zip');
    docxPath = fullfile(outDir, 'sample_paper.docx');
    zip(zipPath, {'word'}, tmpDir);
    movefile(zipPath, docxPath, 'f');
    rmdir(tmpDir, 's');
end

function figPaths = createSampleFigs(outDir)
    figPaths = cell(1, 5);
    x = linspace(0, 10, 100);

    % 图 1：超大原始窗口 + 18pt 字体
    f1 = figure('Visible', 'off', 'Position', [100, 100, 900, 650]);
    plot(x, sin(x), 'b-', 'LineWidth', 1.5); hold on;
    plot(x, cos(x), 'r--', 'LineWidth', 1.5);
    xlabel('Time (s)', 'FontSize', 18); ylabel('Amplitude', 'FontSize', 18);
    legend({'sin', 'cos'}, 'Location', 'northeast', 'FontSize', 16);
    set(gca, 'FontSize', 16, 'FontName', 'Helvetica');
    grid on;
    figPaths{1} = fullfile(outDir, 'fig1_large.fig');
    savefig(f1, figPaths{1}); close(f1);

    % 图 2：小窗口 + 7pt 字体
    f2 = figure('Visible', 'off', 'Position', [100, 100, 320, 260]);
    plot(x, exp(-0.2*x).*sin(2*x), 'k-', 'LineWidth', 1.2);
    xlabel('x (mm)', 'FontSize', 7); ylabel('Response', 'FontSize', 7);
    set(gca, 'FontSize', 7, 'FontName', 'Courier');
    grid on;
    figPaths{2} = fullfile(outDir, 'fig2_small.fig');
    savefig(f2, figPaths{2}); close(f2);

    % 图 3：柱状图 + 14pt 字体
    f3 = figure('Visible', 'off', 'Position', [100, 100, 600, 500]);
    bar([1 2 3 4], [2.4 3.8 3.1 4.5]);
    xlabel('Group', 'FontSize', 14); ylabel('Score', 'FontSize', 14);
    set(gca, 'FontSize', 13);
    figPaths{3} = fullfile(outDir, 'fig3_bar.fig');
    savefig(f3, figPaths{3}); close(f3);

    % 图 4：宽幅散点拟合图
    f4 = figure('Visible', 'off', 'Position', [100, 100, 800, 350]);
    y4 = 0.5*x + 0.4*sin(x);
    scatter(x(1:4:end), y4(1:4:end), 24, 'filled'); hold on;
    plot(x, 0.5*x, 'r-', 'LineWidth', 1.4);
    xlabel('Input Voltage (V)', 'FontSize', 15); ylabel('Current (mA)', 'FontSize', 15);
    legend({'Measured', 'Linear fit'}, 'Location', 'northwest');
    set(gca, 'FontSize', 14);
    figPaths{4} = fullfile(outDir, 'fig4_wide1.fig');
    savefig(f4, figPaths{4}); close(f4);

    % 图 5：宽幅多曲线对比图
    f5 = figure('Visible', 'off', 'Position', [100, 100, 750, 320]);
    plot(x, log1p(x), 'LineWidth', 1.4); hold on;
    plot(x, 0.8*sqrt(x), '--', 'LineWidth', 1.4);
    xlabel('Iteration', 'FontSize', 12); ylabel('Loss', 'FontSize', 12);
    legend({'Method A', 'Method B'}, 'Location', 'southeast');
    set(gca, 'FontSize', 11);
    figPaths{5} = fullfile(outDir, 'fig5_wide2.fig');
    savefig(f5, figPaths{5}); close(f5);
end
