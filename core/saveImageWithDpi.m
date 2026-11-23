function saveImageWithDpi(imgRgb, outPath, dpi)
%SAVEIMAGEWITHDPI 保存高清 PNG 或 TIFF 图像，并写入精确物理 DPI 元数据
%   这样直接将图片拖入 Microsoft Word 时，Word 会自动按 100% 原始厘米宽度显示，
%   不会产生缩放，确保图内字体与 Word 正文字号 1:1 完全一致。

    arguments
        imgRgb (:,:,3) uint8
        outPath (1,:) char
        dpi (1,1) double {mustBePositive} = 600
    end

    [parentDir, ~, ext] = fileparts(outPath);
    if ~isempty(parentDir) && ~isfolder(parentDir)
        mkdir(parentDir);
    end

    switch lower(ext)
        case '.png'
            ppm = round(dpi / 0.0254); % pixels per meter
            imwrite(imgRgb, outPath, ...
                'ResolutionUnit', 'meter', ...
                'XResolution', ppm, ...
                'YResolution', ppm);
        case {'.tif', '.tiff'}
            imwrite(imgRgb, outPath, ...
                'Compression', 'lzw', ...
                'Resolution', [dpi, dpi]);
        otherwise
            imwrite(imgRgb, outPath);
    end
end
