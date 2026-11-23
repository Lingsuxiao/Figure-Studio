%START_APP 一键启动「科研绘图排版与字体统一工具」
%   在 MATLAB 中运行本脚本，或在命令行直接输入 start_app 即可打开界面。

thisDir = fileparts(mfilename('fullpath'));
addpath(thisDir);
addpath(fullfile(thisDir, 'core'));

app = FigureLayoutApp();
