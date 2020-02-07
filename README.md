### How to use

Steps:

1. 安装nodejs (https://nodejs.org/en/download/), 直接下载安装
2. 将需要遍历的文件放入 `resources` 文件夹
3. 打开termial(mac) / 命令行(windows, 好像是在搜索那里输入cmd?)
4. 进入当前工作目录，比如我放在桌面的xlsx文件夹里，我就输入 `cd ~/Desktop/xlsx` (windows有点不同有需要再解释给你)
5. 输入`npm i node-xlsx`
6. 运行 `node main.js` 
7. 生成的文件会在 `dest` 文件夹中



### 解释

目前的流程大致是，遍历 `resources` 文件夹下的所有文件 > 遍历每个文件的所有sheet，找到所有指定sheets > 遍历每个sheet的每一行，查看指定项目 (cell) 名称是否包含在内，如果是，则保存该行内容。最后输出所有保存的内容到 `dest/text.xlsx` 文件