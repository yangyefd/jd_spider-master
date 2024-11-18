# 功能
京东爬虫，读入excel表格，定位 "商品关键词", 和 "所属商品SPUID"，通过自动化脚本依次搜索关键词，并搜索是否存在SPUID从而判断商品链接是否存在异常
![alt text](./image/result.png)
# notes
京东对爬虫有限制，所以访问较多时，需要进行扫码登录，单次任务仅需登录一次

# python打包方法
1. pip install pyinstaller 
2. 下载upx（减小打包后程序大小）
3. pyinstaller --onefile --windowed --strip  --upx-dir=/path/to/upx your_script.py
