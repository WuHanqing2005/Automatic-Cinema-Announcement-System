# Automatic-Cinema-Announcement-System
It is a system that can play film announcement automatically.

软件名称：沈北万瑞电影城-自动广播系统

版本号：2025.02.10 21:00 (正式版)

软件版权归属：吴瀚庆

未经允许，禁止盗用，侵权必究


有意请联系软件作者 吴瀚庆

微信：whq20050121

手机：19528873640

邮箱：m19528873640@outlook.com

欢迎提出宝贵意见，感谢支持！

————————————————————
更新日志：

2025.02.10 21:00

程序诞生，基于电影院自动广播系统（通过爬虫）修改而来，之前的程序为万瑞电影院（黎明店）所用

————————————————————
广播模板：

756.wav                      --  756提示音

template_cn\\1.wav           --  各位观众请注意

hall_cn\\5.wav               --  五号厅

hour_cn\\17.wav              --  十七点

minute_cn\\15.wav            --  十五分

template_cn\\2.wav           --  播放的电影

filmname_cn\\熊出没.wav       --  熊出没

template_cn\\3.wav           --  现在开始检票入场，谢谢！

————————————————————
软件使用说明：

请勿擅自修改本软件目录下的任何文件！！！

请确保在联网条件下运行本软件！！！

本软件为“电影院自动广播测试系统”，通过爬取猫眼的售票信息，获取电影排片信息，并自动生成广播语音，由吴瀚庆开发。

1.启动软件：双击main.exe文件启动软件。

2.设置参数：启动后会弹出设置窗口，让用户设置提前检票分钟数和广播循环播放次数。根据需要选择合适的分钟数和播放次数，点击“确认”按钮完成设置。

3.检查电影名文件：设置完成后，软件会自动检查filmname_cn文件夹中是否存在电影名称对应的.wav文件。若发现缺失文件，会弹出警告窗口列出缺失文件。

4.主界面操作：

播放广播：在主界面的下拉列表中选择一部电影，点击“播放广播”按钮，软件会根据设置生成并播放广播语音。

停止播放：点击“停止播放”按钮可停止当前播放的广播。

清空缓存：点击“清空缓存”按钮可清空output文件夹中的缓存文件。

清空并退出：点击“清空并退出”按钮可清空数据并退出软件。

读取并刷新：点击“读取并刷新”按钮可从data.xlsx文件中读取电影信息并更新表格。

修改电影信息：选中表格中的电影信息，点击“修改电影信息”按钮可对电影信息进行修改。

5.文件夹说明：

material文件夹：存放原始音频素材，包括提示音、电影名称语音、模板语音、厅号语音、时间语音等。

filmname_cn文件夹：存放电影名称对应的.wav文件。

output文件夹：存放生成的广播语音文件。

6.注意事项：

确保info.txt文件中的配置信息正确无误。

定期清空output文件夹中的缓存文件，以免占用过多空间。

避免过于频繁地停止播放广播，以免影响播放效果。

如有疑问或建议，请联系软件作者吴瀚庆。

7.语音包更新

更新语音包，请联系网页作者吴瀚庆，微信：whq20050121，或邮箱：m19528873640@outlook.com

将需要更新的语音包，放入对应的material文件夹中，例如电影名称语音包需要放到material文件夹下的filmname_cn文件夹中。

