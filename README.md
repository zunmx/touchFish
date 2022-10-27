# touchFish
居家办公的摸鱼神器
# 介绍

监控任务栏图标变化，根据是否有图标变化来判断是否有消息(可别在托盘里放一些经常变的图标，比如任务管理器)，根据进程信息，判断是否explorer未响应，并且支持自动重启explorer，定期移动鼠标，避免离开状态以及锁屏，如果有消息(任务栏图标变化)，循环播放音乐，并且支持自动调高音量，人工介入后自动恢复原来的音量，如果正在操作电脑，程序不工作，可设置超时时间，超过多久不操作自动进行检测和移动鼠标。
⚠：本工具使用了Hook，可能有杀毒软件提示危险。

屏幕左上角就是间隔时间后点击的地方，原则上是置顶的，红色说明程序空闲，绿色说明程序运行

# 需要的依赖

os
threading
psutil
win32gui
win32api
win32con
PIL
pygame
keyboard
time
mouse
ctypes
comtypes

pycaw

# 配置段

```python
self.moveTime = 45  # 检查鼠标移动的周期
        self.get_tray_timeout = 0  # 获取托盘失败次数
        self.error_max = 10  # 最大获取失败次数
        self.restart_explorer_sleep = 5  # 重启explorer等待时间
        self.rect_right_offset = 24  # 右侧偏移量，如果想只截取部分应用，需把那部分托盘移动到最左边，然后修改此偏移量
        self.container_max = 3  # 容器存放数量
        self.timer = 0  # 检测周期
        self.active_check_time = 5  # 多久不动之后才会生效
        self.trayClassName = "ToolbarWindow32"
        self.playSound = True  # 是否允许播放音乐
        self.canUpperSoundLevel = True  # 允许调整电脑音量
        self.canRestartExplorer = False  # 是否允许explorer崩溃时重启explorer
```
