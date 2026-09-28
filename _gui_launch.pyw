# -*- coding: utf-8 -*-
"""GUI 启动器：pythonw 无控制台，崩溃写日志；另加心跳线程供外部判活。"""
import traceback
import sys, os, threading, time

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, "_gui_err.log")
BEAT = os.path.join(HERE, "_gui_heartbeat.txt")

def _beat():
    while True:
        try:
            with open(BEAT, "a", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
        except OSError:
            pass
        time.sleep(5)

threading.Thread(target=_beat, daemon=True).start()

try:
    sys.path.insert(0, HERE)
    import gui
    gui.main()
except Exception:
    with open(LOG, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
