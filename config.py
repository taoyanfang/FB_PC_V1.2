"""配置项：路径、WiFi、设备与账号、开关等所有可调参数"""

import os
import time

# ========== WiFi 配置（V1.2新增） ==========
WIFI_SSID = 'ASUS_F8_5G'  # 未连上WiFi时自动连接的WiFi名称

# ========== 日志 ==========
LOG_DIR = r'E:\PC-Facebook'  # 日志文件所在目录

# ========== 浏览器 ==========
USER_DATA_DIR = r'D:\user\23456789'  # 浏览器用户数据目录（原写在 browser() 内）

# ========== 路径（按脚本启动时间生成） ==========
ROOT_PATH = os.path.join(r'E:\PC-Facebook\fb_result', time.strftime("%Y%m%d", time.localtime(time.time())))
PCAP_PATH = os.path.join(r'E:\PC-Facebook\fb_pcap', time.strftime("%Y%m%d", time.localtime(time.time())))
EXCEL_NAME = os.path.join(ROOT_PATH,
                          'fb播放结果' + time.strftime("%Y%m%d_%H%M%S", time.localtime(time.time())) + '.xlsx')
CREATE_CMD = "mkdir /appslog/fb_pc_" + time.strftime("%Y%m%d", time.localtime(time.time()))

# ========== 设备与工具 ==========
WIRESHARK_PATH = r'C:\Program Files\Wireshark'  # wireshark安装目录
# WIRESHARK_PATH = r'E:\software\Wireshark'
ETH_NAME = '9DC9F098-86A3-4933-AE10-C0AE2B4D3FDA'  # 内网机PC捕包网卡编号
# ETH_NAME = '74DEFCE8-3876-4318-8F7D-AB110DE6F019'  # 外网机
IP_LIST_PATH = r'E:\PC-Facebook\fb_result\ip_list.xlsx'  # IP切换列表
INTERFACE = 'WLAN'  # PC上网网卡名称

# ========== IAS 抓包 ==========
TCPDUMP_PATH = '/apps/ddriver/tools/dumpcap'  # ddriver抓包
IAS_ETH = '0'  # ias 捕包网卡名称
HOSTNAME = '192.168.80.93'  # ias IP
USERNAME = 'root'  # ias root账号
PASSWORD = 'rzx1218'  # ias root 密码
# HOSTNAME = '192.168.70.20'
# USERNAME = 'root'
# PASSWORD = '111111'
# TCPDUMP_PATH = '/pag/sbin/tcpdump_macinmac'  # tcpdump目录+tcpdump程序名称
# IAS_ETH = 'pag0'  # ias 捕包网卡名称

# ========== Facebook 账号 ==========
FB_USERNAME = 'tyf8989@163.com'  # fb账号
FB_PASSWORD = 'qweasd123'  # fb密码

# ========== 视频播放列表 ==========
VID_FILE = r'E:\PC-Facebook\fb_result\fb视频列表_1.xlsx'
# VID_FILE = r'E:\PC-Facebook\yfb_result\fb视频列表非目标.xlsx'

# ========== 开关 ==========
IS_SCREEN = False  # 是否有界面执行，True时，无界面播放
IS_IAS_PCAP = True  # 是否开启ias捕包 False不开启，True开启
IS_TIME_LIMIT = False  # 是否限制ip封堵使用时长，True限制，False不限制
IS_IP_SWITCH = True  # True每次都切换IP，False封住再切换
