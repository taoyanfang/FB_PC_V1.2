"""基础方法：日志重定向、WiFi管理、进程清理、IP切换、通用小工具"""

import os
import sys
import time
import socket
import subprocess

import ping3
import psutil
import openpyxl
from datetime import datetime

from config import WIFI_SSID, LOG_DIR

# 日志重定向：同时输出到控制台和文件
log_file_path = os.path.join(LOG_DIR, f'fb_log_{time.strftime("%Y%m%d_%H%M%S")}.log')
class TeeLogger:
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, 'a', encoding='utf-8')
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()
    def flush(self):
        self.terminal.flush()
        self.log.flush()
sys.stdout = TeeLogger(log_file_path)
print(f'[INFO] 日志文件: {log_file_path}')
print('[INFO] 脚本版本: V1.2（新增WiFi自动连接/强制切换到指定WiFi功能）')

# ========== V1.2新增：WiFi自动重连功能 ==========
def get_wifi_ssid(interface='WLAN'):
    """获取WLAN当前连接的WiFi名称（SSID）。
    未连接WiFi时，netsh输出中没有非空SSID行，返回None"""
    try:
        output = subprocess.check_output('netsh wlan show interfaces', shell=True)
        text = output.decode('gbk', errors='ignore')
    except Exception as e:
        print(f'[ERROR] 获取WLAN状态失败: {e}')
        return None

    ssid = None
    for line in text.split('\n'):
        line = line.strip()
        # 匹配SSID行（排除BSSID行），格式: SSID                   : ASUS_F8_5G
        if line.upper().startswith('SSID') and not line.upper().startswith('BSSID'):
            parts = line.split(':', 1)
            if len(parts) == 2 and parts[1].strip():
                ssid = parts[1].strip()
    return ssid


def wait_for_valid_ip(interface='WLAN', timeout=60):
    """等待网卡通过DHCP获取有效IP（排除APIPA地址169.254.x.x）"""
    for i in range(timeout // 5):
        try:
            addrs = psutil.net_if_addrs()
            if interface in addrs and len(addrs[interface]) > 1:
                ip = addrs[interface][1].address
                if ip and not str(ip).startswith('169.254.'):
                    print(f'[OK] 网卡 {interface} 获取到有效IP: {ip}')
                    return ip
        except Exception as e:
            print(f'[WARN] 获取网卡IP异常: {e}')
        print(f'[INFO] 等待网卡 {interface} 获取有效IP... ({(i + 1) * 5}秒)')
        time.sleep(5)
    print(f'[WARN] 等待网卡 {interface} 获取有效IP超时({timeout}秒)')
    return None


def ensure_wifi_connected(ssid=WIFI_SSID, interface='WLAN', wait_seconds=60):
    """检查WLAN连接状态：
    1. 已连接指定WiFi -> 直接返回True
    2. 连接的是其他WiFi -> 强制断开并切换到指定WiFi
    3. 未连接任何WiFi -> 直接连接指定WiFi
    返回True表示WLAN已连接到指定WiFi"""
    curr_ssid = get_wifi_ssid(interface)
    if curr_ssid == ssid:
        print(f'[OK] WLAN已连接指定WiFi: {curr_ssid}')
        return True

    if curr_ssid:
        print(f'[WARN] WLAN当前连接的是其他WiFi: {curr_ssid}，强制切换到: {ssid}')
        try:
            subprocess.run(f'netsh wlan disconnect interface="{interface}"', shell=True,
                           capture_output=True, timeout=30)
            time.sleep(3)
        except Exception as e:
            print(f'[WARN] 断开当前WiFi失败: {e}')
    else:
        print(f'[WARN] WLAN未连接WiFi，尝试自动连接: {ssid}')
    try:
        # 确保无线网卡已启用
        subprocess.run(f'netsh interface set interface "{interface}" enabled', shell=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
        time.sleep(2)

        # 检查WiFi配置文件是否存在（需手动连接过一次才会保存）
        try:
            profiles = subprocess.check_output('netsh wlan show profiles', shell=True).decode('gbk', errors='ignore')
            if ssid not in profiles:
                print(f'[ERROR] 系统中不存在WiFi配置文件: {ssid}')
                print('[INFO] 请先手动连接一次该WiFi并勾选自动连接，之后脚本才能自动重连')
        except Exception:
            pass

        # 循环尝试连接直到成功或超时（每轮10秒，共wait_seconds秒）
        for i in range(wait_seconds // 10):
            result = subprocess.run(f'netsh wlan connect name="{ssid}" interface="{interface}"',
                                    shell=True, capture_output=True, timeout=30)
            print(f'[INFO] 连接命令输出: {(result.stdout or b"").decode("gbk", errors="ignore").strip()}')

            for j in range(5):
                time.sleep(2)
                curr_ssid = get_wifi_ssid(interface)
                if curr_ssid:
                    print(f'[OK] WiFi连接成功: {curr_ssid} (耗时约{i * 10 + (j + 1) * 2}秒)')
                    # WiFi连接后等待DHCP分配有效IP
                    wait_for_valid_ip(interface=interface, timeout=60)
                    return True
            print(f'[INFO] 等待WiFi连接中... ({(i + 1) * 10}秒)')
    except Exception as e:
        print(f'[ERROR] 执行WiFi连接操作失败: {e}')

    print(f'[ERROR] 连接WiFi {ssid} 失败或超时({wait_seconds}秒)')
    return False

# 关闭进程的函数
def kill_processes(exe_name_list):
    for name in exe_name_list:
        # 在Windows上可能需要获取完整的路径
        if os.name == 'nt':
            name = name + '.exe'
        for process in psutil.process_iter(['name']):
            # 检查进程名称
            if process.info['name'] == name:
                try:
                    process.kill()  # 终止进程
                    print(f'{name} process with PID {process.pid} has been terminated.')
                except Exception as e:
                    print(f"Failed to kill {name} process with PID {process.pid}: {e}")


# 修改IP
def set_ip_address(ip_list_path, interface, ip_filter, is_time_limit=False):
    workbook = openpyxl.load_workbook(filename=ip_list_path)
    sheet = workbook.active
    edit_sheet = workbook.active
    num = 1
    max_row = sheet.max_row
    for row in sheet.iter_rows(values_only=True, min_col=1, max_col=4, min_row=2):
        num = num + 1
        edit_ip_cell = 'C' + str(num)
        edit_status_cell = 'D' + str(num)
        if row[3]:
            print('row[3]', row[3], max_row, num)
            if max_row - num <= 2 or len(str(row[0])) < 7:
                for i in range(2, max_row + 1):
                    edit_status_cell = 'D' + str(i)
                    edit_sheet[edit_status_cell] = None
                workbook.save(ip_list_path)
            continue
        if row[0] == ip_filter:
            curr_tiem = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
            edit_sheet[edit_ip_cell] = curr_tiem
            edit_sheet[edit_status_cell] = True
            workbook.save(ip_list_path)
            continue
        try:
            if row[2]:
                try:
                    old_timestamp = int(datetime.strptime(row[2], "%Y-%m-%d %H:%M:%S").timestamp())
                    time_cal = int(time.time()) - old_timestamp

                except Exception as e:
                    time_cal = 0
                if time_cal > 1800 and is_time_limit:
                    print('time_cal')
                    ip_status = ping3.ping(row[0], timeout=10)
                    if not ip_status and socket.gethostbyname(socket.getfqdn(socket.gethostname())) != row[0]:
                        subprocess.run(
                            ["netsh", "interface", "ip", "set", "address", interface, "static", row[0], "255.255.255.0",
                             row[1]])
                        edit_sheet[edit_ip_cell] = ''
                        edit_sheet[edit_status_cell] = True
                        workbook.save(ip_list_path)
                        return num
            else:
                ip_status = ping3.ping(row[0], timeout=10)
                print(ip_status, 'ip_status')
                if not ip_status and socket.gethostbyname(socket.getfqdn(socket.gethostname())) != row[0]:
                    subprocess.run(
                        ["netsh", "interface", "ip", "set", "address", interface, "static", row[0], "255.255.255.0",
                         row[1]])
                    edit_sheet[edit_status_cell] = True
                    workbook.save(ip_list_path)
                    return num
        except Exception as e:
            return num


def time_to_seconds(time_str):
    # 将时间字符串拆分为小时、分钟和秒数
    time_length = len(time_str.split(':'))
    if time_length == 2:
        minutes, seconds = time_str.split(':')
        total_seconds = int(minutes) * 60 + int(seconds)
        return total_seconds
    elif time_length == 3:
        hours, minutes, seconds = time_str.split(':')
        total_seconds = int(hours) * 3600 + int(minutes) * 60 + int(seconds)
        return total_seconds


# 获取浏览器版本
def get_browser_version():
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True
        )
        print(output.decode().strip().split())
        browser_version = int(output.decode().strip().split()[-1].split('.')[0])
        return browser_version
    except subprocess.CalledProcessError as e:
        print(f"Error getting Chrome version: {e}")
        return 128
