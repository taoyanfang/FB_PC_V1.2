"""业务逻辑：FB 播放任务类（自 FB_PC_V1.2.py 原样迁入，仅 user_data_dir 改为引用配置）"""

import os
import time
import socket
import subprocess

import ping3
import psutil
import openpyxl
import paramiko
import xlsxwriter
import multiprocessing
from multiprocessing.dummy import Process

from selenium.webdriver import Chrome, ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
import undetected_chromedriver as uc

from config import WIFI_SSID, USER_DATA_DIR
from utils import (
    get_wifi_ssid,
    wait_for_valid_ip,
    ensure_wifi_connected,
    kill_processes,
    set_ip_address,
    time_to_seconds,
    get_browser_version,
)

class FB(multiprocessing.Process):
    def __init__(self, root_path, pcap_path, excel_name, wireshark_path, eth_name, ip_list_path, interface,
                 tcpdump_path, ias_eth, hostname, username, password, is_screen, is_ias_pcap, vid_file, fb_username,
                 fb_password, is_time_limit, is_ip_switch):
        super().__init__()
        self.root_path = root_path
        self.pcap_path = pcap_path
        self.excel_name = excel_name
        self.wireshark_path = wireshark_path  # wireshark安装目录
        self.eth_name = eth_name  # PC捕包网卡编号
        self.ip_list_path = ip_list_path  # IP切换列表
        self.interface = interface  # PC上网网卡名称
        self.tcpdump_path = tcpdump_path  # tcpdump目录+tcpdump程序名称
        self.ias_eth = ias_eth  # ias 捕包网卡名称
        self.hostname = hostname  # ias IP
        self.username = username  # ias root账号
        self.password = password  # ias root 密码
        self.ip = None
        self.driver = None
        self.protocol_dict = {}
        self.tcpdump_pid = None
        self.cmd = None
        self.driver = None
        self.is_screen = is_screen
        self.browser_version = None  # 浏览器版本
        self.is_ias_pcap = is_ias_pcap  # 是否需要ias pcap
        self.vid_file = vid_file
        self.cookies = None
        self.fb_username = fb_username
        self.fb_password = fb_password
        self.uc = False
        self.is_time_limit = is_time_limit
        self.is_ip_switch = is_ip_switch

    def browser(self, sslkey_log_path=None):
        dri = Options()
        dri.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
        dri.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
        dri.add_argument('--hide-scrollbars')  # 隐藏滚动条，应对一些特殊页面
        dri.add_argument('--ignore-certificate-errors')  # 忽略不安全提示页面
        dri.add_argument("--window-size=1920,1080")
        dri.add_argument("--host-resolver-rules=MAP optimizationguide-pa.googleapis* 100.100.100.100")

        # SSL/TLS密钥日志相关参数（确保SSLKEYLOGFILE能够正常工作）
        if sslkey_log_path:
            # 直接指定SSL密钥日志文件的完整路径（这是关键！）
            dri.add_argument(f'--ssl-key-log-file={sslkey_log_path}')
            dri.add_argument('--disable-features=VizDisplayCompositor')  # 某些情况下需要禁用
            dri.add_argument('--disable-software-rasterizer')  # 确保SSL功能正常
            print(f'[INFO] SSL密钥日志文件路径已设置: {sslkey_log_path}')
        else:
            print('[WARN] 未传入SSL密钥日志文件路径，SSL解密将无法工作！')

        user_data_dir = USER_DATA_DIR
        # dri.add_argument("--disable-software-rasterizer")
        # dri.add_argument("--disable-features=OptimizationGuide")
        # dri.add_argument("--disable-features=PreloadMediaEngagementData,PreloadSRPAndMedia")
        # dri.add_argument("--disable-background-networking")
        # dri.add_argument("--disable-sync")
        # dri.add_argument("--disable-background-timer-throttling")
        # dri.add_argument("--disable-blink-features=AutomationControlled")
        # dri.add_argument("--disable-features=OptimizationGuide")

        # opt.binary_location = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" # 手动指定使用的浏览器位置

        if self.is_screen:
            dri.add_argument('--headless')  # 浏览器不提供可视化界面。Linux下如果系统不支持可视化不加这条会启动失败
            if self.uc:
                try:
                    self.driver = uc.Chrome(options=dri, version_main=get_browser_version(), use_subprocess=True)
                except Exception as e:
                    print(f"使用 undetected_chromedriver 失败: {e}")
                    print("尝试使用 Selenium Chrome...")
                    self.driver = Chrome(options=dri)
            else:
                self.driver = Chrome(options=dri)
        else:
            if self.uc:
                try:
                    self.driver = uc.Chrome(options=dri, version_main=get_browser_version(), use_subprocess=True)
                except Exception as e:
                    print(f"使用 undetected_chromedriver 失败: {e}")
                    print("尝试使用 Selenium Chrome...")
                    self.driver = Chrome(options=dri)
            else:
                self.driver = Chrome(options=dri)
            self.driver.maximize_window()
        return self.driver

    # 获取播放列表
    def get_data_dict(self):
        workbook = openpyxl.load_workbook(filename=self.vid_file)
        sheet = workbook.active
        num = 0
        for row in sheet.iter_rows(values_only=True, min_col=2, max_col=5, min_row=2):
            num = num + 1
            # 检查必要字段是否为空
            if not row[0]:  # vid不能为空
                print(f'[WARN] 第{num+1}行：视频ID为空，跳过')
                continue
            if not row[2]:  # play_time_limit不能为空
                print(f'[WARN] 第{num+1}行：播放时长为空，使用默认值60秒')
                row = (row[0], row[1], 60, row[3])  # 设置默认值60秒
            self.protocol_dict[num] = [row[0], str(row[1]).upper().split('P')[0] if row[1] else 'AUTO', row[2], row[3]]
        return self.protocol_dict

    # 获取网卡IP
    def get_ip(self, interface):
        """获取网卡IP并进行合理性判断"""
        # V1.2新增：检查WiFi连接状态，未连接或连了其他WiFi时强制切换到指定WiFi
        curr_ssid = get_wifi_ssid(interface)
        if curr_ssid != WIFI_SSID:
            if ensure_wifi_connected(ssid=WIFI_SSID, interface=interface):
                valid_ip = wait_for_valid_ip(interface=interface, timeout=60)
                if valid_ip:
                    self.ip = valid_ip
                    print(f'[OK] WiFi重连成功，当前IP: {self.ip}')
                    return self.ip
                print('[ERROR] WiFi已连接但未获取到有效IP，继续按原逻辑处理')
            else:
                print('[ERROR] WiFi自动连接失败，继续按原逻辑处理')

        ip_list = psutil.net_if_addrs()
        self.ip = ip_list[interface][1].address

        # IP合理性判断
        if self.is_ip_invalid(self.ip):
            print(f'[WARN] 当前WLAN IP不合理: {self.ip}')
            print('[INFO] 尝试切换到IP列表中下一个可用IP...')

            if self.is_ip_switch:
                try:
                    # 调用set_ip_address切换到下一个可用IP
                    next_ip_row = set_ip_address(
                        ip_list_path=self.ip_list_path,
                        interface=self.interface,
                        ip_filter=self.ip,  # 当前不合理的IP作为过滤
                        is_time_limit=self.is_time_limit
                    )

                    # 如果找到下一个可用IP，重新获取IP
                    if next_ip_row:
                        print(f'[INFO] 找到下一个可用IP，等待网络重新配置...')
                        time.sleep(10)  # 等待网络配置生效

                        # 重新获取IP
                        ip_list = psutil.net_if_addrs()
                        self.ip = ip_list[interface][1].address

                        # 再次检查合理性
                        if self.is_ip_invalid(self.ip):
                            print(f'[ERROR] 新IP仍然不合理: {self.ip}')
                            print('[ERROR] IP设置不成功，无法继续')
                            return None
                        else:
                            print(f'[OK] IP切换成功，新IP: {self.ip}')
                    else:
                        print('[ERROR] IP列表中没有其他可用IP')
                        return None
                except Exception as e:
                    print(f'[ERROR] IP切换失败: {e}')
                    return None
            else:
                print('[ERROR] IP切换功能未启用，无法解决IP不合理问题')
                return None
        else:
            print(f'[OK] 当前WLAN IP合理: {self.ip}')

        return self.ip

    def is_ip_invalid(self, ip):
        """判断IP是否不合理"""
        if not ip:
            return True

        # 检查是否是APIPA地址（169.254.0.0/16）
        if ip.startswith('169.254.'):
            print(f'[REASON] 检测到APIPA地址（DHCP分配失败）: {ip}')
            return True
        return False

    # PC捕包
    def get_pcap(self, pcap_name, *args):
        # 捕包目录
        if not os.path.exists(self.pcap_path):
            os.mkdir(self.pcap_path)
        pcap_full_name = os.path.join(self.pcap_path, pcap_name)
        cmd = self.wireshark_path + r'\tshark.exe -i \Device\NPF_{%s}  -w %s' % (self.eth_name, pcap_full_name)
        subprocess.run(cmd)

    # IAS捕包
    def ias_opt(self, cmd, *args):
        if not self.is_ias_pcap:
            return
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        print(self.get_ip(self.interface))
        time.sleep(10)
        try:
            ip_status = ping3.ping(self.hostname, timeout=10)
            print('IP是否可用', ip_status)

            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            # sock.bind((self.get_ip(self.interface), 0))
            sock.bind(('10.10.163.86', 0))
            ##sock.bind(('10.10.188.185', 0))
            sock.settimeout(30)
            sock.connect((self.hostname, 22))
            transport = paramiko.Transport(sock)
            transport.connect(username=self.username, password=self.password)
            client._transport = transport
            filename = '/appslog/fb_pc_' + time.strftime("%Y%m%d", time.localtime(time.time()))
            file_dir = '[ -d "{0}" ] && echo "Exists" || mkdir {0}'.format(filename)
            client.exec_command(file_dir)

            # ========== 在启动抓包之前，先清理残留的dumpcap进程 ==========
            if 'tcpdump' in cmd or 'dumpcap' in cmd:
                print('[INFO] 检查并清理IAS上残留的dumpcap进程...')

                try:
                    # 查找所有dumpcap进程
                    print('[INFO] 执行: ps aux | grep dumpcap')
                    _, stdout, _ = client.exec_command('ps aux | grep dumpcap')
                    ps_output = stdout.read().decode().strip()
                    print(f'[INFO] ps aux输出:\n{ps_output}')

                    # 解析PID并kill
                    if ps_output and 'dumpcap' in ps_output:
                        lines = ps_output.split('\n')
                        killed_pids = []

                        for line in lines:
                            if 'dumpcap' in line and 'grep' not in line:
                                parts = line.split()
                                if len(parts) >= 2:
                                    pid = parts[1]  # PID通常是第二列
                                    if pid.isdigit():
                                        try:
                                            # 执行kill -9
                                            kill_cmd = f'kill -9 {pid}'
                                            print(f'[INFO] 执行: {kill_cmd}')
                                            _, stdout, _ = client.exec_command(kill_cmd)
                                            stdout.channel.recv_exit_status()  # 等待命令完成
                                            killed_pids.append(pid)
                                            print(f'[OK] 已终止进程: PID {pid}')
                                        except Exception as kill_error:
                                            print(f'[ERROR] 终止进程失败 PID {pid}: {kill_error}')

                        if killed_pids:
                            print(f'[OK] 成功终止 {len(killed_pids)} 个dumpcap进程: {killed_pids}')
                            time.sleep(2)  # 等待进程完全终止
                        else:
                            print('[INFO] 没有找到需要终止的dumpcap进程')
                    else:
                        print('[INFO] 没有发现dumpcap进程在运行')

                except Exception as ps_error:
                    print(f'[ERROR] 清理dumpcap进程失败: {ps_error}')

            print('[INFO] 开始执行抓包命令...')
            client.exec_command(cmd)
            #if 'tcpdump' in cmd:
            #    time.sleep(2)
            #    stdin, stdout, stderr = client.exec_command('cat /appslog/runing_pid')
            #    self.tcpdump_pid = stdout.read().decode()
            if 'tcpdump' in cmd:
                time.sleep(2)  # 增加等待时间
                stdin, stdout, stderr = client.exec_command('cat /appslog/runing_pid')
                exit_status = stdout.channel.recv_exit_status()  # 等待命令完成
                pid_content = stdout.read().decode().strip()  # 读取并去除空白
                print(f'读取PID文件内容: [{pid_content}]')  # 调试信息
                if pid_content:
                    self.tcpdump_pid = pid_content
                    print(f'成功获取PID: {self.tcpdump_pid}')
                else:
                    print('PID文件为空')
                    self.tcpdump_pid = None
            elif 'kill' in cmd:
                self.tcpdump_pid = None
            client.close()

        except Exception as e:
            print('IAS操作失败{0}'.format(e))

    def login(self, *args):
        self.uc = True
        print('123')
        if self.is_screen:
            self.is_screen = False
            print('ttttt')
            self.driver = self.browser(sslkey_log_path=None)
            self.is_screen = True
            print('uuuuuu')
        else:
            print('11222222222222')
            self.driver = self.browser(sslkey_log_path=None)
            print('32222231213')
        print('2311222222222')
        # self.driver.get('https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fwww.facebook.com%2Fsignin%3Faction_handle_signin%3Dtrue%26app%3Ddesktop%26hl%3Dzh-CN%26next%3D%252F&hl=zh-CN&passive=false&service=facebook&uilel=0&flowName=GlifWebSignIn&flowEntry=AddSession&dsh=S-1543343706%3A1718094935469557&ddm=0')
        self.driver.get('http://www.facebook.com/')
        # time.sleep(5)
        # self.driver.find_element(by=By.XPATH,value='//*[@id="email"]').click()
        # time.sleep(1)
        self.driver.find_element(by=By.XPATH, value='//*[@id="email"]').send_keys(self.fb_username)
        # self.driver.find_element(by=By.XPATH,value='//*[@id="identifierNext"]/div/button/span').click()
        time.sleep(10)
        self.driver.find_element(by=By.XPATH, value='//*[@id="pass"]').click()
        time.sleep(1)
        self.driver.find_element(by=By.XPATH, value='//*[@id="pass"]').send_keys(self.fb_password)
        self.driver.find_element(by=By.XPATH, value='//*[@id="loginbutton"]').click()
        time.sleep(30)
        # 保存Cookies
        self.cookies = self.driver.get_cookies()
        # print(self.cookies)
        # self.driver.quit()
        self.uc = False
        # kill_processes(['chrome'])

    def _cleanup_ias_capture(self):
        """清理 IAS 上所有残留的抓包进程"""
        try:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.bind(('10.10.163.86', 0))
            sock.settimeout(30)
            sock.connect((self.hostname, 22))
            transport = paramiko.Transport(sock)
            transport.connect(username=self.username, password=self.password)
            client._transport = transport

            # 查找所有 dumpcap 进程
            print('[INFO] 执行: ps aux | grep dumpcap')
            stdin, stdout, stderr = client.exec_command('ps aux | grep dumpcap')
            ps_output = stdout.read().decode().strip()
            print(f'[INFO] ps aux输出:\n{ps_output}')

            # 解析 PID 并 kill
            if ps_output and 'dumpcap' in ps_output:
                lines = ps_output.split('\n')
                killed_pids = []

                for line in lines:
                    if 'dumpcap' in line and 'grep' not in line:
                        parts = line.split()
                        if len(parts) >= 2:
                            pid = parts[1]
                            if pid.isdigit():
                                try:
                                    kill_cmd = f'kill -15 {pid}'
                                    print(f'[INFO] 执行: {kill_cmd}')
                                    client.exec_command(kill_cmd)
                                    killed_pids.append(pid)
                                    print(f'[OK] 已终止进程: PID {pid}')
                                except Exception as kill_error:
                                    print(f'[ERROR] 终止进程失败 PID {pid}: {kill_error}')

                if killed_pids:
                    print(f'[OK] 成功终止 {len(killed_pids)} 个 dumpcap 进程: {killed_pids}')
                else:
                    print('[INFO] 没有找到需要终止的 dumpcap 进程')
            else:
                print('[INFO] 没有发现 dumpcap 进程在运行')

            # 清空 PID 文件
            client.exec_command('echo "" > /appslog/runing_pid')

            client.close()
        except Exception as e:
            print(f'[ERROR] 清理 IAS 抓包进程失败: {e}')

    def open_fb(self, vid, play_time_limit, resolution, *args):
        # ========== 第一阶段：初始化和准备工作 ==========
        resolution_dict = {
            "1440": "1", "1080": "2", "720": "3", "640": "4", "540": "5", "480": "6", "360": "7", "270": "8",
        }
        url = 'https://www.facebook.com/watch?v=' + vid

        ip = self.get_ip(self.interface)

        # 检查IP是否获取成功
        if not ip:
            print('[ERROR] 无法获取有效的WLAN IP地址')
            print('[ERROR] IP设置不成功，无法继续播放视频')

            start_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))
            end_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))

            # 返回错误结果
            return '无', '无', '无', 'IP设置不成功', '无', start_time, end_time, '无', '无', '无'

        pcap_name = time.strftime("%Y%m%d", time.localtime(
            time.time())) + '_win10_chrome' + '_' + vid + '_' + resolution.lower() + '_' + ip + '_pc.pcap'

        if self.is_ias_pcap:
            pcap_ias_name = time.strftime("%Y%m%d", time.localtime(
                time.time())) + '_win10_chrome' + '_' + vid + '_' + resolution.lower() + '_' + ip + '_ias.pcap'
        else:
            pcap_ias_name = '未开启IAS捕包'

        sslkey_name = time.strftime("%Y%m%d", time.localtime(
            time.time())) + '_win10_chrome' + '_' + vid + '_' + resolution.lower() + '_' + ip + '_pc.key'

        self.cmd = 'nohup {0} -i {1} -f "host {2}" -w /appslog/fb_pc_{3}/{4} &> /dev/null & echo $! > /appslog/runing_pid'.format(
            self.tcpdump_path, self.ias_eth, ip, time.strftime("%Y%m%d", time.localtime(time.time())), pcap_ias_name)

        # 设置SSL密钥日志文件环境变量（必须在浏览器启动前设置）
        sslkey_log_path = os.path.join(self.pcap_path, sslkey_name)
        os.environ['SSLKEYLOGFILE'] = sslkey_log_path

        # 验证密钥日志文件路径是否可写
        try:
            # 确保目录存在
            os.makedirs(self.pcap_path, exist_ok=True)
            # 测试文件创建权限
            test_file = os.path.join(self.pcap_path, 'test_write.tmp')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            print(f'[INFO] SSL密钥日志文件路径: {sslkey_log_path}')
            print('[INFO] 密钥日志文件目录可写')
        except Exception as e:
            print(f'[ERROR] 密钥日志文件路径有问题: {e}')
            print('[WARN] SSL密钥解密可能无法工作')

        # ========== 第一阶段：清理所有残留进程（严格顺序） ==========
        print('[INFO] ========== 开始清理所有残留进程 ==========')
        try:
            # 1.1 先清理 IAS 上残留的抓包进程
            if self.is_ias_pcap:
                print('[INFO] 清理 IAS 上残留的抓包进程...')
                self._cleanup_ias_capture()
                time.sleep(2)
        except Exception as e:
            print(f'[ERROR] 清理 IAS 抓包进程失败: {e}')

        try:
            # 1.2 清理本机残留的抓包进程
            print('[INFO] 清理本机残留的抓包进程...')
            kill_processes(exe_name_list=['chrome', 'tshark'])
            time.sleep(2)
        except Exception as e:
            print(f'[ERROR] 清理本机抓包进程失败: {e}')
        print('[OK] 所有残留进程清理完成')

        # ========== 第二阶段：启动 IAS 抓包 ==========
        print('[INFO] ========== 启动 IAS 抓包 ==========')
        if self.is_ias_pcap:
            try:
                p0 = Process(target=self.ias_opt, args=(self.cmd, self.hostname, self.username, self.password))
                p0.start()
                print('[INFO] IAS 抓包进程已启动')

                # 2.1 等待 PID 获取成功（确保抓包已启动）
                pid_time = 0
                while not self.tcpdump_pid:
                    time.sleep(1)
                    try:
                        client = paramiko.SSHClient()
                        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                        sock.bind(('10.10.163.86', 0))
                        sock.settimeout(5)
                        sock.connect((self.hostname, 22))
                        transport = paramiko.Transport(sock)
                        transport.connect(username=self.username, password=self.password)
                        client._transport = transport
                        stdin, stdout, stderr = client.exec_command('cat /appslog/runing_pid')
                        stdout.channel.recv_exit_status()
                        pid_content = stdout.read().decode().strip()
                        client.close()
                        if pid_content and pid_content.isdigit():
                            self.tcpdump_pid = pid_content
                            print(f'[OK] 成功获取 IAS 抓包 PID: {self.tcpdump_pid}')
                    except Exception as e:
                        print(f'[WARN] 读取 PID 失败: {e}')

                    pid_time = pid_time + 1
                    if pid_time > 30:
                        print(f'[ERROR] 等待 IAS 抓包 PID 超时 ({pid_time}秒)')
                        break
                    print(f'[INFO] 等待 IAS 抓包 PID: {pid_time}秒')

                if not self.tcpdump_pid:
                    print('[ERROR] IAS 抓包启动失败，无法继续')
                    start_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))
                    end_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))
                    return '无', '无', '无', 'IAS抓包启动失败', '无', start_time, end_time, '无', '无', '无'

                # 2.2 等待 IAS 抓包完全启动（额外等待）
                print('[INFO] 等待 IAS 抓包完全启动...')
                time.sleep(3)
            except Exception as e:
                print(f'[ERROR] 启动 IAS 抓包失败: {e}')
        else:
            print('[INFO] IAS 抓包未启用，跳过')

        # ========== 第三阶段：启动 PC 抓包 ==========
        print('[INFO] ========== 启动 PC 抓包 ==========')
        try:
            p1 = Process(target=self.get_pcap, args=(pcap_name, self.pcap_path, self.wireshark_path, self.eth_name))
            p1.start()
            print('[OK] PC 抓包进程已启动')

            # 3.1 等待 PC 抓包完全启动
            print('[INFO] 等待 PC 抓包完全启动...')
            time.sleep(3)
        except Exception as e:
            print(f'[ERROR] 启动 PC 抓包失败: {e}')

        print('[OK] ========== 所有抓包已启动，准备打开浏览器 ==========')
        start_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))
        try:
            self.driver = self.browser(sslkey_log_path=sslkey_log_path)
            time.sleep(5)
            self.driver.get(url)
            time.sleep(10)

            # ========== 检查URL是否访问成功 ==========
            url_access_success = False
            page_title = ""
            page_source = ""

            try:
                # 获取页面标题
                page_title = self.driver.title
                print(f'[INFO] 页面标题: {page_title}')

                # 获取页面源代码（前5000字符）
                page_source = self.driver.page_source[:5000]
                print(f'[INFO] 页面源代码长度: {len(self.driver.page_source)} 字符')

            except Exception as get_page_error:
                print(f'[ERROR] 获取页面信息失败: {get_page_error}')

            # 检查是否有"无法访问此网站"的错误提示
            error_indicators = [
                '无法访问此网站',
                '无法连接',
                'ERR_CONNECTION_REFUSED',
                'ERR_NAME_NOT_RESOLVED',
                'ERR_CONNECTION_TIMED_OUT',
                "This site can't be reached",
                "This page isn't available",
                'Site cannot be reached',
                'Page cannot be found',
                '无法找到服务器',
                '服务器未找到',
                '连接被拒绝'
            ]

            page_has_error = False
            found_error_indicator = None

            # 检查页面标题
            for indicator in error_indicators:
                if indicator.lower() in page_title.lower():
                    page_has_error = True
                    found_error_indicator = indicator
                    print(f'[ERROR] 页面标题包含错误提示: {indicator}')
                    break

            # 如果标题中没有错误，检查页面内容
            if not page_has_error and page_source:
                for indicator in error_indicators:
                    if indicator.lower() in page_source.lower():
                        page_has_error = True
                        found_error_indicator = indicator
                        print(f'[ERROR] 页面内容包含错误提示: {indicator}')
                        break

            if page_has_error:
                print(f'[ERROR] 检测到"无法访问此网站"错误: {found_error_indicator}')
                print(f'[INFO] URL: {url}')
                print(f'[INFO] 页面标题: {page_title}')

                # 结束该vid的播放
                print('[INFO] 结束该vid的播放并返回失败结果')

                # 先结束IAS抓包进程
                if self.is_ias_pcap and self.tcpdump_pid:
                    try:
                        kill_cmd = 'kill -15 {0}'.format(self.tcpdump_pid)
                        print(f'[INFO] 结束IAS抓包进程，PID: {self.tcpdump_pid}')
                        self.ias_opt(kill_cmd, self.hostname, self.username, self.password)
                        self.tcpdump_pid = None
                    except Exception as kill_error:
                        print(f'[ERROR] 结束抓包进程失败: {kill_error}')

                # 清理本地资源
                try:
                    self.driver.quit()
                    time.sleep(1)
                    kill_processes(exe_name_list=['chrome', 'tshark'])
                except Exception as cleanup_error:
                    print(f'[ERROR] 清理本地资源失败: {cleanup_error}')

                # V1.2新增：网络访问异常时，先检查并尝试重连WiFi
                try:
                    print('[INFO] 检测到网络访问异常，先检查并尝试重连WiFi...')
                    ensure_wifi_connected(ssid=WIFI_SSID, interface=self.interface)
                except Exception as wifi_error:
                    print(f'[ERROR] WiFi重连检查失败: {wifi_error}')

                # IP切换
                if self.is_ip_switch:
                    try:
                        set_ip_address(ip_list_path=self.ip_list_path, interface=self.interface, ip_filter=ip,
                                       is_time_limit=self.is_time_limit)
                    except Exception as ip_error:
                        print(f'[ERROR] IP切换失败: {ip_error}')

                end_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))

                # 返回失败结果
                print('[FAIL] 视频链接打开失败 - 无法访问此网站')
                return '无', '无', '无', '视频链接打开失败', '无', start_time, end_time, ip, '无', '无'
            else:
                print('[OK] URL访问成功，未检测到错误提示')
        except WebDriverException as e:
            # 先结束IAS抓包进程
            if self.is_ias_pcap and self.tcpdump_pid:
                try:
                    kill_cmd = 'kill -15 {0}'.format(self.tcpdump_pid)
                    print(f'[INFO] 异常情况下结束IAS抓包进程，PID: {self.tcpdump_pid}')
                    self.ias_opt(kill_cmd, self.hostname, self.username, self.password)
                    self.tcpdump_pid = None
                except Exception as kill_error:
                    print(f'[ERROR] 异常情况下结束抓包进程失败: {kill_error}')

            # 清理本地资源
            self.driver.quit()
            kill_processes(exe_name_list=['chrome', 'tshark'])

            # 检查是否是连接超时错误
            error_msg = str(e)
            if 'ERR_CONNECTION_TIMED_OUT' in error_msg or 'net::ERR_CONNECTION_TIMED_OUT' in error_msg:
                print(f'❌ 视频链接打开失败: {url}')
                print(f'错误信息: {error_msg}')
                print('🔧 请检查测试网络是否正常！')
            else:
                print(f'❌ 视频链接打开失败: {url}')
                print(f'错误信息: {error_msg}')

            # V1.2新增：网络访问异常时，先检查并尝试重连WiFi
            try:
                print('[INFO] 网络访问异常，先检查并尝试重连WiFi...')
                ensure_wifi_connected(ssid=WIFI_SSID, interface=self.interface)
            except Exception as wifi_error:
                print(f'[ERROR] WiFi重连检查失败: {wifi_error}')

            if self.is_ip_switch:
                set_ip_address(ip_list_path=self.ip_list_path, interface=self.interface, ip_filter=ip,
                               is_time_limit=self.is_time_limit)
            return '无', '无', '无', '视频链接打开失败', '无', start_time, start_time, ip, '无', '无'

        # ========== 第三阶段：关闭登录弹窗 ==========
        login_close_success = False
        try:
            # 尝试多种aria-label模式
            close_btn = None
            aria_labels = ['关闭', 'Close', '关闭对话框', 'Close dialog', '关闭登录', 'Close login']

            for label in aria_labels:
                try:
                    close_btn = self.driver.find_element(By.XPATH, f'//*[@aria-label="{label}"]')
                    if close_btn:
                        print(f'找到关闭按钮: {label}')
                        break
                except:
                    continue

            # 如果aria-label没找到，尝试通过文本内容查找
            if not close_btn:
                try:
                    close_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), '×') or contains(text(), '✕') or contains(text(), '关闭') or contains(text(), 'Close')]")
                    print('通过文本内容找到关闭按钮')
                except:
                    pass

            if close_btn:
                action = ActionChains(self.driver)
                action.move_to_element(close_btn).perform()
                action.click(close_btn).perform()
                print('✅ 成功点击关闭登录按钮')
                login_close_success = True
            else:
                print('⚠️ 未找到关闭登录按钮')

        except Exception as e:
            print(f'⚠️ 关闭登录弹窗异常: {e}')

        time.sleep(2)

        # ========== 第四阶段：获取视频时长信息 ==========
        # 新的HTML结构：通过slider元素获取时长信息
        # <div aria-label="Change Position" aria-valuemax="895.633333" aria-valuenow="36.723402" role="slider">
        fb_time_duration_attribute = '//*[@id="mount_0_0_hb"]/div/div[1]/div/div[3]/div/div/div[1]/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div[2]/div/div[2]/div[2]/div/span[1]'
        fb_time_current_attribute = '//*[@id="mount_0_0_hb"]/div/div[1]/div/div[3]/div/div/div[1]/div[2]/div[1]/div/div/div[1]/div[1]/div/div/div/div[2]/div/div[2]/div[2]/div/span[3]'

        # 尝试获取视频时长和当前播放时长
        video_time_info_obtained = False
        fb_time_duration = '0:00'
        fb_time_current = '0:00'

        # 方法1：尝试通过新的slider元素获取（推荐方法）
        try:
            print('🔍 尝试通过slider元素获取视频时长信息...')
            slider_element = self.driver.find_element(By.XPATH, '//div[@aria-label="Change Position" and @role="slider"]')
            if slider_element:
                # 获取aria-valuemax（视频总时长）和aria-valuenow（当前播放时长）
                aria_valuemax = slider_element.get_attribute('aria-valuemax')
                aria_valuenow = slider_element.get_attribute('aria-valuenow')

                if aria_valuemax and aria_valuenow:
                    # 转换为秒数
                    duration_seconds = float(aria_valuemax)
                    current_seconds = float(aria_valuenow)

                    # 转换为时间格式 (分:秒)
                    duration_minutes = int(duration_seconds // 60)
                    duration_secs = int(duration_seconds % 60)
                    fb_time_duration = f'{duration_minutes}:{duration_secs:02d}'

                    current_minutes = int(current_seconds // 60)
                    current_secs = int(current_seconds % 60)
                    fb_time_current = f'{current_minutes}:{current_secs:02d}'

                    video_time_info_obtained = True
                    print(f'✅ 通过slider元素成功获取视频时长信息:')
                    print(f'   总时长: {aria_valuemax}秒 ({fb_time_duration})')
                    print(f'   当前: {aria_valuenow}秒 ({fb_time_current})')
                else:
                    print('⚠️ slider元素存在但aria属性为空')
        except Exception as e:
            print(f'⚠️ 通过slider元素获取失败: {e}')

        # 方法2：如果方法1失败，尝试通过旧的span元素获取
        if not video_time_info_obtained:
            try:
                print('🔍 尝试通过旧的span元素获取视频时长信息...')
                fb_time_current = self.driver.find_element(by=By.XPATH, value=fb_time_current_attribute).text
                fb_time_duration = self.driver.find_element(by=By.XPATH, value=fb_time_duration_attribute).text
                if fb_time_duration and fb_time_current:
                    video_time_info_obtained = True
                    print(f'✅ 通过span元素成功获取视频时长信息 - 总时长: {fb_time_duration}, 当前: {fb_time_current}')
                else:
                    print('⚠️ span元素文本为空')
            except Exception as e:
                print(f'⚠️ 通过span元素获取失败: {e}')

        # 如果两种方法都失败，设置默认值
        if not video_time_info_obtained:
            print('❌ 所有方法都无法获取视频时长信息，使用默认值')
            fb_time_duration = '0:00'
            fb_time_current = '0:00'

        # ========== 第五阶段：获取当前播放画质 ==========
        resolution_nums = '自动'  # 默认使用自动画质
        video_itag = '-'
        current_quality_obtained = False

        # 鼠标移动到播放器以显示控制栏
        try:
            action = ActionChains(self.driver)
            action.move_to_element(self.driver.find_element(by=By.XPATH, value='//div[@aria-label="Video player"]')).perform()
            action.move_by_offset(1, 1).perform()
            time.sleep(1)

            # 尝试获取当前播放画质
            try:
                import re
                # 查找包含选中标记的元素
                quality_selectors = [
                    '//span[contains(@class, "x3nfvp2") or contains(@class, "xwklpps")]/preceding-sibling::span[1]',
                    '//span[contains(@class, "x3nfvp2") or contains(@class, "xwklpps")]/../span[1]',
                    '//span[contains(@class, "x3nfvp2") or contains(@class, "xwklpps")]/parent::*/span[position()=1]',
                ]

                for selector in quality_selectors:
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            text = elem.text.strip()
                            if not text:
                                continue

                            # 检查是否是"自动"
                            if text == "自动" or text.lower() == "auto":
                                resolution_nums = "自动"
                                current_quality_obtained = True
                                print(f'✅ 当前画质: {resolution_nums}')
                                break
                            # 检查是否是数字画质 (如720p)
                            elif re.match(r'^\d+[pP]$', text):
                                resolution_nums = text
                                current_quality_obtained = True
                                print(f'✅ 当前画质: {resolution_nums}')
                                break

                        if current_quality_obtained:
                            break
                    except Exception:
                        continue

                # 如果上述方法失败，尝试遍历包含选中标记的元素
                if not current_quality_obtained:
                    checkmark_spans = self.driver.find_elements(By.XPATH, '//span[contains(@class, "x3nfvp2") or contains(@class, "xwklpps")]')
                    for checkmark in checkmark_spans:
                        try:
                            parent = checkmark.find_element(By.XPATH, '..')
                            all_spans = parent.find_elements(By.TAG_NAME, 'span')
                            if len(all_spans) >= 2:
                                text = all_spans[0].text.strip()
                                if text == "自动" or text.lower() == "auto":
                                    resolution_nums = "自动"
                                    current_quality_obtained = True
                                    print(f'✅ 当前画质: {resolution_nums}')
                                    break
                                elif re.match(r'^\d+[pP]$', text):
                                    resolution_nums = text
                                    current_quality_obtained = True
                                    print(f'✅ 当前画质: {resolution_nums}')
                                    break
                        except Exception as e:
                            continue

            except Exception as e:
                print(f'⚠️ 获取当前画质失败: {e}')
                resolution_nums = '自动'

        except Exception as e:
            print(f'⚠️ 移动鼠标到播放器失败: {e}')
            resolution_nums = '自动'

        # ========== 第六阶段：尝试打开画质清单并切换到预设画质 ==========
        quality_menu_opened = False
        preset_resolution_switched = False

        try:
            # 打开设置菜单
            action = ActionChains(self.driver)
            action.move_to_element(self.driver.find_element(by=By.XPATH, value='//div[@aria-label="Video player"]')).perform()
            action.move_by_offset(1, 1).perform()
            time.sleep(1)

            try:
                self.driver.find_element(By.XPATH, '//div[@aria-label="设置" and @role="button"]').click()
                print('✅ 成功打开设置菜单')
                time.sleep(1)

                # 点击画质按钮
                try:
                    quality_menu_item = self.driver.find_element(By.XPATH,
                        '//div[contains(text(), "画质") or contains(text(), "Quality") and @role="button"]')
                    quality_menu_item.click()
                    print('✅ 成功打开画质菜单')
                    quality_menu_opened = True
                    time.sleep(1)
                except Exception as e:
                    print(f'⚠️ 打开画质菜单失败: {e}')

            except Exception as e:
                print(f'⚠️ 打开设置菜单失败: {e}')

        except Exception as e:
            print(f'⚠️ 移动鼠标失败: {e}')

        # 如果画质清单打开成功，尝试读取并切换画质
        if quality_menu_opened:
            try:
                # 先尝试找到画质菜单容器（可能是role="dialog"或特定的弹出层）
                quality_menu_container = None
                container_selectors = [
                    '//div[@role="dialog"]',
                    '//div[contains(@class, "x1n2onr6")]',  # Facebook常见的弹出层class
                    '//div[@aria-label="画质"]',
                    '//div[@aria-label="Quality"]',
                ]

                for selector in container_selectors:
                    try:
                        quality_menu_container = self.driver.find_element(By.XPATH, selector)
                        if quality_menu_container:
                            print(f'✅ 找到画质菜单容器: {selector}')
                            break
                    except Exception:
                        continue

                if quality_menu_container:
                    # 在容器内查找画质选项
                    # 根据实际HTML结构调整：画质文字通常在 class="x1ypdohk x1rg5ohu" 的div中
                    # 当前选中的画质会额外包含 x1s688f 类
                    fb_menu_items = quality_menu_container.find_elements(
                        By.XPATH, './/div[contains(@class, "x1ypdohk") and contains(@class, "x1rg5ohu")]'
                    )

                    print(f'📋 在菜单容器内找到 {len(fb_menu_items)} 个选项')
                    target_resolution_found = False
                    available_qualities = []

                    for menu_item in fb_menu_items:
                        try:
                            # 获取画质文字
                            menu_text = menu_item.text.strip()

                            if not menu_text:
                                continue

                            # 获取当前元素的class属性
                            class_attr = menu_item.get_attribute('class') or ''

                            # 判断是否为当前选中的画质（包含 x1s688f）
                            is_selected = 'x1s688f' in class_attr

                            # 过滤：只保留画质相关的选项（"自动" 或 数字+p 格式）
                            import re
                            if menu_text == "自动" or re.match(r'^\d+[pP]$', menu_text):
                                available_qualities.append(menu_text)
                                if is_selected:
                                    print(f'   ✓ 当前画质: {menu_text} (选中)')
                                else:
                                    print(f'   ✓ 画质选项: {menu_text}')

                                # 检查是否匹配预设分辨率（支持 "720" 和 "720P" 等格式）
                                # 移除可能的 'P' 或 'p' 后缀进行比较
                                resolution_clean = resolution.lower().replace('p', '')
                                menu_text_clean = menu_text.lower().replace('p', '')

                                if not target_resolution_found and (resolution_clean == menu_text_clean or resolution.lower() in menu_text.lower()):
                                    target_resolution_found = True
                                    # 点击该分辨率选项（点击父级button）
                                    try:
                                        parent_button = menu_item.find_element(By.XPATH, './ancestor::div[@role="button"]')
                                        parent_button.click()
                                        resolution_nums = menu_text
                                        preset_resolution_switched = True
                                        print(f'✅ 成功切换到预设画质: {resolution_nums}')
                                        time.sleep(2)
                                    except Exception as e:
                                        # 如果找不到父级button，尝试直接点击
                                        try:
                                            menu_item.click()
                                            resolution_nums = menu_text
                                            preset_resolution_switched = True
                                            print(f'✅ 成功切换到预设画质: {resolution_nums}')
                                            time.sleep(2)
                                        except Exception as click_error:
                                            print(f'⚠️ 点击分辨率失败: {e}')
                        except Exception as find_error:
                            # 有些元素可能不是画质选项，跳过
                            continue

                    print(f'📋 可用画质选项: {", ".join(available_qualities)}')

                    if not target_resolution_found:
                        print(f'ℹ️ 预设分辨率 {resolution} 不在画质清单中，继续使用当前画质: {resolution_nums}')
                else:
                    print('⚠️ 未读取到画质清单选项，继续使用当前画质')

            except Exception as e:
                print(f'⚠️ 读取画质清单失败: {e}')
        else:
            print('ℹ️ 画质清单未打开，继续使用当前画质')

        print(f'📊 最终画质设置: {resolution_nums}, 视频总时长: {fb_time_duration}')

        # ========== 第七阶段：处理特殊情况下的逻辑 ==========
        # 情况1: 如果关闭登录按钮失败，或者没有获取到视频时长信息
        if not login_close_success or not video_time_info_obtained:
            print('⚠️ 检测到异常情况（登录关闭失败或无法获取视频时长）')
            print(f'📋 将保持chrome窗口达到预设视频播放时长: {play_time_limit}秒')
            print(f'📋 视频播放画质将输出为: 自动')
            print(f'📋 视频实际播放时长将输出为: {play_time_limit}秒')

            # 设置为默认值
            resolution_nums = '自动'
            final_play_time = int(play_time_limit)

            # 等待预设时长
            print(f'⏳ 开始等待 {play_time_limit} 秒...')
            for remaining in range(int(play_time_limit), 0, -10):
                time.sleep(min(10, remaining))
                if remaining % 60 == 0 or remaining <= 10:
                    print(f'⏰ 剩余等待时间: {remaining} 秒')

            print(f'✅ 等待完成，准备关闭浏览器')

            # 记录结束时间
            end_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))

            # 清理资源
            try:
                self.driver.quit()
                time.sleep(1)
                kill_processes(exe_name_list=['chrome'])
                self.ias_opt(self.cmd, self.hostname, self.username, self.password)
                kill_processes(exe_name_list=['tshark'])
            except Exception as e:
                print(f'⚠️ 清理资源时出错: {e}')

            # 返回特殊结果
            print(f'📊 异常情况处理完成 - 画质: 自动, 实际播放时长: {final_play_time}秒')
            return final_play_time, fb_time_duration, resolution_nums, False, video_itag, start_time, end_time, ip, pcap_name, pcap_ias_name

        # ========== 第八阶段：正常播放视频并监控 ==========
        fb_time_current_seconds = time_to_seconds(fb_time_current)
        play_time_limit = int(play_time_limit)
        interept = False
        wait_time = 0
        paly_time_seconds_next = 0
        fb_time_duration_seconds = time_to_seconds(fb_time_duration)

        print(f'📋 开始正常播放流程 - 预设时长: {play_time_limit}秒, 视频总时长: {fb_time_duration_seconds}秒')

        # 检查视频时长是否小于预设播放时长
        if fb_time_duration_seconds < play_time_limit:
            print(f'ℹ️ 视频时长({fb_time_duration_seconds}秒) < 预设播放时长({play_time_limit}秒)')
            print(f'📋 视频实际播放时长将以最后的当前播放时长为准')

        # 开始播放
        try:
            # 尝试找到视频播放器元素并移动鼠标
            video_player_found = False
            video_player_selectors = [
                '//div[@aria-label="Video player"]',
                '//video',
                '//div[contains(@class, "video")]',
                '//*[@id="movie_player"]',  # YouTube风格（备用）
            ]

            for selector in video_player_selectors:
                try:
                    video_player_element = self.driver.find_element(By.XPATH, selector)
                    if video_player_element:
                        action = ActionChains(self.driver)
                        action.move_to_element(video_player_element).perform()
                        action.move_by_offset(1, 1).perform()
                        print(f'✅ 移动鼠标到播放器成功: {selector}')
                        video_player_found = True
                        break
                except Exception:
                    continue

            if not video_player_found:
                print('⚠️ 无法找到视频播放器元素，跳过鼠标移动')

        except Exception as e:
            print(f'⚠️ 移动鼠标到播放器失败: {e}')

        # 尝试控制视频播放
        try:
            # 首先尝试通过JavaScript直接控制video元素
            try:
                is_paused = self.driver.execute_script("""
                    var video = document.querySelector('video');
                    return video ? video.paused : null;
                """)

                if is_paused is True:
                    print('⏸️ 视频暂停，尝试播放...')
                    try:
                        self.driver.execute_script("document.querySelector('video').play();")
                        print('✅ 通过JavaScript成功播放视频')
                    except Exception as js_play_error:
                        print(f'⚠️ JavaScript播放失败: {js_play_error}')
                        # 尝试快捷键
                        try:
                            self.driver.find_element(By.TAG_NAME, 'body').send_keys(' ')
                            print('✅ 通过空格键成功播放视频')
                        except Exception as space_error:
                            print(f'⚠️ 空格键播放也失败: {space_error}')
                elif is_paused is False:
                    print('▶️ 视频已在播放中')
                else:
                    print('⚠️ 无法检测视频播放状态')

            except Exception as js_error:
                print(f'⚠️ JavaScript检测播放状态失败，尝试UI方法: {js_error}')

                # 备用方法：通过UI按钮控制
                play_button_selectors = [
                    '//div[@aria-label="播放" and @role="button"]',
                    '//div[@aria-label="Play" and @role="button"]',
                    '//button[@aria-label="播放"]',
                    '//button[@aria-label="Play"]',
                    '//div[@aria-label="播放"]',
                    '//div[@aria-label="Play"]',
                ]

                play_button_clicked = False
                for selector in play_button_selectors:
                    try:
                        play_button = self.driver.find_element(By.XPATH, selector)
                        if play_button:
                            play_button.click()
                            print(f'✅ 通过UI按钮成功播放: {selector}')
                            play_button_clicked = True
                            break
                    except Exception:
                        continue

                if not play_button_clicked:
                    print('⚠️ 所有播放按钮选择器都失败，尝试快捷键')
                    try:
                        self.driver.find_element(By.TAG_NAME, 'body').send_keys('k')
                        print('✅ 通过k键尝试播放')
                    except Exception as key_error:
                        print(f'⚠️ 快捷键播放失败: {key_error}')

        except Exception as e:
            print(f'⚠️ 播放控制失败: {e}')

        start_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))
        print(f'🕒 视频开始播放时间: {start_time}')

        # ========== 第九阶段：视频播放监控循环 ==========
        while fb_time_current_seconds < play_time_limit:
            time.sleep(1)
            try:
                # 检查视频是否卡住（播放进度没有变化）
                if fb_time_current_seconds == paly_time_seconds_next:
                    wait_time = wait_time + 1
                    if wait_time >= 60:
                        print(f'⚠️ 视频播放卡住超过60秒，可能被封堵')
                        # V1.2新增：卡住时先检查WiFi连接状态
                        try:
                            ensure_wifi_connected(ssid=WIFI_SSID, interface=self.interface)
                        except Exception as wifi_error:
                            print(f'[ERROR] WiFi重连检查失败: {wifi_error}')
                        interept = True
                        if self.is_ip_switch:
                            set_ip_address(ip_list_path=self.ip_list_path, interface=self.interface, ip_filter=ip,
                                           is_time_limit=self.is_time_limit)
                        break
                    else:
                        interept = False
                else:
                    paly_time_seconds_next = fb_time_current_seconds
                    wait_time = 0  # 重置等待时间

                # 尝试通过slider元素获取播放进度（推荐方法）
                slider_time_obtained = False
                try:
                    slider_element = self.driver.find_element(By.XPATH, '//div[@aria-label="Change Position" and @role="slider"]')
                    if slider_element:
                        aria_valuenow = slider_element.get_attribute('aria-valuenow')
                        aria_valuemax = slider_element.get_attribute('aria-valuemax')

                        if aria_valuenow:
                            fb_time_current_seconds = int(float(aria_valuenow))
                            slider_time_obtained = True
                            if aria_valuemax:
                                fb_time_duration_seconds = int(float(aria_valuemax))
                                #print(f'✅ 通过slider获取播放进度: {fb_time_current_seconds}秒 / {fb_time_duration_seconds}秒')
                            else:
                                print(f'✅ 通过slider获取播放进度: {fb_time_current_seconds}秒')
                except Exception as slider_error:
                    print(f'⚠️ slider获取失败: {slider_error}')

                # 如果slider方法失败，尝试通过JavaScript获取
                if not slider_time_obtained:
                    try:
                        fb_time_current_seconds, _ = self.driver.execute_script(
                            "var video = document.querySelector('video'); "
                            "if (video) return [video.currentTime, video.duration]; "
                            "return [0, 0];"
                        )
                        fb_time_current_seconds = int(fb_time_current_seconds)
                        print(f'✅ 通过JavaScript获取播放进度: {fb_time_current_seconds}秒')
                    except Exception as js_error:
                        print(f'⚠️ JavaScript获取播放进度失败: {js_error}')
                        # 尝试通过DOM元素获取
                        try:
                            ytp_time_current = self.driver.find_element(by=By.XPATH, value=fb_time_current_attribute).text
                            fb_time_current_seconds = time_to_seconds(ytp_time_current)
                            print(f'✅ 通过DOM元素获取播放进度: {fb_time_current_seconds}秒')
                        except Exception as dom_error:
                            print(f'⚠️ DOM获取播放进度也失败: {dom_error}')
                            fb_time_current_seconds = paly_time_seconds_next  # 使用上次的值

                #print(f'📊 播放进度: {fb_time_current_seconds}秒 / {play_time_limit}秒 (预设), 视频总时长: {fb_time_duration_seconds}秒')

                # 移动鼠标保持视频控制栏可见
                try:
                    action = ActionChains(self.driver)

                    # 尝试多种方式查找视频播放器元素
                    video_player = None
                    video_player_selectors = [
                        ('By.CLASS_NAME', "html5-video-player"),
                        ('By.XPATH', '//div[@aria-label="Video player"]'),
                        ('By.CSS_SELECTOR', '[aria-label="Video player"]'),
                        ('By.TAG_NAME', 'video'),
                    ]

                    for selector_type, selector_value in video_player_selectors:
                        try:
                            if selector_type == 'By.CLASS_NAME':
                                video_player = self.driver.find_element(By.CLASS_NAME, selector_value)
                            elif selector_type == 'By.XPATH':
                                video_player = self.driver.find_element(By.XPATH, selector_value)
                            elif selector_type == 'By.CSS_SELECTOR':
                                video_player = self.driver.find_element(By.CSS_SELECTOR, selector_value)
                            elif selector_type == 'By.TAG_NAME':
                                video_player = self.driver.find_element(By.TAG_NAME, selector_value)

                            if video_player:
                                #print(f'✅ 找到视频播放器元素: {selector_type} = {selector_value}')
                                break
                        except Exception:
                            continue

                    if video_player:
                        action.move_to_element(video_player).perform()
                        time.sleep(0.2)

                        # 尝试通过JavaScript触发鼠标事件
                        try:
                            self.driver.execute_script("""
                                var videoElement = document.querySelector('video') || document.querySelector('[aria-label="Video player"]');
                                if (videoElement) {
                                    videoElement.dispatchEvent(new MouseEvent('mousemove', {bubbles: true}));
                                }
                            """)
                        except Exception as js_error:
                            print(f'⚠️ JavaScript触发鼠标事件失败: {js_error}')

                        action.move_by_offset(1, 1).perform()
                    else:
                        print('⚠️ 所有视频播放器选择器都失败，跳过鼠标移动')

                except Exception as move_error:
                    print(f'⚠️ 移动鼠标失败: {move_error}')

                # 尝试获取更新的视频时长
                try:
                    # 方法1：通过slider元素获取（推荐）
                    slider_update_success = False
                    try:
                        slider_element = self.driver.find_element(By.XPATH, '//div[@aria-label="Change Position" and @role="slider"]')
                        if slider_element:
                            aria_valuemax = slider_element.get_attribute('aria-valuemax')
                            aria_valuenow = slider_element.get_attribute('aria-valuenow')

                            if aria_valuemax and aria_valuenow:
                                # 转换为时间格式
                                duration_seconds = float(aria_valuemax)
                                current_seconds = float(aria_valuenow)

                                duration_minutes = int(duration_seconds // 60)
                                duration_secs = int(duration_seconds % 60)
                                fb_time_duration = f'{duration_minutes}:{duration_secs:02d}'

                                current_minutes = int(current_seconds // 60)
                                current_secs = int(current_seconds % 60)
                                ytp_time_current = f'{current_minutes}:{current_secs:02d}'

                                # 更新秒数
                                fb_time_duration_seconds = int(duration_seconds)
                                slider_update_success = True
                                #print(f'✅ 通过slider更新时长 - 总: {fb_time_duration}, 当前: {ytp_time_current}')
                    except Exception as slider_error:
                        print(f'⚠️ slider更新失败: {slider_error}')

                    # 方法2：如果slider失败，使用旧的span元素方法
                    if not slider_update_success:
                        ytp_time_current = self.driver.find_element(by=By.XPATH, value=fb_time_current_attribute).text
                        fb_time_duration = self.driver.find_element(by=By.XPATH, value=fb_time_duration_attribute).text

                        if len(fb_time_duration) < 5:
                            print('⚠️ 视频时长格式异常，尝试重新获取')
                            self.driver.find_element(By.TAG_NAME, 'body').send_keys('C')
                            time.sleep(0.2)
                            ytp_time_current = self.driver.find_element(by=By.XPATH, value=fb_time_current_attribute).text
                            fb_time_duration = self.driver.find_element(by=By.XPATH, value=fb_time_duration_attribute).text
                            print(f'📋 重新获取 - 当前: {ytp_time_current}, 总时长: {fb_time_duration}')

                        # 更新视频时长
                        fb_time_duration_seconds = time_to_seconds(fb_time_duration)
                        print(f'✅ 通过span元素更新时长 - 总: {fb_time_duration} ({fb_time_duration_seconds}秒), 当前: {ytp_time_current}')

                except Exception as time_error:
                    print(f'⚠️ 获取视频时长失败: {time_error}')

                # 检查是否播放结束（视频时长 < 预设播放时长）
                try:
                    if fb_time_current_seconds >= fb_time_duration_seconds and fb_time_duration_seconds > 0:
                        print(f'✅ 视频播放完成 - 当前: {fb_time_current_seconds}秒 >= 总时长: {fb_time_duration_seconds}秒')
                        print(f'📋 由于视频时长({fb_time_duration_seconds}秒) < 预设播放时长({play_time_limit}秒)')
                        print(f'📋 视频实际播放时长=最后的当前播放时长: {fb_time_current_seconds}秒')
                        break
                except Exception as e:
                    print(f'⚠️ 时间比较异常: {e}')

                print(f'📹 播放状态 - 视频ID: {vid}, 已播放: {fb_time_current_seconds}秒, 总时长: {fb_time_duration}秒, 画质: {resolution_nums}, 封堵: {interept}, 等待: {wait_time}秒')

            except Exception as e:
                print(f'⚠️ 播放监控异常: {e}')
                break

        # ========== 第十阶段：清理和返回结果 ==========
        end_time = time.strftime("%Y-%m-%d_%H:%M:%S", time.localtime(time.time()))

        # 确定实际播放时长
        if interept:
            # 情况1: 检测到封堵(interept=true)，使用实际播放时长
            actual_play_time = fb_time_current_seconds
        else:
            # 情况2: 未检测到封堵(interept=false)
            if play_time_limit >= fb_time_duration_seconds:
                # 预设时长 >= 视频总时长，使用视频总时长
                actual_play_time = fb_time_duration_seconds
            else:
                # 预设时长 < 视频总时长，使用预设时长
                actual_play_time = play_time_limit

        print(f'🎬 视频播放完成 - 实际播放时长: {actual_play_time}秒')
        print(f'📊 最终结果统计:')
        print(f'    视频ID: {vid}')
        print(f'    预设画质: {resolution}')
        print(f'    实际播放画质: {resolution_nums}')
        print(f'    预设播放时长: {play_time_limit}秒')
        print(f'    视频总时长: {fb_time_duration} ({fb_time_duration_seconds}秒)')
        print(f'    实际播放时长: {actual_play_time}秒')
        print(f'    开始时间: {start_time}')
        print(f'    结束时间: {end_time}')
        print(f'    播放IP: {ip}')
        print(f'    是否封堵: {interept}')
        print(f'    PCAP文件: {pcap_name}')
        print(f'    IAS_PCAP文件: {pcap_ias_name}')

        # ========== 第十一阶段：清理资源（严格顺序） ==========
        print('[INFO] ========== 开始清理资源 ==========')

        # 11.1 关闭浏览器
        try:
            if self.driver:
                self.driver.quit()
                time.sleep(1)
                print('[OK] 浏览器已关闭')
        except Exception as e:
            print(f'[ERROR] 关闭浏览器失败: {e}')

        # 11.2 停止 PC 抓包
        try:
            print('[INFO] 停止 PC 抓包...')
            kill_processes(exe_name_list=['tshark'])
            time.sleep(2)
            print('[OK] PC 抓包已停止')
        except Exception as e:
            print(f'[ERROR] 停止 PC 抓包失败: {e}')

        # 11.3 停止 IAS 抓包（并等待完全停止）
        if self.is_ias_pcap and self.tcpdump_pid:
            try:
                print('[INFO] 停止 IAS 抓包...')
                kill_cmd = 'kill -15 {0}'.format(self.tcpdump_pid)
                print(f'[INFO] 执行命令: {kill_cmd}')

                # 执行kill命令
                self.ias_opt(kill_cmd, self.hostname, self.username, self.password)
                print('[OK] IAS 抓包进程已发送停止信号')

                # 等待进程完全停止
                print('[INFO] 等待 IAS 抓包进程完全停止...')
                time.sleep(3)

                # 验证进程是否已停止
                stopped = False
                for i in range(10):
                    try:
                        client = paramiko.SSHClient()
                        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                        sock.bind(('10.10.163.86', 0))
                        sock.settimeout(5)
                        sock.connect((self.hostname, 22))
                        transport = paramiko.Transport(sock)
                        transport.connect(username=self.username, password=self.password)
                        client._transport = transport
                        stdin, stdout, stderr = client.exec_command(f'ps -p {self.tcpdump_pid}')
                        exit_status = stdout.channel.recv_exit_status()
                        client.close()

                        # exit_status = 1 表示进程不存在
                        if exit_status != 0:
                            stopped = True
                            print(f'[OK] IAS 抓包进程已完全停止 (验证: 第{i+1}次)')
                            break
                    except Exception as e:
                        # 如果连接失败，可能进程已停止
                        stopped = True
                        print(f'[OK] IAS 抓包进程已完全停止 (验证: 第{i+1}次，异常: {e})')
                        break

                    time.sleep(1)

                if not stopped:
                    print('[WARN] IAS 抓包进程可能未完全停止，但继续执行')

                # 清除PID状态
                self.tcpdump_pid = None
                print('[OK] IAS 抓包已完全停止')
            except Exception as e:
                print(f'[ERROR] 停止 IAS 抓包失败: {e}')
        elif self.is_ias_pcap and not self.tcpdump_pid:
            print('[WARN] IAS 捕包已启用但没有有效的PID')
        else:
            print('[INFO] IAS 捕包未启用，跳过')

        # 11.4 清理残留的 Chrome 进程
        try:
            print('[INFO] 清理残留的 Chrome 进程...')
            kill_processes(exe_name_list=['chrome'])
            time.sleep(1)
            print('[OK] 残留进程清理完成')
        except Exception as e:
            print(f'[ERROR] 清理残留进程失败: {e}')

        print('[OK] ========== 所有资源清理完成 ==========')       

        # IP切换
        if self.is_ip_switch:
            try:
                set_ip_address(ip_list_path=self.ip_list_path, interface=self.interface, ip_filter=ip,
                               is_time_limit=self.is_time_limit)
                print('✅ IP切换完成')
            except Exception as e:
                print(f'⚠️ IP切换失败: {e}')

        # 返回结果
        return actual_play_time, fb_time_duration, resolution_nums, interept, video_itag, start_time, end_time, ip, pcap_name, pcap_ias_name

    def one_inster(self, *args):
        # root_path、pcap_path是否存在，不存在创建
        if not os.path.exists(self.root_path):
            os.makedirs(self.root_path)
        if not os.path.exists(self.pcap_path):
            os.makedirs(self.pcap_path)
        vid_dict = self.get_data_dict()
        # 创建播放失败表格文件
        errorfile_name = os.path.join(self.root_path,
                                      'fb播放失败URL列表' + time.strftime("%Y%m%d_%H%M%S", time.localtime(
                                          time.time())) + '.txt')
        # 创建结果表格和sheet并立即关闭
        workbook1 = xlsxwriter.Workbook(self.excel_name)
        workbook1.add_worksheet()
        workbook1.close()
        # 价值刚创建的表格，获取名为Sheet1的工作表
        workbook = openpyxl.load_workbook(self.excel_name)
        sheet = workbook['Sheet1']
        # 将表头list_first添加到sheet，并保存
        list_first = ['vid', '画质', '实际播放画质', '视频总时长', 'itag', '是否封堵成功', '预设视频播放时长(秒)',
                      '视频实际播放时长(秒)', '视频开始播放时间', '视频停止播放时间', '播放IP', 'PCAP文件名',
                      'IAS_PCAP文件名']
        sheet._current_row = sheet.max_row
        sheet.append(list_first)
        workbook.save(self.excel_name)
        num = 1
        #注释pcap目录的逻辑，避免多次运行时清空之前的文件
        #play_status = []
        #for k, v in vid_dict.items():
        #    play_status.append(v[3])
        #if '已播放' not in play_status:
        #    shutil.rmtree(self.pcap_path)
        for k, v in vid_dict.items():
            num = num + 1
            if v[3] == '已播放':
                continue
            play_time_limit = v[2] if v[2] is not None else 60  # 如果为空，使用默认值60秒
            try:
                fb_time_current_seconds, fb_time_duration, resolution_nums, interept, video_itag, start_time, end_time, ip, pcap_name, pcap_ias_name = self.open_fb(
                    vid=str(v[0]).strip(), play_time_limit=int(play_time_limit), resolution=v[1])
                video_itag = '-'
                print(fb_time_current_seconds, resolution_nums, interept, video_itag, start_time, end_time, ip,
                      pcap_name, pcap_ias_name)
                inster_value = [v[0], v[1], resolution_nums, fb_time_duration, video_itag, interept, play_time_limit,
                                fb_time_current_seconds, start_time, end_time, ip, pcap_name, pcap_ias_name]
                workbook = openpyxl.load_workbook(self.excel_name)
                sheet = workbook['Sheet1']
                sheet._current_row = sheet.max_row
                sheet.append(inster_value)
                workbook.save(self.excel_name)
            except Exception as e:
                print('播放异常', v[0], e)
                with open(errorfile_name, 'a+') as f:
                    f.write(str(v[0]) + str(e))
                    f.write('\n')

            # 标记视频为已播放
            edit_workbook = openpyxl.load_workbook(self.vid_file)
            edit_sheet = edit_workbook.active
            edit_cell = 'E' + str(num)
            edit_sheet[edit_cell] = '已播放'
            edit_workbook.save(self.vid_file)
