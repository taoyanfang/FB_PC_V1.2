"""入口：组装配置并启动 FB 播放任务（原 FB_PC_V1.2.py 的 __main__ 逻辑）"""

import config as cfg
from utils import kill_processes, ensure_wifi_connected
from business import FB


# 主函数
if __name__ == '__main__':
    print('========== FB_PC V1.2（新增WiFi自动重连功能） ==========')
    kill_processes(exe_name_list=['chrome', 'tshark'])
    # 路径类配置（root_path/pcap_path/excel_name/create_cmd）已在 config.py 中按启动时间生成

    # V1.2新增：启动时检查WiFi连接状态，未连接时自动连接指定WiFi
    ensure_wifi_connected(ssid=cfg.WIFI_SSID, interface=cfg.INTERFACE)
    fb_obj = FB(cfg.ROOT_PATH, cfg.PCAP_PATH, cfg.EXCEL_NAME, cfg.WIRESHARK_PATH, cfg.ETH_NAME,
                cfg.IP_LIST_PATH, cfg.INTERFACE, cfg.TCPDUMP_PATH, cfg.IAS_ETH, cfg.HOSTNAME,
                cfg.USERNAME, cfg.PASSWORD, cfg.IS_SCREEN, cfg.IS_IAS_PCAP, cfg.VID_FILE,
                cfg.FB_USERNAME, cfg.FB_PASSWORD, cfg.IS_TIME_LIMIT, cfg.IS_IP_SWITCH)
    if cfg.IS_IAS_PCAP:
        fb_obj.ias_opt(cfg.CREATE_CMD)
    # fb_obj.login()
    fb_obj.one_inster()
