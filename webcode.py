import argparse
import threading
from logic import open_xlsx,process_data_in_threads,save_xlsx

from logo import logo

#传入参数
def parse_args():
    # 1. 定义命令行解析器对象
    parser = argparse.ArgumentParser(description='webcode是一个批量扫描ip的项目')

    # 2. 添加命令行参数
    parser.add_argument('-th',type=int,default=50,help="设置线程数量，默认是50线程")
    parser.add_argument('-ti', type=int,default=5,help="设置延迟时间，默认为5秒")
    # 3. 从命令行中结构化解析参数
    args = parser.parse_args()
    return args



if __name__ == '__main__':
    logo()
    args = parse_args()
    # 设置最大的线程
    thread = args.th
    # 设置request最大的超时时间
    waittime = args.ti
    data_list = open_xlsx()
    result = process_data_in_threads(data_list,thread,waittime)
    save_xlsx(result)