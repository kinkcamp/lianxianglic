import requests
import json
import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import messagebox
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from datetime import datetime
import os
import csv
import time
from typing import Dict, List, Set, Any
from dataclasses import dataclass
from queue import Queue, Empty
from openpyxl import Workbook

@dataclass
class QueryResult:
    """查询结果数据类"""
    serial: str
    index: int
    total: int
    success: bool
    data: dict
    retry_count: int = 0
    valid_services: int = 0
    expired_services: int = 0
    total_services: int = 0

    def to_dict(self):
        """转换为字典以便JSON序列化"""
        return {
            'serial': self.serial,
            'index': self.index,
            'total': self.total,
            'success': self.success,
            'data': self.data,
            'retry_count': self.retry_count,
            'valid_services': self.valid_services,
            'expired_services': self.expired_services,
            'total_services': self.total_services
        }
    
    @classmethod
    def from_dict(cls, data):
        """从字典创建QueryResult对象"""
        return cls(**data)

class WarrantyCheckerApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        self.initialize_system()

    def setup_ui(self):
        """初始化UI组件"""
        self.root.title("联想保修批量查询")
        self.root.geometry("1000x800")
        self.root.minsize(800, 600)
        
        # 创建主框架和组件
        self.create_main_frame()
        self.create_input_area()
        self.create_result_area()
        self.setup_grid_weights()

    def initialize_system(self):
        """初始化系统组件"""
        self.text_lock = threading.Lock()
        self.executor = ThreadPoolExecutor(
            max_workers=96,
            thread_name_prefix="warranty_query"
        )
        self.message_queue = Queue()
        self.query_cache: Dict[str, QueryResult] = {}
        self.query_results: Dict[str, QueryResult] = {}
        
        # 系统参数
        self.batch_size = 200
        self.min_interval = 0.01
        self.max_retries = 2
        self.timeout = 3
        
        # 启动消息处理
        self.root.after(100, self.process_message_queue)
        self.load_previous_results()

    def query_with_retry(self, serial: str, index: int, total: int) -> QueryResult:
        """优化的查询函数"""
        if serial in self.query_cache:
            return self.query_cache[serial]

        session = requests.Session()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Referer': 'https://newsupport.lenovo.com.cn/',
            'Connection': 'keep-alive'
        }

        for retry_count in range(self.max_retries + 1):
            try:
                response = session.get(
                    f"https://newsupport.lenovo.com.cn/api/drive/{serial}/drivewarrantyinfo",
                    headers=headers,
                    timeout=self.timeout
                )
                response.raise_for_status()
                data = response.json()

                if data['statusCode'] == 200:
                    result = QueryResult(
                        serial=serial,
                        index=index,
                        total=total,
                        success=True,
                        data=data,
                        retry_count=retry_count
                    )
                    self.query_cache[serial] = result
                    return result

                if retry_count < self.max_retries:
                    time.sleep(0.1)
                    continue

                return QueryResult(
                    serial=serial,
                    index=index,
                    total=total,
                    success=False,
                    data={'error': f"API返回状态码错误: {data['statusCode']}"},
                    retry_count=retry_count
                )

            except Exception as e:
                if retry_count < self.max_retries:
                    time.sleep(0.1)
                    continue

                return QueryResult(
                    serial=serial,
                    index=index,
                    total=total,
                    success=False,
                    data={'error': str(e)},
                    retry_count=retry_count
                )

    def check_warranty(self):
        """执行保修查询"""
        serials = self.parse_serial_numbers(self.serial_text.get("1.0", tk.END))
        if not self.validate_query(serials):
            return

        self.prepare_query(len(serials))
        progress = self.create_progress_bar(len(serials))

        try:
            self.execute_query(serials, progress)
        finally:
            self.cleanup_query(progress)

    def execute_query(self, serials: List[str], progress: ttk.Progressbar):
        """执行查询逻辑"""
        completed = 0
        failed_serials: Set[str] = set()
        futures = []

        # 提交所有查询任务
        for index, serial in enumerate(serials, 1):
            future = self.executor.submit(self.query_with_retry, serial, index, len(serials))
            futures.append(future)
            if len(futures) % 20 == 0:
                self.root.update()

        # 处理查询结果
        for future in as_completed(futures):
            try:
                result = future.result()
                completed += 1
                progress['value'] = completed
                
                self.update_result_text(result)
                if completed % 20 == 0:
                    self.root.update()

                if result.success and result.data['statusCode'] == 200:
                    self.query_results[result.serial] = result
                    if completed % 100 == 0:
                        self.save_results()
                else:
                    failed_serials.add(result.serial)

            except Exception as e:
                print(f"处理结果时出错: {e}")

        self.save_results()
        self.handle_failed_queries(failed_serials)

    def handle_failed_queries(self, failed_serials: Set[str]):
        """处理失败的查询"""
        if failed_serials:
            # 不要清空输入框，只显示失败信息
            self.message_queue.put(f"\n共有 {len(failed_serials)} 个序列号查询失败")
            # 保留失败的序列号在输入框中，但不清除已有的查询结果
            failed_text = "\n".join(failed_serials)
            self.serial_text.delete("1.0", tk.END)
            self.serial_text.insert("1.0", failed_text)

    # ... (其他方法保持不变，但建议类似地优化代码结构)

    def parse_serial_numbers(self, text):
        """
        处理输入文本，支持多种分隔符并验证序列号格式
        """
        text = text.strip()
        if not text:
            return []
        
        # 首先按换行符分割
        items = text.splitlines()
        # 将每一行再按逗号和空格分割
        result = []
        invalid_serials = []
        
        for item in items:
            # 使用正则表达式分割（支持逗号、空格、制表符等）
            parts = re.split(r'[,\s]+', item.strip())
            for part in parts:
                serial = part.strip()
                if serial:
                    # 验证序列号格式（假设联想序列号为字母数字组合，长度在8-20位之间）
                    if re.match(r'^[A-Za-z0-9]{8,20}$', serial):
                        result.append(serial.upper())  # 转换为大写
                    else:
                        invalid_serials.append(serial)
        
        # 检查重复的序列号
        seen = set()
        duplicates = []
        unique_result = []
        
        for serial in result:
            if serial not in seen:
                seen.add(serial)
                unique_result.append(serial)
            else:
                duplicates.append(serial)
        
        # 如果有无效或重复的序列号，显示警告
        warning_message = ""
        if invalid_serials:
            warning_message += f"发现 {len(invalid_serials)} 个无效的序列号:\n"
            warning_message += ", ".join(invalid_serials) + "\n\n"
        
        if duplicates:
            warning_message += f"发现 {len(duplicates)} 个重复的序列号:\n"
            warning_message += ", ".join(duplicates)
        
        if warning_message:
            messagebox.showwarning("序列号验证", warning_message)
        
        return unique_result

    def process_message_queue(self):
        """处理消息队列中的UI更新请求"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                if isinstance(message, dict):
                    self._update_result_display(message)
                else:
                    self.result_text.insert(tk.END, str(message) + "\n")
                self.result_text.see(tk.END)
        except Empty:
            pass
        finally:
            # 继续监听消息队列
            self.root.after(100, self.process_message_queue)

    def _update_result_display(self, result: QueryResult):
        """实际更新显示结果的函数"""
        try:
            self.result_text.insert(tk.END, f"\n序列号 {result.index}/{result.total}: {result.serial}\n")
            self.result_text.insert(tk.END, "="*80 + "\n")
            
            if result.success:
                data = result.data
                if data['statusCode'] == 200:
                    detail_info = data['data'].get('detailinfo', {})
                    service_types = ['warranty', 'onsite', 'other']
                    
                    valid_services = 0
                    expired_services = 0
                    
                    for service_type in service_types:
                        if service_type in detail_info:
                            self.result_text.insert(tk.END, f"\n【{service_type}类型服务】\n")
                            self.result_text.insert(tk.END, "-"*40 + "\n")
                            
                            for item in detail_info[service_type]:
                                remaining_days = int(item.get('DateDifference', 0))
                                warranty_status = "在保" if remaining_days > 0 else "已过保"
                                
                                if remaining_days > 0:
                                    valid_services += 1
                                else:
                                    expired_services += 1
                                
                                self.result_text.insert(tk.END, f"服务名称: {item.get('ServiceProductName', '未知')}\n")
                                self.result_text.insert(tk.END, f"开始时间: {item.get('StartDate', '未知')}\n")
                                self.result_text.insert(tk.END, f"结束时间: {item.get('EndDate', '未知')}\n")
                                self.result_text.insert(tk.END, f"剩余天数: {remaining_days} 天\n")
                                self.result_text.insert(tk.END, f"保修状态: {warranty_status}\n")
                                if 'ServiceDescription' in item:
                                    self.result_text.insert(tk.END, f"服务描述: {item['ServiceDescription']}\n")
                                self.result_text.insert(tk.END, "-"*40 + "\n")
                    
                    result.valid_services = valid_services
                    result.expired_services = expired_services
                    result.total_services = valid_services + expired_services
                    
                    self.result_text.insert(tk.END, f"\n服务统计:\n")
                    self.result_text.insert(tk.END, f"在保服务数量: {valid_services}\n")
                    self.result_text.insert(tk.END, f"过保服务数量: {expired_services}\n")
                    self.result_text.insert(tk.END, f"总服务数量: {valid_services + expired_services}\n")
                else:
                    self.result_text.insert(tk.END, f"查询失败: {data.get('message', '未知错误')}\n")
                    self.result_text.insert(tk.END, "建议稍后重试此序列号\n")
            else:
                error_msg = result.data.get('error', '未知错误') if isinstance(result.data, dict) else str(result.data)
                self.result_text.insert(tk.END, f"查询出错: {error_msg}\n")
                self.result_text.insert(tk.END, "可能是网络问题，建议稍后重试此序列号\n")
            
            self.result_text.insert(tk.END, "\n" + "="*80 + "\n")
            self.result_text.see(tk.END)  # 自动滚动到底部
            
        except Exception as e:
            print(f"更新显示出错: {e}")  # 用于调试

    def update_result_text(self, result: QueryResult):
        """将结果放入消息队列"""
        with self.text_lock:
            self._update_result_display(result)

    def load_previous_results(self):
        """加载之前的查询结果"""
        self.query_results = {}
        try:
            if os.path.exists("query_results.json"):
                with open("query_results.json", "r", encoding='utf-8') as f:
                    saved_results = json.load(f)
                    for serial, result_dict in saved_results.items():
                        self.query_results[serial] = QueryResult.from_dict(result_dict)
        except Exception as e:
            print(f"加载历史记录失败: {e}")

    def save_results(self):
        """保存查询结果到文件"""
        try:
            # 将QueryResult对象转换为可序列化的字典
            serializable_results = {
                serial: result.to_dict() 
                for serial, result in self.query_results.items()
            }
            with open("query_results.json", "w", encoding='utf-8') as f:
                json.dump(serializable_results, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存结果失败: {e}")

    def clear_input(self):
        """只清除输入框"""
        self.serial_text.delete("1.0", tk.END)

    def clear_all(self):
        """清除所有内容"""
        if messagebox.askyesno("确认", "确定要清除所有内容吗？\n这将清除输入框、查询结果和导出记录！"):
            self.serial_text.delete("1.0", tk.END)
            self.result_text.delete("1.0", tk.END)
            self.query_results.clear()  # 清空存储的结果
            if os.path.exists("query_results.json"):
                os.remove("query_results.json")  # 删除结果文件

    def export_to_csv(self):
        """导出结果到Excel文件"""
        try:
            # 获取当前输入框中的序列号列表作为排序参考
            current_serials = self.parse_serial_numbers(self.serial_text.get("1.0", tk.END))
            if not current_serials:
                messagebox.showwarning("警告", "请先输入要导出的序列号！")
                return
            
            # 检查是否有查询结果
            has_results = any(serial in self.query_results for serial in current_serials)
            if not has_results:
                messagebox.showwarning("警告", "没有可导出的查询结果！")
                return

            if not os.path.exists("导出结果"):
                os.makedirs("导出结果")
            
            filename = f"导出结果/联想保修查询结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            wb = Workbook()
            ws_summary = wb.active
            ws_summary.title = "汇总信息"
            ws_detail = wb.create_sheet("详细信息")
            
            # 统计数据
            total_queries = len(current_serials)
            success_queries = sum(1 for serial in current_serials 
                                if serial in self.query_results 
                                and self.query_results[serial].success 
                                and self.query_results[serial].data['statusCode'] == 200)
            failed_queries = total_queries - success_queries
            
            # 收集在保和过保的序列号
            in_warranty_serials = []
            out_warranty_serials = []
            
            # 收集所有序列号的信息
            in_warranty_info = []
            out_warranty_info = []
            all_serials_info = []
            
            # 处理每个序列号
            for serial in current_serials:
                if serial in self.query_results:
                    result = self.query_results[serial]
                    if result.success and result.data['statusCode'] == 200:
                        # 获取最新的维保信息
                        latest_start_date = None
                        latest_end_date = None
                        remaining_days = 0
                        detail_info = result.data['data'].get('detailinfo', {})
                        
                        for service_type in ['warranty', 'onsite', 'other']:
                            if service_type in detail_info:
                                for item in detail_info[service_type]:
                                    end_date = item.get('EndDate', '')
                                    start_date = item.get('StartDate', '')
                                    current_remaining_days = int(item.get('DateDifference', 0))
                                    if end_date:
                                        try:
                                            current_end_date = datetime.strptime(end_date, '%Y-%m-%d')
                                            current_start_date = datetime.strptime(start_date, '%Y-%m-%d')
                                            if latest_end_date is None or current_end_date > latest_end_date:
                                                latest_end_date = current_end_date
                                                latest_start_date = current_start_date
                                                remaining_days = current_remaining_days
                                        except ValueError:
                                            continue
                        
                        if latest_start_date and latest_end_date:
                            info = (
                                serial,
                                latest_start_date.strftime('%Y-%m-%d'),
                                latest_end_date.strftime('%Y-%m-%d'),
                                remaining_days
                            )
                            all_serials_info.append(info)
                            
                            if remaining_days > 0:
                                in_warranty_serials.append(serial)
                                in_warranty_info.append(info)
                            else:
                                out_warranty_serials.append(serial)
                                out_warranty_info.append(info)
                        else:
                            all_serials_info.append((serial, '', '', 0))
                    else:
                        all_serials_info.append((serial, '', '', 0))
                else:
                    all_serials_info.append((serial, '', '', 0))
            
            # 写入汇总表
            ws_summary.append(['汇总信息'])
            ws_summary.append(['总查询数量', '查询成功数量', '查询失败数量', '在保机器数量', '过保机器数量'])
            ws_summary.append([
                total_queries, 
                success_queries, 
                failed_queries, 
                len(in_warranty_serials), 
                len(out_warranty_serials)
            ])
            
            # 添加空行
            ws_summary.append([])
            
            # 写入三组数据的表头
            ws_summary.append([
                '在保机器序列号', '维保开始时间', '维保到期时间', '剩余天数',
                '过保机器序列号', '维保开始时间', '维保到期时间', '过保天数',
                '所有机器序列号', '维保开始时间', '维保到期时间', '剩余/过保天数'
            ])
            
            # 获取最大长度并补齐数据
            max_length = max(len(in_warranty_info), len(out_warranty_info), len(all_serials_info))
            in_warranty_padded = in_warranty_info + [('', '', '', '')] * (max_length - len(in_warranty_info))
            out_warranty_padded = out_warranty_info + [('', '', '', '')] * (max_length - len(out_warranty_info))
            all_serials_padded = all_serials_info
            
            # 并排写入三组数据
            for in_info, out_info, all_info in zip(in_warranty_padded, out_warranty_padded, all_serials_padded):
                row = []
                # 添加在保信息
                row.extend(in_info)
                # 添加过保信息
                row.extend(out_info)
                # 添加所有序列号信息
                row.extend(all_info)
                ws_summary.append(row)
            
            # 写入详细信息表头
            ws_detail.append(['序列号', '查询状态', '在保服务数', '过保服务数', '总服务数', 
                             '服务名称', '开始时间', '结束时间', '剩余天数', '保修状态'])
            
            # 按输入顺序写入详细数据
            for serial in current_serials:
                if serial in self.query_results:
                    result = self.query_results[serial]
                    if result.success and result.data['statusCode'] == 200:
                        detail_info = result.data['data'].get('detailinfo', {})
                        service_info = []
                        
                        for service_type in ['warranty', 'onsite', 'other']:
                            if service_type in detail_info:
                                for item in detail_info[service_type]:
                                    remaining_days = int(item.get('DateDifference', 0))
                                    warranty_status = "在保" if remaining_days > 0 else "已过保"
                                    service_info.append({
                                        'name': item.get('ServiceProductName', '未知'),
                                        'start': item.get('StartDate', '未知'),
                                        'end': item.get('EndDate', '未知'),
                                        'days': remaining_days,
                                        'status': warranty_status
                                    })
                        
                        if service_info:
                            for idx, service in enumerate(service_info):
                                if idx == 0:
                                    ws_detail.append([
                                        result.serial,
                                        '成功',
                                        result.valid_services,
                                        result.expired_services,
                                        result.total_services,
                                        service['name'],
                                        service['start'],
                                        service['end'],
                                        service['days'],
                                        service['status']
                                    ])
                                else:
                                    ws_detail.append([
                                        '',
                                        '',
                                        '',
                                        '',
                                        '',
                                        service['name'],
                                        service['start'],
                                        service['end'],
                                        service['days'],
                                        service['status']
                                    ])
                        else:
                            ws_detail.append([
                                result.serial,
                                '成功',
                                0, 0, 0,
                                '无服务信息',
                                '', '', '', ''
                            ])
                    else:
                        ws_detail.append([
                            serial,
                            '失败',
                            0, 0, 0,
                            '查询失败',
                            '', '', '', ''
                        ])
                else:
                    ws_detail.append([
                        serial,
                        '未查询',
                        0, 0, 0,
                        '未查询',
                        '', '', '', ''
                    ])
            
            # 调整列宽
            for ws in [ws_summary, ws_detail]:
                for column in ws.columns:
                    max_length = 0
                    column = list(column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column[0].column_letter].width = adjusted_width
            
            # 保存文件
            wb.save(filename)
            messagebox.showinfo("成功", f"结果已导出到：\n{filename}")
            os.startfile(os.path.abspath("导出结果"))
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def __del__(self):
        # 确保程序退出时关闭线程池
        if hasattr(self, 'executor'):
            self.executor.shutdown(wait=False)

    def create_main_frame(self):
        """创建主框架"""
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def create_input_area(self):
        """创建输入区域"""
        input_frame = ttk.LabelFrame(self.main_frame, text="序列号输入", padding="5")
        input_frame.grid(row=0, column=0, pady=10, sticky=tk.W+tk.E)
        
        # 添加说明文本
        ttk.Label(input_frame, 
                 text="请输入机器序列号（支持多个序列号，可以直接从Excel粘贴，每行一个或用逗号、空格分隔）：").grid(
                     row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        # 使用Text替代Entry，支持多行输入
        self.serial_text = tk.Text(input_frame, width=80, height=5)
        self.serial_text.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 按钮区域
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=5)
        
        ttk.Button(button_frame, text="查询", command=self.check_warranty).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清除输入", command=self.clear_input).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清除所有", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出CSV", command=self.export_to_csv).pack(side=tk.LEFT, padx=5)

    def create_result_area(self):
        """创建结果显示区域"""
        result_frame = ttk.LabelFrame(self.main_frame, text="查询结果", padding="5")
        result_frame.grid(row=1, column=0, sticky=tk.W+tk.E+tk.N+tk.S)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, width=100, height=35)
        self.result_text.pack(fill=tk.BOTH, expand=True)

    def setup_grid_weights(self):
        """设置网格权重"""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)

    def validate_query(self, serials):
        """验证查询参数"""
        if not serials:
            messagebox.showwarning("警告", "请输入至少一个有效的机器序列号！")
            return False
        
        if not messagebox.askyesno("确认", f"将要查询 {len(serials)} 个序列号，是否继续？"):
            return False
        
        return True

    def prepare_query(self, total_serials):
        """准备查询"""
        if self.result_text.get("1.0", tk.END).strip():
            self.result_text.insert(tk.END, "\n" + "="*80 + "\n")
            self.result_text.insert(tk.END, f"新查询开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            self.result_text.insert(tk.END, "="*80 + "\n")
        
        # 禁用按钮
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Button):
                child.configure(state='disabled')

    def create_progress_bar(self, total_serials):
        """创建进度条"""
        progress = ttk.Progressbar(self.root, mode='determinate', maximum=total_serials)
        progress.grid(row=1, column=0, sticky='ew', padx=10)
        return progress

    def cleanup_query(self, progress):
        """清理查询"""
        # 恢复按钮状态
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Button):
                child.configure(state='normal')
        
        # 移除进度条
        progress.destroy()

def main():
    root = tk.Tk()
    app = WarrantyCheckerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 