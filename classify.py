from openai import OpenAI
import win32com.client
import time
import logging
import os
from datetime import datetime, timedelta
from typing import List, Optional
import concurrent.futures
from threading import Semaphore  # 直接使用threading的Semaphore

# 全局配置参数
MAX_WORKERS = min(5, os.cpu_count() or 1)  # 最大并发线程数
API_REQUEST_DELAY = 1.0  # API请求间隔（秒）
BATCH_SIZE = 10  # 每批处理邮件数


class DeepseekAPIWrapper:
    """Deepseek API调用工具（修复信号量实现）"""

    def __init__(self, api_key: str):
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://api.deepseek.com"
        )
        self.request_semaphore = Semaphore(MAX_WORKERS)
        self.last_request_time = 0

    def _throttle_request(self):
        """控制API请求频率，避免限流"""
        current_time = time.time()
        elapsed = current_time - self.last_request_time
        if elapsed < API_REQUEST_DELAY:
            time.sleep(API_REQUEST_DELAY - elapsed)
        self.last_request_time = time.time()

    def chat(self, messages: list, model: str = "deepseek-chat") -> Optional[str]:
        with self.request_semaphore:  # 控制并发数
            self._throttle_request()
            try:
                response = self.client.chat.completions.create(
                    model=model,
                    messages=messages
                )
                return response.choices[0].message.content
            except Exception as e:
                logger.error(f"API调用出错: {e}")
                return None


class OutlookManager:
    """Outlook邮件管理工具"""

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)  # 收件箱（固定索引）
        self.inbox.GetExplorer().CurrentFolder = self.inbox  # 刷新收件箱
        logger.info(f"已连接到Outlook，收件箱: {self.inbox.Name}")

    def get_total_email_count(self) -> int:
        """获取收件箱总邮件数"""
        return self.inbox.Items.Count

    def _get_time_filtered_emails(self, time_range: str) -> List:
        """根据时间范围筛选邮件"""
        now = datetime.now()
        start_time = None

        if time_range == "今天":
            start_time = now.replace(hour=0, minute=0, second=0, microsecond=0)
        elif time_range == "本工作日":
            weekday = now.weekday()
            days_back = weekday if weekday < 5 else 0
            start_time = (now - timedelta(days=days_back)).replace(hour=0, minute=0, second=0, microsecond=0)
        elif time_range == "本周":
            days_back = now.weekday()
            start_time = (now - timedelta(days=days_back)).replace(hour=0, minute=0, second=0, microsecond=0)
        elif time_range == "本月":
            start_time = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

        messages = self.inbox.Items
        messages.Sort("[ReceivedTime]", True)
        filtered = []
        for msg in messages:
            try:
                received_time = msg.ReceivedTime
                if hasattr(received_time, 'year'):
                    received_dt = datetime(
                        received_time.year, received_time.month, received_time.day,
                        received_time.hour, received_time.minute, received_time.second
                    )
                else:
                    received_dt = received_time
                if received_dt >= start_time:
                    filtered.append(msg)
                else:
                    break
            except Exception as e:
                logger.warning(f"跳过异常邮件: {e}")
                continue
        logger.info(f"时间范围 '{time_range}' 筛选出 {len(filtered)} 封邮件")
        return filtered

    def get_emails_by_condition(self, filter_type: str, time_range: str = None, count: int = None) -> List:
        """根据条件获取邮件"""
        if filter_type == "time":
            return self._get_time_filtered_emails(time_range)
        else:
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)
            return list(messages)[:count] if count else []

    def create_category_folder(self, category: str) -> bool:
        """创建分类文件夹"""
        try:
            for folder in self.inbox.Folders:
                if folder.Name == category:
                    return True
            self.inbox.Folders.Add(category)
            logger.info(f"创建文件夹: {category}")
            return True
        except Exception as e:
            logger.error(f"创建文件夹失败: {e}")
            return False

    def move_email_to_category(self, email, category: str) -> bool:
        """移动邮件到分类文件夹"""
        if not self.create_category_folder(category):
            return False
        try:
            target_folder = next((f for f in self.inbox.Folders if f.Name == category), None)
            if target_folder:
                email.Move(target_folder)
                logger.info(f"邮件 '{email.Subject}' 移动到 {category}")
                return True
            logger.error(f"未找到文件夹: {category}")
            return False
        except Exception as e:
            logger.error(f"移动邮件失败: {e}")
            return False


def analyze_email_content(api_wrapper: DeepseekAPIWrapper, email_id: int, email_content: str, categories: list) -> dict:
    """并行分析邮件内容"""
    categories_str = ", ".join(categories)
    messages = [
        {"role": "system", "content": f"仅返回以下类别之一：{categories_str}"},
        {"role": "user", "content": f"分类邮件：\n{email_content[:4000]}"}
    ]
    response = api_wrapper.chat(messages)
    category = response.strip() if response and response.strip() in categories else "未分类"
    return {"email_id": email_id, "category": category}


def process_batch_parallel(api_wrapper: DeepseekAPIWrapper, batch_emails: List, categories: list) -> List[dict]:
    """并行处理一批邮件"""
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(
                analyze_email_content,
                api_wrapper,
                i,
                email.Body if hasattr(email, 'Body') else "",
                categories
            ) for i, email in enumerate(batch_emails)
        ]
        results = []
        for future in concurrent.futures.as_completed(futures):
            try:
                results.append(future.result())
            except Exception as e:
                logger.error(f"分析失败: {e}")
        return sorted(results, key=lambda x: x["email_id"])


def main():
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("email_analysis.log", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    global logger
    logger = logging.getLogger(__name__)

    # 配置参数
    DEEPSEEK_API_KEY = "Your API Key"  # 替换为实际API密钥
    CATEGORIES = ["垃圾邮件", "学校", "游戏", "公司","学习"]

    try:
        outlook_manager = OutlookManager()
        total_count = outlook_manager.get_total_email_count()
        print(f"\n当前收件箱总邮件数: {total_count} 封")

        # 选择筛选方式
        filter_options = [
            "1. 今天（处理今天所有邮件）",
            "2. 本工作日（处理本周一至当前所有邮件）",
            "3. 本周（处理本周所有邮件）",
            "4. 本月（处理本月所有邮件）",
            "5. 指定数目（处理最新的N封邮件，不限制时间）"
        ]
        print("\n请选择邮件筛选方式：")
        for option in filter_options:
            print(option)

        # 处理筛选方式输入
        while True:
            try:
                filter_choice = int(input("请输入选项(1-5)：").strip())
                if 1 <= filter_choice <= 5:
                    break
                print("请输入1-5之间的数字")
            except ValueError:
                print("输入错误，请输入数字")

        # 解析筛选方式
        filter_type = "time" if filter_choice <= 4 else "count"
        time_range = ["今天", "本工作日", "本周", "本月"][filter_choice - 1] if filter_choice <= 4 else ""

        # 仅当选择第5项时需要输入数量
        num_emails = None
        if filter_choice == 5:
            while True:
                try:
                    num_emails = int(input("请输入需要处理的邮件数量：").strip())
                    if num_emails > 0:
                        break
                    print("请输入正整数")
                except ValueError:
                    print("输入错误，请输入数字")

        # 获取筛选后的邮件
        emails = outlook_manager.get_emails_by_condition(
            filter_type=filter_type,
            time_range=time_range,
            count=num_emails
        )
        if not emails:
            print(f"\n没有符合条件的邮件")
            return

        # 显示处理信息
        if filter_choice <= 4:
            print(f"\n将处理{time_range}的所有 {len(emails)} 封邮件")
        else:
            print(f"\n将处理最新的 {len(emails)} 封邮件")

        # 初始化API并处理邮件
        api_wrapper = DeepseekAPIWrapper(api_key=DEEPSEEK_API_KEY)
        total_batches = (len(emails) + BATCH_SIZE - 1) // BATCH_SIZE
        print(f"将分{total_batches}批处理")

        # 分批处理
        for batch_idx in range(total_batches):
            print(f"\n===== 第{batch_idx + 1}/{total_batches}批 =====")
            batch_start = batch_idx * BATCH_SIZE
            batch_end = min(batch_start + BATCH_SIZE, len(emails))
            batch_emails = emails[batch_start:batch_end]

            # 并行分析
            start_time = time.time()
            batch_results = process_batch_parallel(api_wrapper, batch_emails, CATEGORIES)
            print(f"分析耗时: {time.time() - start_time:.2f}秒")

            # 移动邮件
            for i, result in enumerate(batch_results):
                email = batch_emails[i]
                category = result["category"]
                email_index = batch_start + i + 1
                short_subject = email.Subject[:30] + "..." if len(email.Subject) > 30 else email.Subject
                print(f"邮件 {email_index}/{len(emails)}: 主题《{short_subject}》分类为【{category}】")
                if category != "未分类":
                    outlook_manager.move_email_to_category(email, category)

        print(f"\n✅ 所有{len(emails)}封邮件处理完成！")
        logger.info("所有邮件处理完成")

    except Exception as e:
        logger.error(f"程序错误: {e}")
        print(f"发生错误: {e}")


if __name__ == "__main__":
    main()