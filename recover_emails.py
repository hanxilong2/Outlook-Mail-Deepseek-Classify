import win32com.client
import logging
from typing import List

# 配置参数
SHOW_PROGRESS_INTERVAL = 10  # 每处理10封邮件显示一次进度


def setup_logger():
    """配置日志记录"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("move_emails.log", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


def get_inbox_subfolders(inbox) -> List:
    """获取收件箱下的所有子文件夹"""
    subfolders = []
    try:
        for folder in inbox.Folders:
            subfolders.append(folder)
        logger.info(f"找到 {len(subfolders)} 个收件箱子文件夹")
    except Exception as e:
        logger.error(f"获取子文件夹失败: {e}")
    return subfolders


def move_subfolder_emails_to_inbox(inbox):
    subfolders = get_inbox_subfolders(inbox)
    if not subfolders:
        logger.info("没有找到子文件夹，无需处理")
        return

    total_moved = 0
    total_folders = len(subfolders)

    print(f"\n开始处理 {total_folders} 个子文件夹")

    for i, folder in enumerate(subfolders, 1):
        folder_name = folder.Name
        try:
            messages = folder.Items
            folder_email_count = messages.Count

            if folder_email_count == 0:
                logger.info(f"文件夹 '{folder_name}' 为空，跳过")
                continue

            logger.info(f"开始处理文件夹 '{folder_name}' ({i}/{total_folders})，包含 {folder_email_count} 封邮件")
            moved_count = 0

            # 转换为列表避免迭代过程中修改集合
            for j, msg in enumerate(list(messages)):
                try:
                    msg.Move(inbox)
                    moved_count += 1
                    total_moved += 1

                    # 显示进度
                    if j % SHOW_PROGRESS_INTERVAL == 0 or j == folder_email_count - 1:
                        print(f"\r正在处理文件夹 '{folder_name}': {j + 1}/{folder_email_count} 封邮件已移动", end="")

                except Exception as e:
                    logger.warning(f"移动邮件失败（主题: {getattr(msg, 'Subject', '未知')}）: {e}")
                    continue

            print(f"\r文件夹 '{folder_name}' 处理完成，成功移动 {moved_count}/{folder_email_count} 封邮件")
            logger.info(f"文件夹 '{folder_name}' 处理完成，成功移动 {moved_count}/{folder_email_count} 封邮件")

        except Exception as e:
            logger.error(f"处理文件夹 '{folder_name}' 时出错: {e}")
            print(f"\n处理文件夹 '{folder_name}' 时出错: {e}")
            continue

    logger.info(f"所有子文件夹处理完成，总计移动 {total_moved} 封邮件到收件箱")
    print(f"\n\n操作完成，总计移动 {total_moved} 封邮件到收件箱")


def main():
    global logger
    logger = setup_logger()

    try:
        # 连接到Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 代表收件箱
        logger.info(f"已连接到Outlook，收件箱: {inbox.Name}")

        # 执行移动操作
        move_subfolder_emails_to_inbox(inbox)

    except Exception as e:
        logger.error(f"程序执行失败: {e}")
        print(f"发生错误: {e}")


if __name__ == "__main__":
    print("===== 收件箱子文件夹邮件迁移工具 =====")
    print("功能：将收件箱下所有子文件夹中的邮件移动回收件箱\n")
    main()