# Outlook收件箱子文件夹邮件迁移工具功能文档

## 一、概述
此工具是用Python编写的脚本，借助`win32com.client`库与Windows系统下的Outlook客户端进行交互，其主要功能是将Outlook收件箱下所有子文件夹中的邮件移动回收件箱。同时，工具还具备详细的日志记录功能，方便用户追踪操作过程和排查问题。

## 二、功能特性
1. **自动检测子文件夹**：自动获取Outlook收件箱下的所有子文件夹。
2. **邮件移动**：将每个子文件夹中的邮件逐一移动回收件箱。
3. **进度显示**：在处理邮件时，每处理10封邮件显示一次进度，让用户了解处理情况。
4. **日志记录**：将操作过程中的关键信息记录到`move_emails.log`文件中，同时在控制台输出日志。
5. **错误处理**：在获取子文件夹、移动邮件等操作中出现异常时，会记录错误信息并继续处理其他任务。

## 三、环境要求
- **操作系统**：Windows系统，因为依赖`win32com.client`库与Outlook客户端进行交互。
- **软件**：已安装Microsoft Outlook并配置好邮箱账户，确保能正常接收邮件。
- **Python库**：需要安装`pywin32`库，可使用以下命令进行安装：
```bash
pip install pywin32
```

## 四、代码结构及函数说明

### 1. 全局变量
```python
SHOW_PROGRESS_INTERVAL = 10  # 每处理10封邮件显示一次进度
```

### 2. 配置日志记录 - `setup_logger`
```python
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
```
- **功能**：配置日志记录，将日志信息同时输出到`move_emails.log`文件和控制台。
- **返回值**：返回配置好的日志记录器。

### 3. 获取收件箱子文件夹 - `get_inbox_subfolders`
```python
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
```
- **功能**：获取指定收件箱下的所有子文件夹。
- **参数**：
  - `inbox`：Outlook收件箱对象。
- **返回值**：包含所有子文件夹对象的列表。

### 4. 移动子文件夹邮件到收件箱 - `move_subfolder_emails_to_inbox`
```python
def move_subfolder_emails_to_inbox(inbox):
    ...
```
- **功能**：将收件箱下所有子文件夹中的邮件移动回收件箱。
- **参数**：
  - `inbox`：Outlook收件箱对象。
- **详细步骤**：
  1. 调用`get_inbox_subfolders`函数获取所有子文件夹。
  2. 若没有找到子文件夹，记录日志并返回。
  3. 遍历每个子文件夹，检查是否为空，若为空则跳过。
  4. 遍历子文件夹中的每封邮件，将其移动回收件箱，并更新移动计数。
  5. 每处理10封邮件显示一次进度，处理完成后记录日志。
  6. 若移动邮件或处理子文件夹时出现异常，记录错误信息并继续处理下一个子文件夹。

### 5. 主函数 - `main`
```python
def main():
    ...
```
- **功能**：程序的入口点，负责初始化日志记录器、连接到Outlook并执行邮件移动操作。
- **详细步骤**：
  1. 调用`setup_logger`函数配置日志记录。
  2. 连接到Outlook客户端，获取收件箱对象。
  3. 调用`move_subfolder_emails_to_inbox`函数执行邮件移动操作。
  4. 若程序执行过程中出现异常，记录错误信息并在控制台输出。

## 五、使用步骤
1. 确保满足环境要求，安装好所需的Python库。
2. 将上述代码保存为Python文件，例如`recover_demo.py`。
3. 打开命令行工具，进入保存文件的目录。
4. 运行脚本：
```bash
python recover_demo.py
```
5. 程序会自动连接到Outlook，检测收件箱子文件夹并将其中的邮件移动回收件箱，同时在控制台显示处理进度和结果。
6. 操作完成后，可查看`move_emails.log`文件获取详细的操作日志。

## 六、注意事项
- 首次运行脚本时，可能需要授予Python访问Outlook的权限，在弹出的对话框中选择“允许”。
- 若在处理过程中出现异常，可查看`move_emails.log`文件中的错误信息进行排查。
- 处理大量邮件时，可能需要一定的时间，请耐心等待。 