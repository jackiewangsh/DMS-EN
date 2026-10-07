# DMS 新电脑安装详细操作步骤

> 当前服务器环境：Python 3.12.10 / MariaDB 12.2.2 / Windows Server
> DMS 路径：C:\AI\QC\DMS
> 最后更新：2026-05-14

---

## 总览

安装分为 **5 个阶段**，预计耗时 30-60 分钟：

| 阶段 | 内容 | 耗时 |
|------|------|------|
| ① 准备工作 | 在旧电脑上导出数据 | 5-10分钟 |
| ② 安装软件 | Python + MariaDB + LibreOffice | 15-20分钟 |
| ③ 复制文件 | DMS 文件夹 + 上传文件 | 10-30分钟 |
| ④ 导入数据库 | MariaDB 建库 + 导入数据 | 2-5分钟 |
| ⑤ 启动验证 | 运行 DMS 并测试 | 5分钟 |

---

## 阶段 ①：在旧电脑上准备数据（如果是从旧电脑迁移）

> ⚠️ 如果是全新安装（无旧数据），跳过此阶段，直接从阶段 ② 开始。

### 步骤 1：备份数据库

在旧电脑上打开命令提示符（CMD），执行：

    "C:\Program Files\MariaDB 12.2\bin\mysqldump.exe" -u dms_user -pDmsUser123! dms > C:\AI\QC\DMS\data\dms_backup.sql

> 💡 也可以直接双击 DMS 目录下的 `databasebackup.bat`。

### 步骤 2：复制 DMS 文件夹

**需要复制的文件和文件夹**：

    C:\AI\QC\DMS\
    ├── app.py                          ← 必须复制
    ├── _run.py                         ← 必须复制
    ├── watchdog.py                     ← 必须复制
    ├── start_production.bat            ← 必须复制
    ├── start.bat                       ← 必须复制
    ├── start_silent.vbs                ← 必须复制
    ├── ExitDMS-endPythonProcess.bat    ← 必须复制
    ├── databasebackup.bat              ← 必须复制
    ├── requirements.txt                ← 必须复制
    ├── INSTALL_GUIDE.md                ← 可选（文档）
    ├── USER_GUIDE.md                   ← 可选（文档）
    ├── 生产部署说明.md                  ← 可选（文档）
    ├── system\                         ← 必须复制整个文件夹
    │   ├── login_title.txt
    │   ├── rebuild_index.py
    │   ├── static\                     ← Logo、CSS、JS、字体
    │   ├── templates\                  ← 22个HTML页面
    │   └── wheels\                     ← 44个离线安装包（约77MB）
    ├── data\
    │   ├── config.json                 ← 系统配置（SMTP等）
    │   ├── dms_backup.sql              ← 数据库备份（刚导出的）
    │   └── .secret_key                 ← Flask密钥
    └── uploads\                        ← 上传文件（按部门子目录）

**⚠️ 不要复制的文件**：

    C:\AI\QC\DMS\
    ├── venv\                           ← ❌ 绝对不要复制！
    │                                    虚拟环境绑定旧电脑Python路径，复制过来100%无法使用
    └── index_debug.log                 ← 调试日志

> ⚠️ **最重要的一点：不要复制 venv 文件夹！** 新电脑上会自动重新创建。

---

## 阶段 ②：在新电脑上安装软件

### 步骤 1：安装 Python 3.12 或 3.14

1. 打开浏览器访问 https://www.python.org/downloads/
2. 下载 Python 3.12.x（推荐）或 Python 3.14.x
3. 运行安装程序，**务必勾选 "Add Python to PATH"**（底部复选框）
4. 点击 "Install Now" 完成安装
5. 打开新的 CMD 窗口，验证：`python --version`
   应显示 `Python 3.12.x` 或 `Python 3.14.x`

> ⚠️ DMS 的离线安装包（system/wheels/）仅包含 cp312 和 cp314 两个版本的二进制包。其他版本需要联网安装。

### 步骤 2：安装 MariaDB 12.x

1. 打开浏览器访问 https://mariadb.org/download/
2. 选择版本：MariaDB 12.x（推荐）或 10.11+
3. 下载 Windows 64-bit MSI 安装包
4. 运行安装程序：
   - 安装类型：选 "Server only" 或 "Custom"
   - 设置 root 密码：**记住这个密码**（后面要用）
   - 端口：保持默认 3306
   - 字符集：utf8mb4
   - 勾选 "Configure MariaDB Server as Windows Service"
5. 安装完成后验证：`"C:\Program Files\MariaDB 12.2\bin\mysql.exe" -u root -p`
   输入 root 密码，看到 `MariaDB [(none)]>` 即成功。输入 `exit` 退出。

### 步骤 3：创建数据库和用户

打开 CMD，连接 MariaDB，输入 root 密码后依次执行：

    CREATE DATABASE dms CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
    CREATE USER 'dms_user'@'localhost' IDENTIFIED BY 'DmsUser123!';
    GRANT ALL PRIVILEGES ON dms.* TO 'dms_user'@'localhost';
    FLUSH PRIVILEGES;

验证：`SHOW DATABASES;` 和 `SELECT User, Host FROM mysql.user WHERE User = 'dms_user';`

> 💡 如果想修改默认密码 `DmsUser123!`，需要同时修改 app.py 第 99 行。

### 步骤 4：安装 LibreOffice（用于文档预览）

1. 访问 https://www.libreoffice.org/download/download/
2. 下载 Windows x86_64 版本（推荐 7.x+）
3. 安装完成后确认 `C:\Program Files\LibreOffice\program\soffice.exe` 存在

> 💡 不安装也能运行 DMS，但 Office 文件预览会降级为纯文本。

---

## 阶段 ③：复制 DMS 文件夹到新电脑

1. 将 DMS 文件夹（**不含 venv/**）复制到 `C:\AI\QC\DMS\`
2. 确认 app.py、system/wheels/、data/dms_backup.sql 存在
3. 如果 venv 目录被误复制，删除它：`rmdir /s /q C:\AI\QC\DMS\venv`

---

## 阶段 ④：导入数据库

> start_production.bat 首次运行会自动导入，但推荐手动导入以便看到错误信息。

    "C:\Program Files\MariaDB 12.2\bin\mysql.exe" -u root -p dms < C:\AI\QC\DMS\data\dms_backup.sql

验证：

    "C:\Program Files\MariaDB 12.2\bin\mysql.exe" -u dms_user -pDmsUser123! dms -e "SHOW TABLES;"
    "C:\Program Files\MariaDB 12.2\bin\mysql.exe" -u dms_user -pDmsUser123! dms -e "SELECT COUNT(*) AS file_count FROM files;"

应看到 19 个表和文件数量。

---

## 阶段 ⑤：启动与验证

### 步骤 1：首次启动

双击 `C:\AI\QC\DMS\start_production.bat`

脚本会自动执行：
1. 创建虚拟环境 venv/
2. 从 system/wheels/ 安装依赖（离线，无需网络）
3. 验证 16 个依赖包全部就绪
4. 检测数据库配置
5. 启动 Waitress 生产服务器

看到以下输出即成功：

    [DMS] All dependencies verified OK
    [DMS] Starting production server...
      URL: http://localhost:5000
      Admin: admin / admin123

### 步骤 2：浏览器访问

1. 打开浏览器访问 http://localhost:5000
2. 默认账号：用户名 `admin`，密码 `admin123`
3. **登录后立即修改管理员密码**

### 步骤 3：功能检查

| 检查项 | 操作 | 预期结果 |
|--------|------|---------|
| 文件库 | 侧边栏展开部门 → 点击文件类型 | 显示文件列表 |
| 文件预览 | 点击 PDF 文件名 | 浏览器内显示 PDF |
| 文件下载 | 点击下载按钮 | 下载文件（原始文件名） |
| 全文搜索 | 搜索框输入关键词 | 显示搜索结果 |
| 上传文件 | 选择部门/类型 → 上传 | 上传成功 |
| 系统配置 | 侧边栏 → 系统配置 | 正常显示 |

### 步骤 4：重建全文索引

登录管理员 → 系统配置 → 点击 **"重建索引"** → 等待进度条完成。

### 步骤 5：配置防火墙（局域网访问需要）

    netsh advfirewall firewall add rule name="DMS服务器" dir=in action=allow protocol=tcp localport=5000

### 步骤 6：查看本机 IP 并通知同事

    ipconfig

找到 IPv4 地址，同事通过 http://你的IP:5000 访问。

---

## 可选：配置开机自启动

### 方式一：启动文件夹（简单）

Win+R → 输入 `shell:startup` → 回车 → 将 `start_production.bat` 快捷方式拖入

### 方式二：计划任务（推荐，管理员 CMD 执行），运行后win+r, taskschd.msc,看是否有DMS...这两个计划了
    schtasks /Create /TN "DMS_AutoStart" /TR "wscript.exe \"C:\AI\QC\DMS\start_silent.vbs\"" /SC ONLOGON /RL HIGHEST /F
    schtasks /Create /TN "DMS_Watchdog" /TR "C:\AI\QC\DMS\venv\Scripts\python.exe C:\AI\QC\DMS\watchdog.py" /SC MINUTE /MO 1 /RL HIGHEST /F

如果上面不行，运行下面这两个
schtasks /Create /TN "DMS_AutoStart" /TR "\"C:\AI\QC\DMS\venv\Scripts\python.exe\" \"C:\AI\QC\DMS\_run.py\"" /SC ONSTART /RU SYSTEM /F
schtasks /Create /TN "DMS_Watchdog" /TR "\"C:\AI\QC\DMS\venv\Scripts\python.exe\" \"C:\AI\QC\DMS\watchdog.py\"" /SC MINUTE /MO 1 /RU SYSTEM /F

---

## 常见问题排查

### Q1: start_production.bat 报 "venv creation failed"
Python 未安装或未加入 PATH。运行 `python --version` 检查。

### Q2: 依赖安装失败
Python 版本不是 3.12/3.14。安装正确版本后删除 venv 重新运行：
`rmdir /s /q C:\AI\QC\DMS\venv` 然后 `start_production.bat`

### Q3: 数据库连接失败
`sc query MySQL` 检查服务，`net start MySQL` 启动服务。

### Q4: venv 创建失败但目录已存在
`rmdir /s /q C:\AI\QC\DMS\venv` 然后 `start_production.bat`

### Q5: 全文搜索返回空
系统配置 → 重建索引，等待进度条完成。

### Q6: Word/Excel 预览显示纯文本
未安装 LibreOffice，安装后重启 DMS。

### Q7: 同事无法访问
确认防火墙开放 5000 端口，确认用服务器 IP（不是 127.0.0.1）。

---

## 安装检查清单

- [ ] Python 3.12 或 3.14 已安装（python --version）
- [ ] MariaDB 服务正在运行（sc query MySQL）
- [ ] dms 数据库已创建（SHOW DATABASES）
- [ ] dms_user 用户已创建并有权限
- [ ] DMS 文件夹已复制（不含 venv/）
- [ ] data/dms_backup.sql 存在
- [ ] uploads/ 已复制
- [ ] start_production.bat 正常启动（无报错）
- [ ] 浏览器能访问 http://localhost:5000
- [ ] admin 账号能登录
- [ ] 文件库显示正常
- [ ] 全文索引已重建
- [ ] 防火墙已开放（如需局域网访问）
- [ ] 已修改管理员默认密码