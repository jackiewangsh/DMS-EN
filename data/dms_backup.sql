-- =====================================================
-- DMS 文档管理系统 - 数据库表结构
-- 生成时间: 2026-06-04
-- 数据库: MariaDB
-- 字符集: utf8mb4
-- =====================================================

-- 设置字符集
SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- =====================================================
-- 1. 用户表
-- =====================================================
DROP TABLE IF EXISTS users;
CREATE TABLE users (
    id VARCHAR(255) PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    password TEXT DEFAULT '',
    email TEXT,
    department TEXT,
    cdsid TEXT,
    status TEXT DEFAULT '在职',
    role TEXT DEFAULT '普通用户',
    created_at TEXT,
    user_theme TEXT DEFAULT 'tech-blue',
    must_change_password INTEGER DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 2. 部门表
-- =====================================================
DROP TABLE IF EXISTS departments;
CREATE TABLE departments (
    id VARCHAR(255) PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    sort_order INTEGER DEFAULT 0,
    created_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 3. 文件类型表（通用，全局类型）
-- =====================================================
DROP TABLE IF EXISTS file_types;
CREATE TABLE file_types (
    id VARCHAR(255) PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    sort_order INTEGER DEFAULT 0,
    created_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 4. 按部门的文件类型表（每个部门独立的文件类型）
-- =====================================================
DROP TABLE IF EXISTS dept_file_types;
CREATE TABLE dept_file_types (
    id VARCHAR(255) PRIMARY KEY,
    department TEXT NOT NULL,
    name TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0,
    created_at TEXT,
    UNIQUE(department, name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 5. 文件表
-- =====================================================
DROP TABLE IF EXISTS files;
CREATE TABLE files (
    id VARCHAR(255) PRIMARY KEY,
    filename TEXT NOT NULL,
    original_name TEXT NOT NULL,
    department TEXT,
    file_type TEXT,
    file_path TEXT NOT NULL,
    file_size INTEGER,
    status TEXT DEFAULT '已发布',
    uploader_id VARCHAR(255),
    uploader_name TEXT,
    created_at TEXT,
    updated_at TEXT,
    is_published INTEGER DEFAULT 1,
    obsolete_reason TEXT,
    FOREIGN KEY(uploader_id) REFERENCES users(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 6. 审批流程表
-- =====================================================
DROP TABLE IF EXISTS workflows;
CREATE TABLE workflows (
    id VARCHAR(255) PRIMARY KEY,
    serial_no TEXT UNIQUE NOT NULL,
    title TEXT NOT NULL,
    initiator_id VARCHAR(255),
    initiator_name TEXT,
    department TEXT,
    current_node TEXT,
    status TEXT DEFAULT '进行中',
    created_at TEXT,
    updated_at TEXT,
    is_closed INTEGER DEFAULT 0,
    closed_at TEXT,
    archive_path TEXT,
    FOREIGN KEY(initiator_id) REFERENCES users(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 7. 审批节点表
-- =====================================================
DROP TABLE IF EXISTS workflow_nodes;
CREATE TABLE workflow_nodes (
    id VARCHAR(255) PRIMARY KEY,
    workflow_id VARCHAR(255) NOT NULL,
    node_name TEXT NOT NULL,
    node_order INTEGER,
    is_required INTEGER DEFAULT 1,
    can_multi INTEGER DEFAULT 0,
    approver_ids TEXT,
    approver_names TEXT,
    waiting_approvers TEXT,
    FOREIGN KEY(workflow_id) REFERENCES workflows(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 8. 审批记录表
-- =====================================================
DROP TABLE IF EXISTS workflow_records;
CREATE TABLE workflow_records (
    id VARCHAR(255) PRIMARY KEY,
    workflow_id VARCHAR(255) NOT NULL,
    node_name TEXT,
    approver_id VARCHAR(255),
    approver_name TEXT,
    action TEXT,
    comment TEXT,
    created_at TEXT,
    FOREIGN KEY(workflow_id) REFERENCES workflows(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 9. 流程附件表
-- =====================================================
DROP TABLE IF EXISTS workflow_files;
CREATE TABLE workflow_files (
    id VARCHAR(255) PRIMARY KEY,
    workflow_id VARCHAR(255) NOT NULL,
    file_id VARCHAR(255),
    filename TEXT,
    original_name TEXT,
    department TEXT,
    file_type TEXT,
    file_path TEXT,
    file_size INTEGER,
    created_at TEXT,
    is_published INTEGER DEFAULT 0,
    is_deleted INTEGER DEFAULT 0,
    FOREIGN KEY(workflow_id) REFERENCES workflows(id),
    FOREIGN KEY(file_id) REFERENCES files(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 10. 操作日志表
-- =====================================================
DROP TABLE IF EXISTS operation_logs;
CREATE TABLE operation_logs (
    id VARCHAR(255) PRIMARY KEY,
    user_id TEXT,
    username TEXT,
    action TEXT,
    detail TEXT,
    ip_address TEXT,
    created_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 11. 系统公告表
-- =====================================================
DROP TABLE IF EXISTS announcements;
CREATE TABLE announcements (
    id VARCHAR(255) PRIMARY KEY,
    content TEXT,
    author_id TEXT,
    created_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 12. 主页富文本内容表
-- =====================================================
DROP TABLE IF EXISTS dashboard_richtext;
CREATE TABLE dashboard_richtext (
    id INT AUTO_INCREMENT PRIMARY KEY,
    content TEXT,
    updated_by VARCHAR(100),
    updated_at VARCHAR(50)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 13. 作废文件记录表
-- =====================================================
DROP TABLE IF EXISTS obsolete_files;
CREATE TABLE obsolete_files (
    id VARCHAR(255) PRIMARY KEY,
    file_no TEXT,
    filename TEXT,
    original_name TEXT,
    department TEXT,
    file_type TEXT,
    file_path TEXT,
    file_size INTEGER DEFAULT 0,
    uploader_name TEXT,
    obsolete_by TEXT,
    obsolete_at TEXT,
    reason TEXT,
    note TEXT,
    restored_at TEXT,
    restored_by TEXT,
    created_at TEXT,
    updated_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 14. 备份记录表
-- =====================================================
DROP TABLE IF EXISTS backup_records;
CREATE TABLE backup_records (
    id VARCHAR(255) PRIMARY KEY,
    backup_path TEXT,
    backup_type TEXT,
    operator_id TEXT,
    operator_name TEXT,
    created_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 15. 通知表
-- =====================================================
DROP TABLE IF EXISTS notifications;
CREATE TABLE notifications (
    id VARCHAR(255) PRIMARY KEY,
    user_id VARCHAR(255) NOT NULL,
    type TEXT,
    title TEXT,
    content TEXT,
    workflow_id VARCHAR(255),
    serial_no TEXT,
    is_read INTEGER DEFAULT 0,
    created_at TEXT,
    FOREIGN KEY(user_id) REFERENCES users(id),
    FOREIGN KEY(workflow_id) REFERENCES workflows(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 16. 系统配置表
-- =====================================================
DROP TABLE IF EXISTS config;
CREATE TABLE config (
    `key` VARCHAR(255) PRIMARY KEY,
    value TEXT,
    updated_at TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 17. 文件分享表
-- =====================================================
DROP TABLE IF EXISTS file_shares;
CREATE TABLE file_shares (
    id INT AUTO_INCREMENT PRIMARY KEY,
    file_id VARCHAR(255) NOT NULL,
    token VARCHAR(255) UNIQUE NOT NULL,
    created_by VARCHAR(100),
    expires_at VARCHAR(50),
    created_at VARCHAR(50) DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- =====================================================
-- 创建索引
-- =====================================================

-- 用户表索引
CREATE INDEX idx_users_username ON users(username(50));
CREATE INDEX idx_users_department ON users(department(50));

-- 文件表索引
CREATE INDEX idx_files_filename ON files(filename(100));
CREATE INDEX idx_files_department ON files(department(50));
CREATE INDEX idx_files_file_type ON files(file_type(50));
CREATE INDEX idx_files_status ON files(status(20));
CREATE INDEX idx_files_uploader_id ON files(uploader_id);
CREATE INDEX idx_files_created_at ON files(created_at);

-- 审批流程表索引
CREATE INDEX idx_workflows_serial_no ON workflows(serial_no);
CREATE INDEX idx_workflows_status ON workflows(status(20));
CREATE INDEX idx_workflows_department ON workflows(department(50));
CREATE INDEX idx_workflows_initiator_id ON workflows(initiator_id);
CREATE INDEX idx_workflows_created_at ON workflows(created_at);

-- 审批节点表索引
CREATE INDEX idx_workflow_nodes_workflow_id ON workflow_nodes(workflow_id);

-- 审批记录表索引
CREATE INDEX idx_workflow_records_workflow_id ON workflow_records(workflow_id);
CREATE INDEX idx_workflow_records_approver_id ON workflow_records(approver_id);

-- 流程附件表索引
CREATE INDEX idx_workflow_files_workflow_id ON workflow_files(workflow_id);
CREATE INDEX idx_workflow_files_file_id ON workflow_files(file_id);

-- 操作日志表索引
CREATE INDEX idx_operation_logs_user_id ON operation_logs(user_id);
CREATE INDEX idx_operation_logs_created_at ON operation_logs(created_at);
CREATE INDEX idx_operation_logs_action ON operation_logs(action(50));

-- 通知表索引
CREATE INDEX idx_notifications_user_id ON notifications(user_id);
CREATE INDEX idx_notifications_is_read ON notifications(is_read);
CREATE INDEX idx_notifications_created_at ON notifications(created_at);

-- 文件分享表索引
CREATE INDEX idx_file_shares_token ON file_shares(token);
CREATE INDEX idx_file_shares_file_id ON file_shares(file_id);

-- =====================================================
-- 初始化默认数据
-- =====================================================

-- 插入默认管理员账户
INSERT INTO users (id, username, password_hash, password, email, department, cdsid, status, role, created_at)
VALUES (
    UUID(),
    'admin',
    '$pbkdf2-sha256$29000$3tFbZR5hxiR8zGLJA9XxOA$44HX60LcNYtMlaztAIM97quusR0clsP0uWBsapJjN8I',
    '',
    'admin@company.com',
    '综合部',
    'admin',
    '在职',
    '系统管理员',
    NOW()
) ON DUPLICATE KEY UPDATE username=username;

-- 插入默认部门
INSERT IGNORE INTO departments (id, name, sort_order, created_at) VALUES
    (UUID(), '综合部文件', 1, NOW()),
    (UUID(), '公司文件', 2, NOW()),
    (UUID(), '财务部', 3, NOW()),
    (UUID(), '研发部', 4, NOW()),
    (UUID(), '市场部', 5, NOW()),
    (UUID(), '人力资源部', 6, NOW()),
    (UUID(), '质量管理部', 7, NOW()),
    (UUID(), 'IT部', 8, NOW()),
    (UUID(), '生产部', 9, NOW());

-- 插入默认文件类型
INSERT IGNORE INTO file_types (id, name, sort_order, created_at) VALUES
    (UUID(), '手册', 1, NOW()),
    (UUID(), '程序文件', 2, NOW()),
    (UUID(), '作业指导书', 3, NOW()),
    (UUID(), '表单', 4, NOW()),
    (UUID(), '外来文件', 5, NOW()),
    (UUID(), '记录', 6, NOW()),
    (UUID(), '其他', 7, NOW());

-- 插入默认系统配置
INSERT IGNORE INTO config (`key`, `value`, updated_at) VALUES
    ('system_version', 'V1.0.0', NOW()),
    ('scheduler_enabled', 'false', NOW());

SET FOREIGN_KEY_CHECKS = 1;
