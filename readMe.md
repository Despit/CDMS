## 需求

### 一、形式

1. 网页/应用

### 二、功能

1. [X] 校验数据录入
2. [X] 历史数据读取-在此基础上修改
3. [X] 报告导出（beta）
4. [X] 全生命周期到期提醒
5. [X] 检定周期提醒管理

### todo

1. 上云
2. 流程审批
3. 网页端

### 运行模式

- 默认使用 MySQL 数据库；设置环境变量 `LOCAL_MODE=1` 可切换为本地 Excel 模式。
- 数据库连接参数通过 `MYSQL_HOST`、`MYSQL_PORT`、`MYSQL_USER`、`MYSQL_PASSWORD` 和 `MYSQL_DATABASE` 配置。
