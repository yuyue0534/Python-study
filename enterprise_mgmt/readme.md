**企业内部业务管理系统开发方案**

完整目录结构：
```
enterprise_mgmt/
├── main.py                    # 程序入口（日志+DB初始化+启动）
├── config.py                  # 全局配置
├── requirements.txt
├── database/
│   └── db_manager.py          # 单例连接 + 自动建表（5张表）
├── dal/                       # 数据访问层
│   ├── employee_dal.py        # 员工 CRUD + 分页查询
│   ├── department_dal.py      # 部门 CRUD
│   ├── inventory_dal.py       # 商品/采购/销售 CRUD
│   └── finance_dal.py         # 财务统计查询
├── bll/                       # 业务逻辑层
│   ├── employee_bll.py        # 工号自动生成、手机/薪资校验
│   ├── department_bll.py      # 删除前检查在职员工
│   ├── inventory_bll.py       # 采购入库/销售扣库+联动财务
│   └── finance_bll.py         # 图表数据加工
├── ui/
│   ├── styles.py              # 完整 QSS 商务主题
│   ├── main_window.py         # 主窗口（左导航+顶栏+StackedWidget）
│   ├── modules/
│   │   ├── hrm_widget.py      # 员工列表+分页+搜索+导出Excel
│   │   ├── department_widget.py # QTreeWidget树形架构+员工详情
│   │   ├── psi_widget.py      # 库存预警高亮+采购+销售三Tab
│   │   └── finance_widget.py  # Matplotlib折线图+饼图+KPI卡片
│   └── dialogs/
│       ├── employee_dialog.py  # 员工新增/编辑弹窗
│       └── psi_dialogs.py     # 采购/销售录入弹窗
└── utils/
    └── excel_exporter.py      # 带样式的Excel导出
```

安装运行
```
# 1. 解压后进入目录
cd enterprise_mgmt

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动
python main.py
```

首次运行自动建库建表，预置 5 个部门，SQLite 文件生成在项目根目录 enterprise.db。
