# data-dictionary2excel

此项目旨在将数据库中的数据字典自动生成并导出为Excel文件，以便于查阅和分享。通过Python脚本实现自动化处理，利用SQLAlchemy连接数据库，并通过pandas库读取表结构信息，最后使用openpyxl库将这些信息格式化为Excel表格输出。

目前支持的数据库：

- MySQL
- GaussDB (待支持)
- Postgres (待支持)

## 功能概述

`gen_datadic.py` 将读取指定数据库中所有表的结构与注释，按表逐项输出为一份结构化的数据字典 Excel 文件，示例输出文件名为：`<db_name>_data_dictionary.xlsx`。

生成内容包括：

- 表名与表注释
- 每个字段的名称、类型（含长度/精度/小数位）、是否可空、是否主键、字段注释
- 基本样式（表头加粗、字段表格带边框）

## 依赖

请确保在运行前安装下列 Python 包：

```bash
# 创建虚拟环境并激活：
# 建议python版本3.10
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip

# 若使用 MySQL，请安装以下包：
pip install pandas sqlalchemy mysql-connector-python openpyxl

# 若使用 GaussDB，请安装以下包：
pip install gaussdb pandas openpyxl
sh install_gaussdb_driver.sh
```

> 注意：如果你使用的是系统的 `mysqlclient` 驱动或其他 DB 驱动，`create_engine` 的连接字符串需要相应修改。


---

## 使用方法

1. 克隆仓库：

```bash
git clone https://github.com/pangpang20/data-dictionary2excel.git
cd data-dictionary2excel
```

2. 编辑 `gen_datadic.py` 中的数据库连接字符串（若不想在文件中硬编码密码，请替换为从环境变量读取或在运行前修改）：

```py
# 原始脚本中的连接字符串示例：
engine = create_engine(f'mysql+mysqlconnector://#username#:#password#@#dbhost#/{db_name}')
# 当前脚本使用（示例）
engine = create_engine(f'mysql+mysqlconnector://cyl:Audaque%40123@localhost/{db_name}')
```

你可以将用户名/密码/host/port 改为适合你环境的值，或在脚本中改用读取环境变量的方式：

```py
import os
user = os.environ.get('DB_USER', 'cyl')
password = os.environ.get('DB_PASSWORD', 'Audaque@123')
host = os.environ.get('DB_HOST', 'localhost')
engine = create_engine(f'mysql+mysqlconnector://{user}:{password}@{host}/{db_name}')
```

3. 运行脚本生成 Excel（示例）：

```bash
python gen_datadic.py my_database_name
python gen_datadic4gaussdb.py my_database_name
```

执行后会在当前目录生成 `my_database_name_data_dictionary.xlsx`。

![alt text](resource/image1.png)

![alt text](resource/image2.png)
