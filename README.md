# 墨韩表格工具箱

墨韩表格工具箱，基于 `PySide6 + pandas`。

## 运行

### 环境要求

- Python >= 3.12
- PDM（Python 包管理器）

### 安装与运行

```bash
# 克隆仓库
git clone <repo-url>
cd mohanxlsx

# 安装依赖
pdm install

# 启动应用
pdm run mohanxlsx
```

### 其他运行方式

```bash
# 直接用 Python 运行（需要先激活虚拟环境或已安装依赖）
eval $(pdm venv activate)
python -m mohanxlsx

# 或者 pdm run 直接执行模块
pdm run python -m mohanxlsx
```

## 功能

- 打开 Excel / CSV
- 配置保存与另存为
- 设置标题行
- 列重命名与列开关
- 按列拆分导出多个 Excel
- 多文件合并
- 去重、空行清理、排序
- 统一导出格式：`xlsx / csv / xlsm`
