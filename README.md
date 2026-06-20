# 墨韩表格工具箱

墨韩表格工具箱，基于 `PySide6 + pandas`。

## 运行

### 环境要求

- Debian 13 (Trixie) 或更新版本
- Python >= 3.12

### 安装系统依赖

```bash
sudo apt install python3-pyside6.qtcore python3-pyside6.qtgui python3-pyside6.qtwidgets \
  python3-pandas python3-openpyxl python3-qtpy
```

### 安装与运行

```bash
# 克隆仓库
git clone <repo-url>
cd mohanxlsx

# 安装到系统（开发模式）
pip install --break-system-packages -e .

# 启动应用
mohanxlsx

# 或者直接模块运行
python -m mohanxlsx
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

## Debian 打包

使用 `gbp buildpackage` 构建 `.deb` 安装包。

### 分支结构

| 分支 | 用途 |
|------|------|
| `upstream/latest` | 上游源码 |
| `debian/latest` | Debian 打包（含 `debian/` 目录） |
| `pristine-tar` | 原始 tarball 元数据 |

### 构建

推荐使用 **sbuild** 在干净 chroot 环境中构建，确保依赖版本约束兼容目标 Debian 版本。

首先创建 sbuild chroot（使用 mmdebstrap）：

```bash
mkdir -p ~/.cache/sbuild
mmdebstrap --skip=output/dev --variant=buildd --include=eatmydata \
  unstable ~/.cache/sbuild/unstable-amd64.tar.zst \
  http://ftp.cn.debian.org/debian \
  --aptopt='Acquire::http { Proxy "http://127.0.0.1:3142"; }'
```

然后构建：

```bash
gbp buildpackage --git-builder=sbuild --git-arch=amd64 -d unstable
```

> [!IMPORTANT]
> 直接在宿主机上跑不带 sbuild 的构建会导致生成的 `.deb` 依赖被锁死在构建环境的精确版本上，换一个 Debian 版本就可能装不上。务必用 sbuild 做隔离构建。

### 依赖说明

`debian/control` 中核心依赖通过 `${python3:Depends}` 和 `pyproject.toml` 的 `>=` 约束联合生成。`pyproject.toml` 中避免使用 `==` 钉子版本，否则 `dh_python3` 会翻译出无法跨版本满足的精确依赖。

### 安装

```bash
sudo dpkg -i python3-mohanxlsx_*.deb
# 如有缺失依赖
sudo apt --fix-broken install
```
