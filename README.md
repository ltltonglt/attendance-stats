# 人员打卡核查工具

成都中兴人员打卡统计与核查工具，用于自动化处理和分析员工打卡数据。

## 功能特性

- 打卡数据导入与解析
- 人员信息管理
- 打卡异常检测与核查
- 经纬度位置验证
- 核查结果导出（Excel格式）

## 项目结构

```
├── attendance_audit_core.py    # 核心业务逻辑
├── attendance_audit_gui.py     # GUI界面
├── 人员信息.xlsx                # 人员信息数据
├── 打卡数据导出/                # 打卡原始数据
├── tests/                      # 测试文件
├── docs/                       # 文档
├── build/                      # 打包构建文件
└── dist/                       # 打包输出
```

## 使用方法

### 运行程序

```bash
python attendance_audit_gui.py
```

### 打包为可执行文件

```bash
pyinstaller 人员打卡核查工具.spec
```

## 技术栈

- Python 3.x
- tkinter (GUI)
- openpyxl (Excel处理)

## 许可证

私有项目