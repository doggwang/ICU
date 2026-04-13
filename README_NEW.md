# ICU 检测报告处理系统 v2.0

重构后的模块化系统，支持配置化扩展，便于适配不同医院的报告格式。

## 项目结构

```
ICU/
├── src/
│   └── icu_report_processor/     # 主包
│       ├── config/               # 配置模块
│       │   ├── __init__.py       # 配置管理类
│       │   └── hospital_config.yaml  # 医院配置文件
│       ├── classifiers/          # 分类器模块
│       │   ├── __init__.py
│       │   └── base.py           # 分类器基类和关键词分类器
│       ├── parsers/              # 解析器模块
│       │   ├── __init__.py
│       │   └── base.py           # 解析器基类和默认解析器
│       ├── exporters/            # 导出器模块
│       │   ├── __init__.py
│       │   └── excel_exporter.py # Excel 导出器
│       ├── __init__.py           # 包入口
│       ├── processor.py          # 主处理器
│       └── pdf_utils.py          # PDF 工具函数
│   └── main.py                   # 命令行入口
├── tests/
│   └── test_processor.py         # 功能测试
├── process_reports_new.py        # 使用示例
└── 检测报告整理/                 # 报告数据目录
```

## 核心改进

### 1. 配置化设计

所有医院特定的规则都提取到 `hospital_config.yaml`：

```yaml
report_types:
  blood_gas:
    name: "血气分析"
    classification_keywords:
      - "(ICU)(POCT)"
      - "血气分析"
    indicator_fields:
      - "pH"
      - "PaO2"

hospitals:
  default:
    patient_info_patterns:
      patient_name: '姓\s*名[:\s]*([^\s\n]+)'
      patient_id: '病\s*历\s*号[:\s]*([^\s\n]+)'
```

### 2. 模块化架构

| 模块 | 职责 | 扩展方式 |
|------|------|----------|
| Classifier | 报告分类 | 继承 BaseClassifier |
| Parser | 数据解析 | 继承 BaseParser |
| Exporter | 数据导出 | 继承 BaseExporter |

### 3. 清晰的接口定义

```python
# 分类器接口
class BaseClassifier(ABC):
    @abstractmethod
    def classify(self, text: str) -> Optional[str]: ...

# 解析器接口
class BaseParser(ABC):
    @abstractmethod
    def parse(self, text: str, filename: str, report_type: str) -> ParseResult: ...
```

## 使用方法

### 命令行方式

```bash
# 基本用法
python src/main.py --input raw --output .

# 整理文件到分类目录
python src/main.py --input raw --output . --organized 检测报告整理

# 指定患者信息
python src/main.py --input raw --output . \
    --patient-name 张三 \
    --patient-id 12345 \
    --patient-gender 男
```

### Python API 方式

```python
from icu_report_processor import create_processor

# 创建处理器
processor = create_processor()

# 处理目录
results = processor.process_directory(
    input_dir=Path("raw"),
    output_dir=Path("output"),
    patient_info={'姓名': '张三', '病历号': '12345'}
)
```

## 适配新医院

当需要适配新医院时，只需：

1. **复制配置文件**
   ```bash
   cp config/hospital_config.yaml config/hospital_X.yaml
   ```

2. **修改正则表达式**
   ```yaml
   hospitals:
     hospital_X:
       patient_info_patterns:
         patient_name: 'Name:\s*(\S+)'  # 英文格式
   ```

3. **调整分类关键词**
   ```yaml
   report_types:
     blood_gas:
       classification_keywords:
         - "ABG"  # 添加新医院的标识
         - "Blood Gas"
   ```

4. **使用新配置**
   ```python
   processor = create_processor(
       config_path="config/hospital_X.yaml",
       hospital_id="hospital_X"
   )
   ```

## 测试

```bash
# 运行功能测试
python tests/test_processor.py
```

## 与原系统的对比

| 特性 | 原系统 | 新系统 |
|------|--------|--------|
| 配置方式 | 硬编码 | YAML 配置文件 |
| 扩展性 | 需修改代码 | 仅需修改配置 |
| 代码复用 | 复制粘贴 | 继承基类 |
| 测试覆盖 | 无 | 有基础测试 |
| 文档 | 注释 | 完整文档 + 类型注解 |

## 下一步建议

1. **添加更多测试**：针对边界情况和异常情况
2. **日志系统**：添加结构化日志记录
3. **错误处理**：增强错误恢复能力
4. **性能优化**：大批量文件处理优化
5. **GUI 界面**：可选的图形界面
