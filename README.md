
- - -

## 项目介绍

> Excel转PDF模板（套打-电子签）

> 说明：类似电子签署，提取Excel数据套到PDF表单模板里面，支持文本和图片的自动填充

## 功能介绍

### 核心功能

- **数据提取**：从Excel文件中提取文本数据和图片
- **模板填充**：将提取的数据填充到PDF表单模板中
- **图片支持**：支持WPS的DISPIMG图片和Office的浮动图片
- **批量处理**：支持批量处理多行数据生成多个PDF文件
- **格式转换**：支持将生成的PDF转换为PNG图片或PPT演示文稿

### 图片提取功能

#### 支持的图片格式

1. **DISPIMG图片（WPS格式）**
   - 适用于WPS Excel中使用`=DISPIMG()`函数嵌入的图片
   - 图片通过ID引用，存储在Excel内部结构中
   - 自动识别`=_xlfn.DISPIMG("ID_xxx",1)`格式

2. **浮动图片（Office格式）**
   - 适用于Microsoft Office Excel中插入的标准浮动图片
   - 图片存储在Excel的xl/media目录中
   - 通过位置关系与单元格关联

#### 自动检测机制

程序会自动判断Excel单元格中的图片类型：
- 优先检测DISPIMG格式图片
- 如果不是DISPIMG，则检测该位置的浮动图片
- 支持WPS和Office两种Excel格式的无缝切换

### 软件截图

### 基础功能 

## 相关技术

> Python版本：3.12+

> 核心库：pandas, openpyxl, PyMuPDF, PIL

> GUI UI使用：tkinter

> 打包程序使用：cx_Freeze

### 目录结构

```
ExcelToPDFTemplate/
├── gui.py                    # GUI界面代码
├── core.py                   # 主要逻辑代码
├── CatchExcelImageTool.py    # 图片提取工具
├── GUIDGen.py               # GUID生成工具
├── setup.py                 # 打包脚本
├── requirements.txt         # 依赖项列表
├── test_image_extraction.py # 图片提取测试脚本
├── README.md               # 项目说明文档
├── presets/                # 预设配置目录
│   ├── 模板-表单.json
│   └── 模板-表单图.json
├── resources/              # 资源目录
│   └── template/
│       └── 模板-表单.pdf
└── build/                  # 构建输出目录
    └── exe.win-amd64-3.12/
        └── ExcelToPDFTemplate.exe
```

### 架构介绍

- **gui.py**: GUI界面代码，提供用户交互界面
- **core.py**: 主要逻辑代码，包含Excel数据处理和PDF生成功能
- **CatchExcelImageTool.py**: 图片提取工具，支持DISPIMG和浮动图片提取
- **GUIDGen.py**: GUID生成工具
- **setup.py**: 用于打包应用程序的脚本
- **requirements.txt**: 列出了所有Python依赖项
- **test_image_extraction.py**: 图片提取功能的测试脚本
- **presets/**: 存储预设配置文件
- **resources/**: 存储模板文件和其他资源

### 图片提取使用示例

#### 1. 在字段映射中配置图片字段

```json
{
  "pdf_field_name": {
    "is_excel_col": false,
    "is_excel_image": true,
    "val": "B"
  }
}
```

#### 2. 支持的Excel图片格式

**WPS DISPIMG格式：**
```
单元格内容：=_xlfn.DISPIMG("ID_12345",1)
程序会自动提取ID为"ID_12345"的图片
```

**Office浮动图片：**
```
单元格位置：B2
程序会自动检测该位置的浮动图片并提取
```

#### 3. 图片提取逻辑

**精确位置匹配：**
- 程序只提取与指定单元格位置完全匹配的图片
- 对于浮动图片，使用0偏差的完全精确匹配
- 不会提取Excel中所有图片，只提取目标位置的图片

**提取优先级：**
1. 首先检查单元格是否包含DISPIMG格式
2. 如果没有DISPIMG，检查该位置是否有浮动图片
3. 如果都没有找到，不会使用备用方案提取其他图片

### 二开说明

> 企业用户，可开启环境检测和修改内部网络地址，目前使用判断环境来控制非企业内部无法使用的处理，也可自行修改逻辑。


## 使用说明

### 程序运行

> 运行文件：ExcelToPDFTemplate.exe

## 安装说明

### 1. 进入项目目录
```sh
cd ExcelToPDFTemplate
```

### 3. 安装相关依赖
```sh
pip install -r requirements.txt
```

### 4. 运行文件
```sh
python -u main.py
```

### 5. 打包exe
```sh
python setup.py build

# 打包msi
# python setup.py bdist_msi
```
