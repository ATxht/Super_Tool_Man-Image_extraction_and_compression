# 超级工具人——图片提取与压缩

![项目状态](https://img.shields.io/badge/状态-稳定运行-brightgreen)
![Python版本](https://img.shields.io/badge/Python-3.12-blue)
![开源协议](https://img.shields.io/badge/License-MIT-green)

## 项目简介
“超级工具人——图片提取与压缩” 是一款基于 Python 和 PyQt5 开发的图片处理工具。它为用户提供提取照片、重命名照片和压缩照片三大主要功能。用户只需通过直观的图形界面进行简单操作，就能轻松提高图片处理的效率。

## 核心功能
1. **照片提取**  
    - 能从指定的照片文件夹中，依据 Excel 文件里的身份证号，精准提取对应的照片到目标文件夹。
    - 支持选择 Excel 文件的单个工作表或全量工作表。

2. **照片重命名**  
    - 可以将源文件夹中的图片文件，根据文件名中的身份证号进行重命名，并复制到目标文件夹。

3. **照片压缩**  
    - 通过质量调整和尺寸缩放，把指定文件夹中的 JPG 图片压缩到指定的目标大小。

## 环境要求
- **操作系统**：支持 Windows、Mac OS、Linux 等主流操作系统。  
- **Python版本**：Python 3.12。  
- **依赖库**：运行前需安装 `PyQt5`、`openpyxl`、`pandas` 和 `pillow`。

## 快速开始

### 步骤1：获取代码
```bash
git clone https://github.com/ATxht/Super_Tool_Man-Image_extraction_and_compression.git
cd Super_Tool_Man-Image_extraction_and_compression
```

### 步骤2：安装依赖
```bash
pip install PyQt5 openpyxl pandas pillow
```

### 步骤3：运行程序
```bash
python main.py
```

## 使用指南

### 照片提取步骤
```plaintext
1. 点击 “选择 Excel 文件” 按钮，挑选包含身份证号的 Excel 文件。
2. 选择要读取的工作表（可选 “全选”）。
3. 点击 “选择源文件夹” 按钮，选定包含照片的文件夹。
4. 点击 “选择目标文件夹” 按钮，确定提取照片的目标文件夹。
5. 点击 “开始提取照片” 按钮，开始提取操作。
```

### 照片重命名步骤
```plaintext
1. 点击 “选择源文件夹” 按钮，选择包含待重命名照片的文件夹。
2. 点击 “选择目标文件夹” 按钮，确定重命名后照片的目标文件夹。
3. 点击 “开始重命名照片” 按钮，进行重命名操作。
```

### 照片压缩步骤
```plaintext
1. 点击 “选择文件夹” 按钮，选择包含待压缩照片的文件夹。
2. 点击 “选择目标文件夹” 按钮，选择压缩后照片的目标文件夹。
3. 在 “目标大小（KB）” 输入框中输入目标大小。
4. 点击 “开始压缩照片” 按钮，启动压缩操作。
```

## 代码结构说明  
```python
# 主程序文件
main.py: 负责图形界面的创建和图片处理功能的实现。
```

## 注意事项  
1. **照片提取规则**  
    - 提取照片功能要求 Excel 文件中包含 “身份证号” 列。

2. **照片重命名规则**  
    - 重命名照片功能要求文件名中包含 18 位身份证号码。

3. **照片压缩限制**  
    - 压缩照片功能仅支持 JPG 格式的图片。

## 贡献与反馈  
- **提交Issue**：欢迎在 [issues](https://github.com/ATxht/Super_Tool_Man-Image_extraction_and_compression/issues) 反馈问题或提出改进建议。  
- **代码贡献**：提交 Pull Request 前请先创建 Issue 说明需求，确保代码风格一致。

## 开源协议  
本项目采用 **MIT License**，允许自由修改、分发和商业使用，但需保留原作者声明和协议文件。

## 联系方式  
- GitHub：[ATxht](https://github.com/ATxht)

如果这个图片处理工具对你有帮助，欢迎点亮 ⭐ 支持！ 
