# Reference-Formatter-Tool
一个用于格式化并导出 Word 专用学术文献列表的 Python/PyQt6 工具


# 📚 文献引用格式化工具 (Word专用)

<div align="center">

![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-GUI%20Framework-green)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-lightgrey)
![License](https://img.shields.io/badge/License-MIT-yellow)

[![GitHub](https://img.shields.io/badge/GitHub-rhj--flash-181717?style=for-the-badge&logo=github)](https://github.com/rhj-flash)
[![邮箱](https://img.shields.io/badge/邮箱-rhjflash@gmail.com-red?style=for-the-badge&logo=gmail)](mailto:rhjflash@gmail.com)
[![访问量](https://komarev.com/ghpvc/?username=rhj-flash&color=blue&style=for-the-badge)](https://github.com/rhj-flash)
[![GitHub Stars](https://img.shields.io/github/stars/rhj-flash?style=for-the-badge&logo=github&color=yellow)](https://github.com/rhj-flash)
[![GitHub Followers](https://img.shields.io/github/followers/rhj-flash?style=for-the-badge&logo=github&color=green)](https://github.com/rhj-flash)

一个基于 PyQt6 开发的智能文献格式化桌面应用程序，专门为学术写作优化，提供完美的 Word 文档兼容性。

**智能分割 · 统一格式 · Word专享 · 中英混排**

![软件界面截图](https://github.com/rhj-flash/Reference-Formatter-Tool/blob/main/example/1.png)

</div>

## ✨ 核心特性

### 🎯 智能化处理
- **智能文献分割** - 自动识别和分割多个文献条目，无需手动分隔
- **格式统一清洗** - 自动剥离旧编号，应用统一的新编号格式
- **字符规范化** - 全角/半角字符自动转换，标点符号智能处理

### 💻 Word深度集成
- **可编辑列表** - 在Word中保持编号列表的完全可编辑性
- **一键粘贴** - 标准Ctrl+V粘贴即可保持完整格式

### 🎨 现代用户体验
- **玻璃磨砂界面** - 采用现代化的Glassmorphism设计风格
- **三步操作流程** - 清晰的步骤引导，降低使用门槛
- **实时可视化预览** - 彩色区块显示分割结果，直观易懂
- **智能状态反馈** - 实时操作状态提示和错误处理

## 🚀 快速开始

### 系统要求
- **操作系统**: Windows 10/11 (推荐) 或支持 PyQt6 的其他系统
- **Python版本**: 3.7 或更高版本
- **依赖软件**: Microsoft Word (用于最终粘贴使用)

### 安装步骤

1. **克隆或下载项目**
```bash
git clone https://github.com/your-username/reference-formatter-tool.git
cd reference-formatter-tool
