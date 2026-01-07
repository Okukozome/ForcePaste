# ForcePaste (强制粘贴工具)

**ForcePaste** 是一款轻量级的模拟键盘输入工具。通过模拟真实按键来输入剪贴板内容，用于解决系统剪贴板被禁用或受限的场景（如远程桌面、VNC、受限网页表单等）。

## ✨ 功能特点

* **模拟输入**：通过逐字敲击键盘的方式“粘贴”文本，绕过任何粘贴限制。
* **防楼梯模式**：针对 Python 代码缩进进行了优化，防止在编辑器中出现排版错乱。
* **可配置延迟**：支持设置启动延迟、字符间隔及随机抖动，模拟真实的人工输入频率。
* **多语言界面**：原生支持中文与英文界面切换。
* **全局快捷键**：支持通过快捷键触发（默认：`Ctrl+Shift+Y`），无需频繁切换窗口。


## 🛠 安装与使用

### 环境要求

* **操作系统**：Windows (需以管理员身份运行)
* **Python 版本**：Python 3.x

### 源码运行步骤

1. **克隆仓库**
```terminal
git clone https://github.com/Okukozome/forcepaste.git
cd forcepaste
```

2. **安装依赖**
```terminal
pip install keyboard
```

3. **启动程序**
```terminal
# 必须以管理员权限运行，否则无法拦截/模拟底层键盘钩子
python main.py
```


## ⚙️ 设置说明

在应用中点击 **Settings (设置)** 按钮可以调整以下参数：

### 1. 延迟设置 (Delays)

* **启动延迟**：按下快捷键后，程序等待多久才开始输入（留出时间切换窗口）。
* **输入速度**：每个字符之间的间隔时间（毫秒）。

### 2. 控制设置 (Controls)

* **快捷键**：自定义触发模拟粘贴的全局组合键。

### 3. 功能开关 (Features)

* **Tab 转 4 空格**：自动转换制表符，确保代码在不同编辑器中格式一致。
* **使用 Shift+Enter 换行**：适配某些特定的即时通讯或终端环境。
* **窗口置顶**：让 ForcePaste 小工具始终浮动在最上层。


---

# ForcePaste

**ForcePaste** is a lightweight tool designed to simulate keyboard input for pasting text. It is particularly useful in environments where the system clipboard is disabled or restricted (e.g., remote desktops, VNC, or restricted web forms).

## ✨ Features

* **Simulated Typing**: Pastes text by typing it out character by character, bypassing standard paste restrictions.
* **Anti-Staircase Mode**: Specifically optimized for Python code indentation to prevent formatting issues in editors.
* **Configurable Delays**: Adjust start delay, character interval, and random jitter to mimic human behavior.
* **Multilingual UI**: Supports both English and Chinese interfaces.
* **Hotkey Support**: Trigger pasting via a global hotkey (Default: `Ctrl+Shift+Y`).


## 🛠 Installation & Usage

### Requirements

* **OS**: Windows (Run as Administrator)
* **Python**: 3.x

### Run from Source

1. **Clone the repository**
```terminal
git clone https://github.com/Okukozome/forcepaste.git
cd forcepaste
```

2. **Install dependencies**
```terminal
pip install keyboard
```

3. **Run the application**
```terminal
# Must be run as Administrator for low-level keyboard hook access
python main.py
```


## ⚙️ Configuration

Click the **Settings** button in the app to configure:

* **Delays**: Adjust **Start Delay** (wait time before typing) and **Typing Speed** (interval between characters).
* **Controls**: Customize the global **Hotkey**.
* **Features**:
* **Convert Tab to 4 Spaces**: Ideal for consistent code formatting.
* **Use Shift+Enter for Newline**: For applications requiring specific newline triggers.
* **Always on Top**: Keep the tool window visible above all other windows.

---

## 📝 License

This project is licensed under the MIT License.