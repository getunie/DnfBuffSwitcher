# DNF Buff动画随机切换工具

一个为DNF设计的Buff动画随机切换工具，通过重命名备用动画文件实现自动切换。

## 功能特点

- **自动重命名**：可设定时间间隔（最低2秒）自动切换Buff动画
- **多职业支持**：支持同时为多个职业配置不同的动画切换
- **配置持久化**：保存预设配置，启动时自动加载上一次使用的预设
- **后台运行**：支持最小化到系统托盘，开机自启动并后台运行
- **一键还原**：支持选择压缩文件或文件夹快速还原默认动画

## 技术栈

- Python 3.x
- PyQt5
- openpyxl
- PyInstaller

## 使用方法

**Q:如何使用？**
**A:**导入一起下载的对应表，点击添加选取你想随机的职业，并将你想为这个职业随机的动画文件全部重命名为含有同个关键词的命名，如将为剑魂的全部批量重命名为：RenameSword，那么设定关键词为：RenameSword就会自动按设定间隔与方式切换剑魂的buff动画

**Q:可以多个职业的BUFF动画都切换吗？**
**A:**可以，上方添加另一个职业的即可。

**Q:可以为不同职业设定不同动画吗？**
**A:**可以，将A职业的关键词为关键词A,B职业的关键词为关键词B,两个关键词区分即可，但是关键词不能包含另一个，如剑魂为RenameSword，剑帝为RenameSwordF就会混到一起，想区分开来可以改剑帝的为RenameFSword

**Q:关键词可以相同吗？**
**A:**可以，但是多职业相同关键词可能会导致动画全部被占用无法切换

## 系统要求

- Windows 10/11
- 管理员权限（用于文件操作和开机自启）

## 安装与运行

### 源代码运行

```bash
pip install -r requirements.txt
python main.py
```

### 打包成可执行文件

```bash
pyinstaller --onefile --windowed --uac-admin --name DNFBuffSwitcher main.py
```

## 项目结构

```
DnfBuffAnime/
├── main.py              # 主程序
├── requirements.txt     # 依赖列表
├── sample_data.py       # 示例数据生成脚本
├── BUFF动画职业名对照表.xlsx  # 职业与动画文件名对照表
├── LICENSE              # MIT开源协议
└── README.md            # 项目说明
```

## 许可证

MIT License