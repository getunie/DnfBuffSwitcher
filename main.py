import sys
import os
import random
import json
import threading
import shutil
import zipfile
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QTableWidget, QTableWidgetItem, QPushButton, QLineEdit, 
                             QComboBox, QSpinBox, QFileDialog, QMessageBox, QLabel,
                             QCheckBox, QTextEdit, QGroupBox, QHeaderView, 
                             QSystemTrayIcon, QMenu, QAction, QDialog, 
                             QDialogButtonBox, QStatusBar)
from PyQt5.QtGui import QIcon
from openpyxl import load_workbook

# 路径处理：config.json放在exe所在目录
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, 'config.json')
DEFAULT_EXCEL_PATH = os.path.join(APP_DIR, 'BUFF动画职业名对照表.xlsx')


class PresetNameDialog(QDialog):
    """自定义预设命名对话框，替代QInputDialog避免打包后阻塞"""
    def __init__(self, parent=None, title='保存预设', label='输入预设名称:', default=''):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(label))
        self.input = QLineEdit(default)
        layout.addWidget(self.input)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
    
    def get_name(self):
        return self.input.text().strip()


class GlobalFileManager:
    """全局文件管理器，解决多职业同关键词冲突"""
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self.keyword_to_vocations = {}
        self.vocation_configs = {}
        self._op_lock = threading.Lock()
    
    def register(self, vocation, config):
        with self._op_lock:
            self.vocation_configs[vocation] = config
            keyword = config.get('keyword', '')
            if keyword not in self.keyword_to_vocations:
                self.keyword_to_vocations[keyword] = []
            if vocation not in self.keyword_to_vocations[keyword]:
                self.keyword_to_vocations[keyword].append(vocation)
    
    def unregister(self, vocation):
        with self._op_lock:
            if vocation in self.vocation_configs:
                config = self.vocation_configs[vocation]
                keyword = config.get('keyword', '')
                if keyword in self.keyword_to_vocations:
                    if vocation in self.keyword_to_vocations[keyword]:
                        self.keyword_to_vocations[keyword].remove(vocation)
                del self.vocation_configs[vocation]
    
    def is_keyword_shared(self, keyword):
        with self._op_lock:
            return len(self.keyword_to_vocations.get(keyword, [])) > 1
    
    def get_occupied_targets(self, keyword, exclude_vocation=''):
        with self._op_lock:
            occupied = set()
            for voc, cfg in self.vocation_configs.items():
                if voc != exclude_vocation and cfg.get('keyword', '') == keyword:
                    t = cfg.get('target_filename', '')
                    if t:
                        occupied.add(t)
            return occupied


class BuffSwitcher(QtCore.QObject):
    def __init__(self, config, file_manager):
        super().__init__()
        self.config = config
        self.file_manager = file_manager
        self.running = False
        self.timer = None
        
    def switch(self):
        try:
            filepath = self.config.get('filepath', '')
            target_filename = self.config.get('target_filename', '')
            keyword = self.config.get('keyword', '')
            vocation = self.config.get('vocation', '')
            
            if not os.path.isdir(filepath) or not target_filename or not keyword:
                return False
            
            target_path = os.path.join(filepath, target_filename)
            backup_files = []
            
            # 如果目标文件已存在，先重命名为备用（解除占用）
            if os.path.exists(target_path):
                new_name = self.generate_unique_name(filepath, keyword) + '.bk2'
                new_path = os.path.join(filepath, new_name)
                os.rename(target_path, new_path)
                backup_files.append(new_path)
            
            # 收集所有含关键词的.bk2文件
            occupied = self.file_manager.get_occupied_targets(keyword, vocation)
            for f in os.listdir(filepath):
                if (keyword.lower() in f.lower() 
                    and f.lower().endswith('.bk2') 
                    and f != target_filename
                    and f not in occupied):
                    backup_files.append(os.path.join(filepath, f))
            
            if backup_files:
                random_file = random.choice(backup_files)
                os.rename(random_file, target_path)
                return True
            else:
                print(f"[{vocation}] 警告: 没有找到可用文件")
                return False
        except Exception as e:
            print(f"[{vocation}] 切换失败: {e}")
            return False
    
    def generate_unique_name(self, filepath, keyword):
        existing = set(os.listdir(filepath))
        for _ in range(100):
            name = f"{keyword}{random.randint(1000, 9999)}"
            if name not in existing:
                return name
        return f"{keyword}{random.randint(10000, 99999)}"
    
    def start(self):
        if self.running:
            return
        self.running = True
        self.switch()
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.switch)
        self.timer.start(self.config.get('interval', 30) * 1000)
    
    def stop(self):
        self.running = False
        if self.timer:
            self.timer.stop()
            self.timer = None


class VocationMapping:
    def __init__(self):
        self.mappings = {}
    
    def load_from_excel(self, filepath):
        try:
            wb = load_workbook(filepath, data_only=True)
            ws = wb.active
            
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            if not header_row:
                print("导入Excel失败: 表头为空")
                return False
            
            vocation_col_idx = -1
            filename_col_idx = -1
            for idx, cell in enumerate(header_row):
                if cell and '职业' in str(cell):
                    vocation_col_idx = idx
                elif cell and '文件' in str(cell):
                    filename_col_idx = idx
            
            if vocation_col_idx == -1 or filename_col_idx == -1:
                print(f"导入Excel失败: 未找到职业名列或文件名列，表头: {header_row}")
                return False
            
            self.mappings = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row:
                    if len(row) > max(vocation_col_idx, filename_col_idx):
                        vocation_cell = row[vocation_col_idx]
                        filename_cell = row[filename_col_idx]
                        
                        if vocation_cell and filename_cell:
                            vocation = str(vocation_cell).strip()
                            target_filename = str(filename_cell).strip()
                            
                            if vocation and target_filename:
                                if '【' in vocation or '】' in vocation:
                                    continue
                                self.mappings[vocation] = target_filename
            
            return True
        except Exception as e:
            print(f"导入Excel失败: {e}")
            return False
    
    def get_vocations(self):
        return list(self.mappings.keys())
    
    def get_target_filename(self, vocation):
        filename = self.mappings.get(vocation, '')
        if filename and not filename.lower().endswith('.bk2'):
            return filename + '.bk2'
        return filename


class PresetManager:
    def __init__(self):
        self.presets = {}
        self.last_preset = ''
        self.last_error = ''
        self.global_settings = {}
        self.load_presets()
    
    def load_presets(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.presets = data.get('presets', {})
                    self.last_preset = data.get('last_preset', '')
                    self.global_settings = data.get('global_settings', {})
        except Exception as e:
            self.last_error = str(e)
            self.presets = {}
            self.global_settings = {}
    
    def save_presets(self):
        try:
            # 先序列化到内存，确保数据可序列化
            data = {
                'presets': self.presets,
                'last_preset': self.last_preset,
                'global_settings': self.global_settings
            }
            json_str = json.dumps(data, ensure_ascii=False, indent=2)
            
            # 写入临时文件再替换，避免写入失败导致配置损坏
            tmp_file = CONFIG_FILE + '.tmp'
            with open(tmp_file, 'w', encoding='utf-8') as f:
                f.write(json_str)
            
            # 替换原文件
            if os.path.exists(CONFIG_FILE):
                os.replace(tmp_file, CONFIG_FILE)
            else:
                os.rename(tmp_file, CONFIG_FILE)
            
            self.last_error = ''
            return True
        except Exception as e:
            self.last_error = str(e)
            return False
    
    def set_last_preset(self, name):
        self.last_preset = name
        self.save_presets()
    
    def save_global_settings(self, settings):
        """单独保存全局设置，不依赖预设"""
        self.global_settings = settings
        self.save_presets()
    
    def get_global_settings(self):
        return self.global_settings
    
    def add_preset(self, name, config):
        self.presets[name] = config
        return self.save_presets()
    
    def remove_preset(self, name):
        if name in self.presets:
            del self.presets[name]
            return self.save_presets()
        return False
    
    def get_preset(self, name):
        return self.presets.get(name, None)
    
    def get_all_presets(self):
        return list(self.presets.keys())


class MainWindow(QMainWindow):
    def __init__(self, start_minimized=False):
        super().__init__()
        self.setWindowTitle("DNF Buff动画随机切换工具")
        self.setGeometry(100, 100, 900, 650)
        
        self.vocation_mapping = VocationMapping()
        self.preset_manager = PresetManager()
        self.switchers = {}
        self.global_filepath = ''
        self.file_manager = GlobalFileManager()
        self._closing = False
        self._loading = False  # 加载预设时禁止触发保存
        
        self.init_ui()
        self.init_tray()
        
        # 自动加载默认对照表
        if os.path.exists(DEFAULT_EXCEL_PATH):
            self.vocation_mapping.load_from_excel(DEFAULT_EXCEL_PATH)
        
        # 先恢复全局设置（不依赖预设）
        self._loading = True
        gs = self.preset_manager.get_global_settings()
        self.auto_start_check.setChecked(gs.get('auto_start', False))
        self.boot_start_check.setChecked(gs.get('boot_start', False))
        self.hide_after_start_check.setChecked(gs.get('hide_after_start', False))
        self.close_to_tray_check.setChecked(gs.get('close_to_tray', True))
        if gs.get('filepath', ''):
            self.global_filepath = gs.get('filepath', '')
            self.path_edit.setText(self.global_filepath)
        self.default_interval_spin.setValue(gs.get('default_interval', 30))
        self._loading = False
        
        # 启动时自动加载上次使用的预设（静默）
        if self.preset_manager.last_preset and self.preset_manager.last_preset in self.preset_manager.get_all_presets():
            self.load_preset_by_name(self.preset_manager.last_preset, silent=True)
        
        self.verify_boot_start()
        
        # 仅开机自启时最小化到托盘
        if start_minimized:
            self.hide()
    
    def init_tray(self):
        # 用系统标准图标确保托盘可见
        style = self.style()
        icon = style.standardIcon(QtWidgets.QStyle.SP_MediaPlay)
        
        self.tray_icon = QSystemTrayIcon(icon, self)
        self.tray_icon.setToolTip("DNF Buff切换工具")
        
        tray_menu = QMenu()
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show_from_tray)
        tray_menu.addAction(show_action)
        
        tray_menu.addSeparator()
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()
    
    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_from_tray()
    
    def show_from_tray(self):
        self.show()
        self.activateWindow()
        self.raise_()
    
    def quit_app(self):
        self._closing = True
        self.stop_all()
        self.tray_icon.hide()
        QApplication.quit()
    
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 全局设置
        path_group = QGroupBox('全局设置')
        path_layout = QHBoxLayout(path_group)
        
        path_layout.addWidget(QLabel('Buff动画路径:'))
        self.path_edit = QLineEdit()
        path_layout.addWidget(self.path_edit)
        self.path_btn = QPushButton('浏览')
        self.path_btn.clicked.connect(self.browse_global_path)
        path_layout.addWidget(self.path_btn)
        
        path_layout.addWidget(QLabel('默认间隔:'))
        self.default_interval_spin = QSpinBox()
        self.default_interval_spin.setRange(2, 300)
        self.default_interval_spin.setValue(30)
        path_layout.addWidget(self.default_interval_spin)
        
        self.auto_start_check = QCheckBox('自动全部启动')
        self.auto_start_check.stateChanged.connect(self.on_global_setting_changed)
        path_layout.addWidget(self.auto_start_check)
        
        self.boot_start_check = QCheckBox('开机自启动')
        self.boot_start_check.stateChanged.connect(self.on_boot_start_changed)
        path_layout.addWidget(self.boot_start_check)
        
        self.hide_after_start_check = QCheckBox('启动后后台运行')
        self.hide_after_start_check.stateChanged.connect(self.on_global_setting_changed)
        path_layout.addWidget(self.hide_after_start_check)
        
        self.close_to_tray_check = QCheckBox('关闭时后台运行')
        self.close_to_tray_check.setChecked(True)
        self.close_to_tray_check.stateChanged.connect(self.on_global_setting_changed)
        path_layout.addWidget(self.close_to_tray_check)
        
        layout.addWidget(path_group)
        
        # 职业配置表
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['职业', '关键词', '间隔(秒)', '状态'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.table)
        
        # 操作按钮
        btn_layout = QHBoxLayout()
        
        self.add_btn = QPushButton('添加')
        self.add_btn.clicked.connect(self.add_row)
        btn_layout.addWidget(self.add_btn)
        
        self.delete_btn = QPushButton('删除')
        self.delete_btn.clicked.connect(self.delete_row)
        btn_layout.addWidget(self.delete_btn)
        
        self.import_btn = QPushButton('导入对应表')
        self.import_btn.clicked.connect(lambda: self.import_excel())
        btn_layout.addWidget(self.import_btn)
        
        self.start_all_btn = QPushButton('全部启动')
        self.start_all_btn.clicked.connect(self.start_all)
        btn_layout.addWidget(self.start_all_btn)
        
        self.stop_all_btn = QPushButton('全部暂停')
        self.stop_all_btn.clicked.connect(self.stop_all)
        btn_layout.addWidget(self.stop_all_btn)
        
        self.restore_btn = QPushButton('还原默认')
        self.restore_btn.clicked.connect(self.restore_default)
        btn_layout.addWidget(self.restore_btn)
        
        layout.addLayout(btn_layout)
        
        # 预设管理
        preset_layout = QHBoxLayout()
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(self.preset_manager.get_all_presets())
        if self.preset_manager.last_preset:
            self.preset_combo.setCurrentText(self.preset_manager.last_preset)
        preset_layout.addWidget(QLabel('预设:'))
        preset_layout.addWidget(self.preset_combo)
        
        self.save_preset_btn = QPushButton('保存预设')
        self.save_preset_btn.clicked.connect(self.save_preset)
        preset_layout.addWidget(self.save_preset_btn)
        
        self.load_preset_btn = QPushButton('加载预设')
        self.load_preset_btn.clicked.connect(self.load_preset)
        preset_layout.addWidget(self.load_preset_btn)
        
        self.del_preset_btn = QPushButton('删除预设')
        self.del_preset_btn.clicked.connect(self.delete_preset)
        preset_layout.addWidget(self.del_preset_btn)
        
        layout.addLayout(preset_layout)
        
        # 使用说明
        help_group = QGroupBox('使用说明')
        help_layout = QVBoxLayout(help_group)
        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        self.help_text.setMaximumHeight(150)
        self.help_text.setHtml("""
        <strong>Q:如何使用？</strong><br>
        A:导入一起下载的对应表，点击添加选取你想随机的职业，并将你想为这个职业随机的动画文件全部重命名为含有同个关键词的命名，如将为剑魂的全部批量重命名为：RenameSword，那么设定关键词为：RenameSword就会自动按设定间隔与方式切换剑魂的buff动画<br><br>
        <strong>Q:可以多个职业的BUFF动画都切换吗？</strong><br>
        A:可以，上方添加另一个职业的即可。<br><br>
        <strong>Q:可以为不同职业设定不同动画吗？</strong><br>
        A:可以，将A职业的关键词为关键词A，B职业的关键词为关键词B，两个关键词区分即可，但是关键词不能包含另一个，如剑魂为RenameSword，剑帝为RenameSwordF就会混到一起，想区分开来可以改剑帝的为RenameFSword<br><br>
        <strong>Q:关键词可以相同吗？</strong><br>
        A:可以，支持多职业共享同一关键词，自动处理文件冲突。
        """)
        help_layout.addWidget(self.help_text)
        layout.addWidget(help_group)
        
        # 状态栏（替代弹窗提示）
        self.setStatusBar(QStatusBar())
    
    def show_status(self, msg):
        """状态栏显示消息，避免弹窗阻塞"""
        self.statusBar().showMessage(msg, 3000)
    
    def on_global_setting_changed(self):
        """全局设置变化时自动保存"""
        if self._loading:
            return
        settings = {
            'auto_start': self.auto_start_check.isChecked(),
            'boot_start': self.boot_start_check.isChecked(),
            'hide_after_start': self.hide_after_start_check.isChecked(),
            'close_to_tray': self.close_to_tray_check.isChecked(),
            'filepath': self.global_filepath,
            'default_interval': self.default_interval_spin.value()
        }
        self.preset_manager.save_global_settings(settings)
    
    def on_boot_start_changed(self, state):
        """开机自启动变化时，先设置注册表再保存全局设置"""
        if self._loading:
            return
        self.toggle_boot_start(state)
        self.on_global_setting_changed()
    
    def browse_global_path(self):
        path = QFileDialog.getExistingDirectory(self, '选择Buff动画文件夹')
        if path:
            self.path_edit.setText(path)
            self.global_filepath = path
            self.on_global_setting_changed()
            self.show_status(f'路径已设置: {path}')
    
    def import_excel(self, filepath=None):
        if filepath is None:
            filepath, _ = QFileDialog.getOpenFileName(self, '导入Excel文件', '', 'Excel文件 (*.xlsx)')
        if filepath:
            if self.vocation_mapping.load_from_excel(filepath):
                self.update_vocation_combos()
                self.show_status(f'导入成功! 共{len(self.vocation_mapping.get_vocations())}个职业')
    
    def update_vocation_combos(self):
        vocations = self.vocation_mapping.get_vocations()
        for row in range(self.table.rowCount()):
            combo = self.table.cellWidget(row, 0)
            if combo:
                current_text = combo.currentText()
                combo.clear()
                combo.addItems(vocations)
                if current_text in vocations:
                    combo.setCurrentText(current_text)
    
    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        vocation_combo = QComboBox()
        vocation_combo.addItems(self.vocation_mapping.get_vocations())
        self.table.setCellWidget(row, 0, vocation_combo)
        
        keyword_edit = QLineEdit()
        self.table.setCellWidget(row, 1, keyword_edit)
        
        interval_spin = QSpinBox()
        interval_spin.setRange(2, 300)
        interval_spin.setValue(self.default_interval_spin.value())
        self.table.setCellWidget(row, 2, interval_spin)
        
        start_btn = QPushButton('启动')
        start_btn.clicked.connect(lambda checked, r=row: self.toggle_switcher(r))
        self.table.setCellWidget(row, 3, start_btn)
    
    def delete_row(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.stop_switcher(current_row)
            self.table.removeRow(current_row)
    
    def get_row_config(self, row):
        """获取行配置，只返回可序列化的基本类型"""
        vocation_combo = self.table.cellWidget(row, 0)
        keyword_edit = self.table.cellWidget(row, 1)
        interval_spin = self.table.cellWidget(row, 2)
        
        vocation = vocation_combo.currentText() if vocation_combo else ''
        keyword = keyword_edit.text() if keyword_edit else ''
        interval = interval_spin.value() if interval_spin else 30
        filepath = self.global_filepath
        target_filename = self.vocation_mapping.get_target_filename(vocation)
        
        return {
            'vocation': str(vocation),
            'keyword': str(keyword),
            'interval': int(interval),
            'filepath': str(filepath),
            'target_filename': str(target_filename)
        }
    
    def toggle_switcher(self, row):
        btn = self.table.cellWidget(row, 3)
        if not btn:
            return
        
        if btn.text() == '启动':
            config = self.get_row_config(row)
            if not config['vocation']:
                QMessageBox.warning(self, '警告', '请选择职业!')
                return
            if not config['keyword']:
                QMessageBox.warning(self, '警告', '请输入关键词!')
                return
            if not config['filepath']:
                QMessageBox.warning(self, '警告', '请设置Buff动画路径!')
                return
            if not config['target_filename']:
                QMessageBox.warning(self, '警告', '未找到职业对应的目标文件名!')
                return
            
            self.file_manager.register(config['vocation'], config)
            switcher = BuffSwitcher(config, self.file_manager)
            switcher.start()
            self.switchers[row] = switcher
            btn.setText('停止')
            btn.setStyleSheet('background-color: #4CAF50; color: white;')
            self.show_status(f'{config["vocation"]} 已启动')
            
        else:
            self.stop_switcher(row)
    
    def stop_switcher(self, row):
        if row in self.switchers:
            config = self.get_row_config(row)
            self.file_manager.unregister(config.get('vocation', ''))
            self.switchers[row].stop()
            del self.switchers[row]
        btn = self.table.cellWidget(row, 3)
        if btn:
            btn.setText('启动')
            btn.setStyleSheet('')
    
    def start_all(self):
        if not self.global_filepath:
            QMessageBox.warning(self, '警告', '请先设置Buff动画路径!')
            return
        
        started = 0
        for row in range(self.table.rowCount()):
            btn = self.table.cellWidget(row, 3)
            if btn and btn.text() == '启动':
                config = self.get_row_config(row)
                if config['vocation'] and config['keyword'] and config['target_filename']:
                    self.file_manager.register(config['vocation'], config)
                    switcher = BuffSwitcher(config, self.file_manager)
                    switcher.start()
                    self.switchers[row] = switcher
                    btn.setText('停止')
                    btn.setStyleSheet('background-color: #4CAF50; color: white;')
                    started += 1
        
        self.show_status(f'已启动 {started} 个职业')
    
    def stop_all(self):
        for row in list(self.switchers.keys()):
            if row in self.switchers:
                config = self.get_row_config(row)
                self.file_manager.unregister(config.get('vocation', ''))
                self.switchers[row].stop()
                del self.switchers[row]
                btn = self.table.cellWidget(row, 3)
                if btn:
                    btn.setText('启动')
                    btn.setStyleSheet('')
        self.show_status('已全部停止')
    
    def restore_default(self):
        """还原默认动画文件：自动识别压缩文件或文件夹恢复"""
        if not self.global_filepath:
            QMessageBox.warning(self, '警告', '请先设置Buff动画路径!')
            return
        
        # 使用Directory模式+非原生对话框，可同时选择文件或文件夹
        fd = QFileDialog(self, '选择备份(zip文件或文件夹)')
        fd.setFileMode(QFileDialog.Directory)
        fd.setOption(QFileDialog.DontUseNativeDialog, True)
        fd.setOption(QFileDialog.ShowDirsOnly, False)
        if not fd.exec_():
            return
        
        selected = fd.selectedFiles()
        if not selected:
            return
        
        src_path = selected[0]
        
        # 自动判断类型
        if os.path.isfile(src_path) and src_path.lower().endswith('.zip'):
            reply = QMessageBox.question(self, '确认还原',
                f'将从压缩文件还原动画到:\n{self.global_filepath}\n\n将覆盖同名文件，是否继续?')
            if reply != QMessageBox.Yes:
                return
            try:
                with zipfile.ZipFile(src_path, 'r') as zf:
                    zf.extractall(self.global_filepath)
                self.show_status('还原成功!')
                QMessageBox.information(self, '成功', '动画文件已还原!')
            except Exception as e:
                QMessageBox.warning(self, '错误', f'还原失败:\n{e}')
        elif os.path.isdir(src_path):
            reply = QMessageBox.question(self, '确认还原',
                f'将从文件夹还原动画到:\n{self.global_filepath}\n\n将覆盖同名文件，是否继续?')
            if reply != QMessageBox.Yes:
                return
            try:
                count = 0
                for item in os.listdir(src_path):
                    src = os.path.join(src_path, item)
                    dst = os.path.join(self.global_filepath, item)
                    if os.path.isfile(src):
                        shutil.copy2(src, dst)
                        count += 1
                    elif os.path.isdir(src):
                        if os.path.exists(dst):
                            shutil.rmtree(dst)
                        shutil.copytree(src, dst)
                        count += 1
                self.show_status(f'还原成功! 共{count}个文件/文件夹')
                QMessageBox.information(self, '成功', f'动画文件已还原! 共{count}个文件/文件夹')
            except Exception as e:
                QMessageBox.warning(self, '错误', f'还原失败:\n{e}')
        else:
            QMessageBox.warning(self, '警告', '请选择zip压缩文件或文件夹!')
    
    def save_preset(self):
        """保存预设 - 使用自定义对话框避免阻塞"""
        # 先弹窗获取名称
        dialog = PresetNameDialog(self, '保存预设', '输入预设名称:')
        if dialog.exec_() != QDialog.Accepted:
            return
        
        name = dialog.get_name()
        if not name:
            return
        
        # 收集配置（确保都是基本类型）
        configs = []
        for row in range(self.table.rowCount()):
            try:
                config = self.get_row_config(row)
                if config['vocation'] and config['keyword']:
                    configs.append(config)
            except Exception as e:
                print(f"读取第{row}行配置失败: {e}")
        
        preset_data = {
            'filepath': str(self.global_filepath),
            'default_interval': int(self.default_interval_spin.value()),
            'auto_start': bool(self.auto_start_check.isChecked()),
            'boot_start': bool(self.boot_start_check.isChecked()),
            'hide_after_start': bool(self.hide_after_start_check.isChecked()),
            'close_to_tray': bool(self.close_to_tray_check.isChecked()),
            'configs': configs
        }
        
        # 保存
        if self.preset_manager.add_preset(name, preset_data):
            self.preset_manager.set_last_preset(name)
            # 更新下拉框
            self.preset_combo.clear()
            self.preset_combo.addItems(self.preset_manager.get_all_presets())
            self.preset_combo.setCurrentText(name)
            self.show_status(f'预设 "{name}" 保存成功!')
        else:
            err = self.preset_manager.last_error or '未知错误'
            QMessageBox.warning(self, '错误', f'保存预设失败!\n错误: {err}\n配置路径: {CONFIG_FILE}')
    
    def load_preset(self):
        preset_name = self.preset_combo.currentText()
        if preset_name:
            self.load_preset_by_name(preset_name, silent=False)
    
    def load_preset_by_name(self, preset_name, silent=False):
        preset_data = self.preset_manager.get_preset(preset_name)
        if not preset_data:
            return
        
        self._loading = True  # 加载时禁止触发保存
        self.stop_all()
        
        # 清空表格
        while self.table.rowCount() > 0:
            self.table.removeRow(0)
        
        # 恢复全局设置
        self.global_filepath = preset_data.get('filepath', '')
        self.path_edit.setText(self.global_filepath)
        self.default_interval_spin.setValue(preset_data.get('default_interval', 30))
        self.auto_start_check.setChecked(preset_data.get('auto_start', False))
        self.boot_start_check.setChecked(preset_data.get('boot_start', False))
        self.hide_after_start_check.setChecked(preset_data.get('hide_after_start', False))
        self.close_to_tray_check.setChecked(preset_data.get('close_to_tray', True))
        self._loading = False  # 恢复后保存一次最新状态
        self.on_global_setting_changed()
        
        self.preset_manager.set_last_preset(preset_name)
        self.preset_combo.setCurrentText(preset_name)
        
        # 恢复职业配置
        configs = preset_data.get('configs', [])
        for config in configs:
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            vocation_combo = QComboBox()
            vocation_combo.addItems(self.vocation_mapping.get_vocations())
            vocation_combo.setCurrentText(config.get('vocation', ''))
            self.table.setCellWidget(row, 0, vocation_combo)
            
            keyword_edit = QLineEdit(config.get('keyword', ''))
            self.table.setCellWidget(row, 1, keyword_edit)
            
            interval_spin = QSpinBox()
            interval_spin.setRange(2, 300)
            interval_spin.setValue(config.get('interval', 30))
            self.table.setCellWidget(row, 2, interval_spin)
            
            start_btn = QPushButton('启动')
            start_btn.clicked.connect(lambda checked, r=row: self.toggle_switcher(r))
            self.table.setCellWidget(row, 3, start_btn)
        
        # 自动启动
        if self.auto_start_check.isChecked():
            self.start_all()
        
        if not silent:
            self.show_status(f'预设 "{preset_name}" 加载成功!')
    
    def delete_preset(self):
        preset_name = self.preset_combo.currentText()
        if preset_name:
            reply = QMessageBox.question(self, '确认', f'确定删除预设 "{preset_name}" 吗?')
            if reply == QMessageBox.Yes:
                self.preset_manager.remove_preset(preset_name)
                self.preset_combo.removeItem(self.preset_combo.currentIndex())
                if self.preset_manager.last_preset == preset_name:
                    self.preset_manager.set_last_preset('')
                self.show_status('预设已删除')
    
    def get_startup_shortcut_path(self):
        """获取Startup文件夹中快捷方式的路径"""
        import os
        startup_folder = os.path.join(os.environ['APPDATA'], 
                                      'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        return os.path.join(startup_folder, 'DNFBuffSwitcher.lnk')
    
    def create_startup_shortcut(self, exe_path, args=''):
        """在Startup文件夹创建快捷方式"""
        import os
        import pythoncom
        from win32com.client import Dispatch
        
        shortcut_path = self.get_startup_shortcut_path()
        working_dir = os.path.dirname(exe_path)
        
        pythoncom.CoInitialize()
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = exe_path
        if args:
            shortcut.Arguments = args
        shortcut.WorkingDirectory = working_dir
        shortcut.Save()
        pythoncom.CoUninitialize()
        return True
    
    def delete_startup_shortcut(self):
        """删除Startup文件夹中的快捷方式"""
        import os
        shortcut_path = self.get_startup_shortcut_path()
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            return True
        return False
    
    def has_startup_shortcut(self):
        """检查是否存在开机自启快捷方式"""
        import os
        return os.path.exists(self.get_startup_shortcut_path())
    
    def verify_boot_start(self):
        """启动时自动校验并修正开机自启路径"""
        import os
        import subprocess
        import winreg
        
        if not self.has_startup_shortcut():
            return
        
        shortcut_path = self.get_startup_shortcut_path()
        
        exe_path = sys.executable if hasattr(sys, 'frozen') else os.path.abspath(__file__)
        exe_path = os.path.abspath(exe_path)
        
        try:
            import pythoncom
            from win32com.client import Dispatch
            
            pythoncom.CoInitialize()
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            existing_target = shortcut.TargetPath
            pythoncom.CoUninitialize()
            
            if existing_target != exe_path:
                self.delete_startup_shortcut()
                self.create_startup_shortcut(exe_path, '--minimized')
                print(f"开机自启路径已更新: {exe_path}")
            else:
                print("开机自启路径正确")
        except Exception as e:
            print(f"校验开机自启快捷方式失败: {e}")
    
    def toggle_boot_start(self, state):
        """用Startup文件夹快捷方式实现开机自启（最稳定）"""
        exe_path = sys.executable if hasattr(sys, 'frozen') else os.path.abspath(__file__)
        exe_path = os.path.abspath(exe_path)
        
        try:
            if state == 2:  # 勾选
                if self.create_startup_shortcut(exe_path, '--minimized'):
                    self.show_status('已设置开机自启动(启动文件夹)')
                else:
                    import winreg
                    cmd_with_args = f'"{exe_path}" --minimized'
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                         r'Software\Microsoft\Windows\CurrentVersion\Run',
                                         0, winreg.KEY_SET_VALUE)
                    winreg.SetValueEx(key, 'DNFBuffSwitcher', 0, winreg.REG_SZ, cmd_with_args)
                    winreg.CloseKey(key)
                    self.show_status('已设置开机自启动(注册表)')
            else:  # 取消
                self.delete_startup_shortcut()
                try:
                    import winreg
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                         r'Software\Microsoft\Windows\CurrentVersion\Run',
                                         0, winreg.KEY_SET_VALUE)
                    winreg.DeleteValue(key, 'DNFBuffSwitcher')
                    winreg.CloseKey(key)
                except FileNotFoundError:
                    pass
                self.show_status('已取消开机自启动')
        except Exception as e:
            print(f"设置开机自启动失败: {e}")
    
    def closeEvent(self, event):
        if self._closing:
            # 真正退出
            event.accept()
            return
        
        if self.close_to_tray_check.isChecked():
            # 关闭时后台运行：隐藏到托盘
            self.hide()
            event.ignore()
        else:
            # 直接关闭
            self.stop_all()
            self.tray_icon.hide()
            event.accept()


if __name__ == '__main__':
    import socket
    
    # 单实例检测
    def is_single_instance():
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.bind(('127.0.0.1', 19876))
            return sock
        except OSError:
            return None
    
    sock = is_single_instance()
    if sock is None:
        QMessageBox.warning(None, '提示', '程序已在运行中！')
        sys.exit(0)
    
    start_minimized = '--minimized' in sys.argv
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    window = MainWindow(start_minimized=start_minimized)
    
    if not start_minimized:
        window.show()
    
    exit_code = app.exec_()
    
    # 关闭时清理子进程
    import os
    import signal
    try:
        parent_pid = os.getpid()
        result = os.popen(f'wmic process where "ParentProcessId={parent_pid}" get ProcessId').read()
        for line in result.strip().split('\n')[1:]:
            line = line.strip()
            if line.isdigit():
                try:
                    os.kill(int(line), signal.SIGTERM)
                except:
                    pass
    except:
        pass
    
    sock.close()
    sys.exit(exit_code)
