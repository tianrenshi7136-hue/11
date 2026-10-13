import json
import os
from pathlib import Path

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPainter, QPen
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (QApplication, QComboBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QTabWidget, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget)

from .core import ProjectStore, export_project, import_project, new_project, validate
from .templates import inspect_excel, readiness
from .sheetpile import SECTIONS, draft as sheetpile_draft, generate as generate_sheetpile, validate_request
from .sections import properties, YIELD_MPA, E_KN_PER_MM2, TABLE_VERSION

ROOT = Path(__file__).resolve().parent.parent
FIELDS = ["member_id", "kind", "case", "x_m", "y_m", "level_m", "effective_length_m", "n_kn", "m_knm", "v_kn", "value_basis", "force_basis"]
HEADERS = ["构件编号", "类型", "工况", "X (m)", "Y (m)", "标高 (m)", "有效长度 (m)", "轴力 (kN)", "弯矩 (kN·m)", "剪力 (kN)", "内力性质", "计量方式"]
NUMERIC = set(FIELDS[3:10])


class SheetpileWorker(QThread):
    completed = Signal(dict)
    failed = Signal(str)

    def __init__(self, request, project_name):
        super().__init__()
        self.request, self.project_name = request, project_name

    def run(self):
        try:
            self.completed.emit(generate_sheetpile(self.request, ROOT / 'templates/sheetpile-structure.local.json', ROOT / 'outputs/sheetpile', self.project_name))
        except Exception as exc:
            self.failed.emit(str(exc))


def configure_fonts():
    app = QApplication.instance()
    if app is None:
        return
    font = Path(os.environ.get("SystemRoot", "C:/Windows")) / "Fonts" / "msyh.ttc"
    if font.is_file() and "Microsoft YaHei" not in QFontDatabase.families():
        QFontDatabase.addApplicationFont(str(font))
    if "Microsoft YaHei" in QFontDatabase.families():
        app.setFont(QFont("Microsoft YaHei", 10))


class PlanPreview(QWidget):
    def __init__(self):
        super().__init__()
        self.project = new_project()
        self.setMinimumHeight(230)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), QColor("#f5f8fc"))
        p.setPen(QColor("#334155"))
        p.drawText(20, 26, "平面位置示意（地面标高为0；非施工图）")
        length, width = self.project["length_m"], self.project["width_m"]
        if length <= 0 or width <= 0:
            p.drawText(20, 65, "输入基坑长宽后可预览构件位置")
            return
        scale = min((self.width() - 100) / length, (self.height() - 100) / width)
        x0, y0 = 50, 60
        p.setPen(QPen(QColor("#2563eb"), 3))
        p.drawRect(x0, y0, int(length * scale), int(width * scale))
        p.setPen(QColor("#334155"))
        p.drawText(x0, y0 - 10, f"{length:g} × {width:g} m")
        seen = set()
        for m in self.project["members"]:
            key = (m["member_id"], m["x_m"], m["y_m"], m["level_m"])
            if key in seen:
                continue
            seen.add(key)
            x, y = x0 + m["x_m"] * scale, y0 + m["y_m"] * scale
            p.drawEllipse(int(x - 3), int(y - 3), 6, 6)
            p.drawText(int(x + 6), int(y - 4), f'{m["member_id"]} / {m["level_m"]:g}m')


class Workbench(QMainWindow):
    def __init__(self, database=None):
        configure_fonts()
        super().__init__()
        self.setWindowTitle("ELS 基坑支护模板工作台 · v0.2")
        self.resize(1250, 830)
        self.store = ProjectStore(database or ROOT / "data" / "projects.sqlite3")
        self.project = new_project()
        outer = QWidget()
        layout = QVBoxLayout(outer)
        title = QLabel("基坑支护 · 计算书与图纸模板工作台")
        title.setStyleSheet("font-size:22px;font-weight:600;padding:8px")
        layout.addWidget(title)
        self.state = QLabel("模板尚未接入：可录入和保存项目，暂不能生成工程设计成果")
        self.state.setStyleSheet("background:#fff1d6;color:#714800;padding:10px")
        layout.addWidget(self.state)
        bar = QHBoxLayout()
        for label, action in [("新建项目", self.new), ("保存版本", self.save), ("历史版本", self.history), ("导入项目", self.import_file), ("导出项目", self.export_file), ("备份数据库", self.backup), ("检查输入", self.check)]:
            b = QPushButton(label)
            b.clicked.connect(action)
            bar.addWidget(b)
        layout.addLayout(bar)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        self.setCentralWidget(outer)
        basic = QWidget()
        form = QFormLayout(basic)
        self.name = QLineEdit()
        form.addRow("项目名称", self.name)
        self.geometry = {}
        for key, label in [("length_m", "基坑长度 (m)"), ("width_m", "基坑宽度 (m)"), ("depth_m", "开挖深度 (m)"), ("embedment_m", "钢板桩入土深度 (m)")]:
            spin = QDoubleSpinBox()
            spin.setRange(0, 10000)
            spin.setDecimals(3)
            self.geometry[key] = spin
            form.addRow(label, spin)
        self.analysis = QLineEdit()
        form.addRow("外部分析来源／布置版本", self.analysis)
        self.site = QTextEdit()
        self.site.setPlaceholderText("记录土层、地下水和周边环境；本工具不从土质名称推导内力")
        form.addRow("场地资料", self.site)
        self.external = QTextEdit()
        self.external.setPlaceholderText("填写整体稳定、抗隆起、渗流和变形的外部验算结论及来源")
        form.addRow("外部验算记录", self.external)
        tabs.addTab(basic, "项目与场地")
        forces = QWidget()
        f = QVBoxLayout(forces)
        f.addWidget(QLabel("一行对应一项构件／工况。轴力以压缩为正；弯矩和剪力保留符号；每延米时单位相应为kN/m、kN·m/m。"))
        self.table = QTableWidget(0, len(FIELDS))
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        f.addWidget(self.table)
        buttons = QHBoxLayout()
        for label, action in [("增加工况", self.add_row), ("删除选中行", self.delete_rows), ("刷新平面示意", self.refresh_preview)]:
            b = QPushButton(label)
            b.clicked.connect(action)
            buttons.addWidget(b)
        f.addLayout(buttons)
        self.preview = PlanPreview()
        f.addWidget(self.preview)
        tabs.addTab(forces, "构件与内力")
        templates = QWidget()
        t = QVBoxLayout(templates)
        t.addWidget(QLabel("保留原始模板。Excel检查仅读取结构，不执行宏、不重算公式、不修改文件。"))
        template_buttons = QHBoxLayout()
        for label, action in [("检查Excel模板", self.inspect), ("选择模板映射JSON", self.choose_manifest), ("检查生成准备情况", self.prepare)]:
            b = QPushButton(label)
            b.clicked.connect(action)
            template_buttons.addWidget(b)
        t.addLayout(template_buttons)
        self.manifest = QLineEdit()
        self.manifest.setPlaceholderText("审核后的模板映射路径；未接入时保持为空")
        t.addWidget(self.manifest)
        self.report = QTextEdit()
        self.report.setReadOnly(True)
        t.addWidget(self.report)
        tabs.addTab(templates, "模板接入")
        steel = QWidget()
        steel_layout = QVBoxLayout(steel)
        steel_layout.addWidget(QLabel('连续钢板桩墙：输入未乘1.4的模型内力，软件取绝对值包络，由Excel乘1.4一次。成果须工程师复核。'))
        sf = QFormLayout()
        self.steel_mode = QComboBox()
        self.steel_mode.addItem('手动指定型号', 'manual')
        self.steel_mode.addItem('选择最轻合格截面', 'lightest')
        sf.addRow('选型方式', self.steel_mode)
        self.steel_section = QComboBox()
        self.steel_section.addItems(['请选择'] + SECTIONS)
        sf.addRow('钢板桩型号', self.steel_section)
        self.steel_fields = {}
        for key, label in [('yield_strength_mpa', '自动材料强度 (MPa)'), ('shear_reduction', '设计折减系数（剪力／弯矩共用）'), ('stiffness_reduction', 'EI折减系数 Bd'), ('analysis_ei_knm2_per_m', '自动折减EI (kN·m²/m)'), ('source', '钢板桩分析文件／版本'), ('parameter_reference', '自动参数来源')]:
            edit = QLineEdit()
            edit.setPlaceholderText('必填，未提供时不使用旧项目默认值')
            self.steel_fields[key] = edit
            sf.addRow(label, edit)
            if key in ('yield_strength_mpa', 'analysis_ei_knm2_per_m', 'parameter_reference'):
                edit.setReadOnly(True)
                edit.setStyleSheet('background:#eef3f8;')
        self.steel_fields['bending_reduction'] = self.steel_fields['shear_reduction']
        self.steel_properties = QTableWidget(3, 8)
        self.steel_properties.setHorizontalHeaderLabels(['范围', 'w (mm)', 'h (mm)', 't (mm)', 'A', 'I', 'Z', 'm'])
        self.steel_properties.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.steel_properties.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.steel_properties.setMaximumHeight(135)
        steel_layout.addLayout(sf)
        steel_layout.addWidget(self.steel_properties)
        steel_layout.addWidget(QLabel('每根：A cm²、I cm⁴、Z cm³、m kg/m；每米墙：A cm²/m、I cm⁴/m、Z cm³/m、m kg/m²。\nE=205 kN/mm²；分析模型直接采用Excel生成的EI，无需重复核对。'))
        self.steel_cases = QTableWidget(0, 3)
        self.steel_cases.setHorizontalHeaderLabels(['工况名称', '模型剪力 (kN/m)', '模型弯矩 (kN·m/m)'])
        self.steel_cases.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        steel_layout.addWidget(self.steel_cases)
        sb = QHBoxLayout()
        for label, action in [('增加钢板桩工况', self.add_steel_case), ('删除钢板桩选中行', self.delete_steel_cases)]:
            button = QPushButton(label)
            button.clicked.connect(action)
            sb.addWidget(button)
        self.steel_generate = QPushButton('生成计算书与PDF（待复核）')
        self.steel_generate.clicked.connect(self.generate_steel)
        sb.addWidget(self.steel_generate)
        steel_layout.addLayout(sb)
        self.steel_result = QTextEdit()
        self.steel_result.setReadOnly(True)
        steel_layout.addWidget(self.steel_result)
        tabs.addTab(steel, '钢板桩结构计算')
        self.steel_worker = None
        for widget in self.steel_fields.values():
            widget.textChanged.connect(self.steel_changed)
        self.steel_mode.currentIndexChanged.connect(self.steel_changed)
        self.steel_section.currentIndexChanged.connect(self.steel_changed)
        self.steel_section.currentIndexChanged.connect(self.update_steel_properties)
        self.steel_fields['stiffness_reduction'].textChanged.connect(self.update_steel_properties)
        self.steel_cases.itemChanged.connect(self.steel_changed)
        self.name.textChanged.connect(self.steel_changed)
        self.load_form(self.project)

    def guarded(self, action):
        try:
            return action()
        except Exception as exc:
            QMessageBox.warning(self, "操作未完成", str(exc))

    def collect(self):
        p = dict(self.project)
        p.update(name=self.name.text().strip(), analysis_reference=self.analysis.text().strip(), site_notes=self.site.toPlainText(), external_checks=self.external.toPlainText(), template_manifest=self.manifest.text().strip())
        p.update({k: widget.value() for k, widget in self.geometry.items()})
        p["members"] = []
        p['sheetpile'] = self.collect_steel()
        for row in range(self.table.rowCount()):
            member = {}
            for col, key in enumerate(FIELDS):
                widget = self.table.cellWidget(row, col)
                text = widget.currentText() if widget else (self.table.item(row, col).text() if self.table.item(row, col) else "")
                try:
                    member[key] = float(text) if key in NUMERIC else text.strip()
                except ValueError:
                    raise ValueError(f"第{row + 1}行 {HEADERS[col]}须填写数字")
            p["members"].append(member)
        return p

    def load_form(self, p):
        self.project = p
        self.name.setText(p["name"])
        self.analysis.setText(p["analysis_reference"])
        self.site.setPlainText(p["site_notes"])
        self.external.setPlainText(p["external_checks"])
        self.manifest.setText(p["template_manifest"])
        for k, widget in self.geometry.items():
            widget.setValue(p[k])
        self.table.setRowCount(0)
        for member in p["members"]:
            self.add_row(member)
        self.preview.project = p
        self.preview.update()
        self.state.setText("项目已载入；输入与模板尚需校验，未生成设计成果")
        self.load_steel(p.get('sheetpile', {}))

    def collect_steel(self):
        if self.steel_cases.rowCount() == 0 and self.steel_section.currentIndex() == 0 and self.steel_mode.currentIndex() == 0 and not any(self.steel_fields[k].text().strip() for k in ('shear_reduction', 'stiffness_reduction', 'source')):
            return {}
        r = sheetpile_draft()
        r['selection_mode'] = self.steel_mode.currentData()
        r['section'] = self.steel_section.currentText() if self.steel_section.currentIndex() else ''
        for key, widget in self.steel_fields.items():
            text = widget.text().strip()
            if key in ('source', 'parameter_reference'):
                r[key] = text
            else:
                try:
                    r[key] = float(text) if text else None
                except ValueError:
                    raise ValueError(f'钢板桩{key}须为数字')
        for row in range(self.steel_cases.rowCount()):
            cells = [self.steel_cases.item(row, col).text().strip() if self.steel_cases.item(row, col) else '' for col in range(3)]
            try:
                r['cases'].append({'name': cells[0], 'v_kn_per_m': float(cells[1]) if cells[1] else None, 'm_knm_per_m': float(cells[2]) if cells[2] else None})
            except ValueError:
                raise ValueError(f'钢板桩第{row+1}工况内力须为数字')
        if getattr(self, 'steel_ei_origin', None) == 'section_table':
            r['analysis_ei_origin'] = 'section_table'
            r['section_table_version'] = TABLE_VERSION
        return r

    def load_steel(self, request):
        self.steel_loading = True
        r = sheetpile_draft()
        r.update(request or {})
        self.steel_mode.setCurrentIndex(1 if r['selection_mode'] == 'lightest' else 0)
        self.steel_section.setCurrentText(r['section'] or '请选择')
        for key, widget in self.steel_fields.items():
            if key == 'bending_reduction':
                continue
            widget.setText('' if r[key] is None else str(r[key]))
        if r['shear_reduction'] != r['bending_reduction']:
            self.steel_fields['shear_reduction'].clear()
        self.steel_fields['yield_strength_mpa'].setText(str(YIELD_MPA))
        if not self.steel_fields['parameter_reference'].text():
            self.steel_fields['parameter_reference'].setText('工程师附表；E=205 kN/mm²；fy=275 MPa（已确认）；两项折减系数按项目输入')
        self.steel_ei_origin = request.get('analysis_ei_origin') if request else None
        self.steel_loading = False
        if r['analysis_ei_knm2_per_m'] is None or self.steel_ei_origin == 'section_table':
            self.update_steel_properties()
        else:
            self.update_steel_properties(preserve_model_ei=True)
        self.steel_cases.setRowCount(0)
        for case in r['cases']:
            self.add_steel_case(case)
        self.steel_result.clear()

    def update_steel_properties(self, *args, preserve_model_ei=False):
        if getattr(self, 'steel_loading', False):
            return
        self.steel_properties.clearContents()
        section = self.steel_section.currentText()
        if section not in SECTIONS:
            if not preserve_model_ei:
                self.steel_fields['analysis_ei_knm2_per_m'].clear()
            return
        p = properties(section)
        rows = [('尺寸', p['width_mm'], p['height_mm'], p['thickness_mm'], '', '', '', ''),
                ('每根', '', '', '', p['pile_area_cm2'], p['pile_inertia_cm4'], p['pile_modulus_cm3'], p['pile_mass_kg_per_m']),
                ('每米墙', '', '', '', p['wall_area_cm2_per_m'], p['wall_inertia_cm4_per_m'], p['wall_modulus_cm3_per_m'], p['wall_mass_kg_per_m2'])]
        for row, values in enumerate(rows):
            for col, value in enumerate(values):
                self.steel_properties.setItem(row, col, QTableWidgetItem(f'{value:g}' if isinstance(value, (int, float)) else value))
        if not preserve_model_ei:
            self.steel_ei_origin = 'section_table'
            try:
                bd = float(self.steel_fields['stiffness_reduction'].text())
                ei = properties(section, bd)['reduced_ei_knm2_per_m']
                self.steel_fields['analysis_ei_knm2_per_m'].setText(f'{ei:.12g}')
            except ValueError:
                self.steel_fields['analysis_ei_knm2_per_m'].clear()

    def add_steel_case(self, case=None):
        case = case if isinstance(case, dict) else {}
        row = self.steel_cases.rowCount()
        self.steel_cases.insertRow(row)
        for col, key in enumerate(('name', 'v_kn_per_m', 'm_knm_per_m')):
            value = case.get(key)
            self.steel_cases.setItem(row, col, QTableWidgetItem('' if value is None else str(value)))
        self.steel_changed()

    def delete_steel_cases(self):
        for row in sorted({i.row() for i in self.steel_cases.selectedIndexes()}, reverse=True):
            self.steel_cases.removeRow(row)
        self.steel_changed()

    def steel_changed(self, *args):
        if hasattr(self, 'steel_result') and self.steel_result.toPlainText():
            self.steel_result.setPlainText('输入已修改，原生成结果已过期；请重新生成。')

    def generate_steel(self):
        def action():
            request = self.collect_steel()
            errors = validate_request(request)
            if errors:
                raise ValueError('\n'.join(errors))
            if not self.name.text().strip():
                raise ValueError('请填写项目名称')
            self.steel_generate.setEnabled(False)
            self.steel_result.setPlainText('正在使用Excel重算并生成独立副本…')
            self.steel_worker = SheetpileWorker(request, self.name.text().strip())
            self.steel_worker.completed.connect(self.steel_completed)
            self.steel_worker.failed.connect(self.steel_failed)
            self.steel_worker.finished.connect(lambda: self.steel_generate.setEnabled(True))
            self.steel_worker.start()
        self.guarded(action)

    def steel_completed(self, result):
        try:
            if self.collect_steel() != result['request'] or self.name.text().strip() != result['project_name']:
                self.steel_result.setPlainText('生成期间输入已修改，该成果已过期。\n' + result['workbook'])
                return
        except ValueError:
            self.steel_result.setPlainText('生成期间输入已修改，该成果已过期。')
            return
        s = result['selected']
        self.steel_result.setPlainText(f"型号：{s['section']}\n剪力利用率：{s['shear_utilization']:.3f}\n弯矩利用率：{s['moment_utilization']:.3f}\n" + ('型号或刚度与分析基准不同：须重新分析。\n' if s['requires_reanalysis'] else '分析模型与计算书共用Excel生成的EI。\n') + f"Excel：{result['workbook']}\nPDF：{result['pdf']}\n仅连续墙剪力及弯矩，不代表整体稳定、节点或图纸已完成。")

    def steel_failed(self, message):
        self.steel_result.setPlainText('生成未完成：\n' + message)

    def closeEvent(self, event):
        if self.steel_worker is not None and self.steel_worker.isRunning():
            QMessageBox.information(self, 'Excel正在生成', '请等待本次生成结束后再关闭窗口。')
            event.ignore()
        else:
            event.accept()

    def add_row(self, member=None):
        if not isinstance(member, dict):
            member = {}
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col, key in enumerate(FIELDS):
            choices = {"kind": ["钢板桩", "围檩", "支撑"], "value_basis": ["请选择", "特征值", "设计值"], "force_basis": ["请选择", "每延米", "单根"]}.get(key)
            if choices:
                widget = QComboBox()
                widget.addItems(choices)
                widget.setCurrentText(str(member.get(key, choices[0])))
                self.table.setCellWidget(row, col, widget)
            else:
                self.table.setItem(row, col, QTableWidgetItem(str(member.get(key, ""))))

    def delete_rows(self):
        for row in sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(row)

    def new(self):
        answer = QMessageBox.question(self, "新建项目", "未保存的修改将丢失，确定新建？")
        if answer == QMessageBox.StandardButton.Yes:
            self.load_form(new_project())

    def save(self):
        def action():
            p = self.collect()
            errors = validate(p)
            revision = self.store.save(p)
            self.project = p
            self.state.setText(f"已保存版本 {revision} · {'草稿（输入待补充）' if errors else '输入完整；工程计算尚未验证'}")
        self.guarded(action)

    def history(self):
        rows = self.store.list_revisions()
        if not rows:
            QMessageBox.information(self, "历史版本", "尚无已保存项目")
            return
        labels = [f"{r[2]} · v{r[1]} · {r[3]} · {r[0][:8]}" for r in rows]
        label, ok = QInputDialog.getItem(self, "打开历史版本（会替换未保存输入）", "选择版本", labels, 0, False)
        if ok:
            r = rows[labels.index(label)]
            self.guarded(lambda: self.load_form(self.store.load(r[0], r[1])))

    def import_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入项目（会替换未保存输入）", "", "ELS项目 (*.json)")
        if path:
            self.guarded(lambda: self.load_form(import_project(path)))

    def export_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出输入完整的项目", "project.json", "ELS项目 (*.json)")
        if path:
            def action():
                p = self.collect()
                errors = validate(p)
                if errors:
                    raise ValueError("请先补全输入；草稿可通过保存版本保留：\n" + "\n".join(errors))
                export_project(p, path)
            self.guarded(action)

    def backup(self):
        path, _ = QFileDialog.getSaveFileName(self, "备份数据库", "projects-backup.sqlite3", "SQLite (*.sqlite3)")
        if path:
            self.guarded(lambda: self.store.backup(path))

    def check(self):
        def action():
            errors = validate(self.collect())
            QMessageBox.information(self, "输入检查", "\n".join(errors) if errors else "输入格式检查通过。此结果不代表构件验算或规范审核通过。")
        self.guarded(action)

    def refresh_preview(self):
        def action():
            p = self.collect()
            errors = validate(p)
            if errors:
                raise ValueError("\n".join(errors))
            self.preview.project = p
            self.preview.update()
        self.guarded(action)

    def inspect(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择原始Excel模板", str(ROOT / "templates"), "Excel模板 (*.xlsx *.xlsm)")
        if path:
            self.guarded(lambda: self.report.setPlainText(json.dumps(inspect_excel(path), ensure_ascii=False, indent=2)))

    def choose_manifest(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择模板映射", str(ROOT / "templates"), "JSON (*.json)")
        if path:
            self.manifest.setText(path)

    def prepare(self):
        self.guarded(lambda: self.report.setPlainText("待处理事项：\n\n" + "\n".join(readiness(self.collect()))))


def main():
    app = QApplication.instance() or QApplication([])
    app.setStyle("Fusion")
    app.setStyleSheet("QPushButton { padding:7px; } QLineEdit,QDoubleSpinBox { padding:5px; } QTabWidget::pane { border:1px solid #cbd5e1; }")
    window = Workbench()
    window.show()
    return app.exec()
