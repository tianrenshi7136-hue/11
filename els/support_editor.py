"""Separate strut/waler entry, searchable sections and repeatable member rows."""
import copy
import re
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QWidget,QVBoxLayout,QHBoxLayout,QLabel,QLineEdit,QCheckBox,
 QTableWidget,QTableWidgetItem,QComboBox,QCompleter,QHeaderView,QPushButton,QTabWidget,
 QFormLayout,QTextEdit,QInputDialog,QMessageBox,QGroupBox,QScrollArea,QStyledItemDelegate,QAbstractItemView,QApplication,QProgressBar)
from .connections import AA_CONNECTIONS, BRACKET_FIELDS

STRUT_FIELDS=('member_id','case','section','grade','anchor_kn','spacing_m','angle_deg','strut_length_m','lx_m','ly_m','lx_equals_ly','kx','ky','klt','use_default_k')
WALER_FIELDS=('member_id','case','section','grade','source_strut','anchor_kn_per_m','load_width_m','load_span_m','lx_m','ly_m','lx_equals_ly','kx','ky','klt','use_default_k')
WELD_FIELDS=('member_id','case','kind','leg_mm','strength_mpa','effective_length_mm','distribution_length_m')
TEXT_FIELDS={'member_id','case','section','grade','source_strut','kind','bracket_id','waler_id','waler_case','load_model'}

class EngineeringNumbers(QStyledItemDelegate):
 def __init__(self,fields,parent=None):super().__init__(parent);self.fields=fields
 def displayText(self,value,locale):
  return str(value)
 def initStyleOption(self,option,index):
  super().initStyleOption(option,index)
  key=self.fields[index.column()]
  if key in TEXT_FIELDS or key=='lx_equals_ly':return
  try:number=float(index.data())
  except (ValueError,TypeError):return
  digits=3 if key.endswith('_m') else 1 if key in ('angle_deg','leg_mm','effective_length_mm') else 0 if key=='strength_mpa' else 2
  option.text=f'{number:.{digits}f}'

class SupportEditor(QWidget):
 changed=Signal()
 generate_requested=Signal()
 reference_requested=Signal()
 def __init__(self,sections,parent=None):
  super().__init__(parent);self.sections=sections;self.busy=False;self._multi_case=False;self._case_synced='LC1'
  self.setObjectName('SupportEditor')
  self.setStyleSheet("QGroupBox {font-weight:600; margin-top:12px; padding:16px 12px 12px;} QGroupBox::title {subcontrol-origin:margin; left:12px;} QHeaderView::section {padding:8px 6px; font-weight:600;} QLineEdit,QComboBox {min-height:28px;} QPushButton {min-height:30px; padding:3px 12px;} QPushButton#Generate {background:#2563eb; color:white; border:0; border-radius:5px; min-height:36px;}")
  layout=QVBoxLayout(self);layout.setContentsMargins(16,12,16,12);layout.setSpacing(12)
  header=QGroupBox('工况与分析来源');form=QFormLayout(header);form.setSpacing(8)
  self.case_name=QLineEdit('LC1');self.case_name.setPlaceholderText('本页统一工况，例如 ULS1')
  self.source=QLineEdit();self.source.setPlaceholderText('填写 PLAXIS / FREW 分析文件及版本')
  form.addRow('统一工况',self.case_name);form.addRow('内力来源',self.source);layout.addWidget(header)
  self.case_notice=QLabel();self.case_notice.setWordWrap(True);self.case_notice.hide();layout.addWidget(self.case_notice)
  toolbar=QHBoxLayout();self.generate=QPushButton('生成计算书及 PDF');self.generate.setObjectName('Generate');self.generate.clicked.connect(lambda *_:self.generate_requested.emit());toolbar.addWidget(self.generate,1)
  refresh=QPushButton('更新参考模板…');refresh.clicked.connect(lambda *_:self.reference_requested.emit());toolbar.addWidget(refresh);layout.addLayout(toolbar)
  self.generation_status=QLabel();self.generation_status.setWordWrap(True);self.generation_status.hide();layout.addWidget(self.generation_status)
  self.generation_bar=QProgressBar();self.generation_bar.setRange(0,0);self.generation_bar.setTextVisible(False);self.generation_bar.setMaximumHeight(6);self.generation_bar.hide();layout.addWidget(self.generation_bar)
  self.combined=QCheckBox('包含钢板桩计算书（使用钢板桩页面已输入的数据）');self.combined.setChecked(True);layout.addWidget(self.combined)
  self.tabs=QTabWidget();self.tabs.setDocumentMode(True);layout.addWidget(self.tabs,1)
  self.struts=self.member_page('Strut',STRUT_FIELDS,['编号','工况','截面（输入搜索）','钢材','墙法线力\nkN（未分项）','支撑间距\nm','角度\ndeg','Strut length\nm','Lx\nm','Ly\nm','长度联动','Kx','Ky','KLT','K 系数设置'])
  self.walers=self.member_page('Waler',WALER_FIELDS,['编号','工况','截面（输入搜索）','钢材','关联支撑','分布荷载\nkN/m（自动）','荷载宽度\nm（自动）','Span\nm','Lx\nm','Ly\nm','长度联动','Kx','Ky','KLT','K 系数设置'])
  # Legacy rows remain available for import/backward compatibility, not manual entry.
  self.welds=QTableWidget(0,len(WELD_FIELDS));self.welds.hide();self.welds.itemChanged.connect(self.emit_changed)
  page=QWidget();pl=QVBoxLayout(page);pl.setContentsMargins(12,12,12,12);pl.setSpacing(14)
  note=QLabel('每类连接填写一组参数。支撑／角撑—围檩焊缝取关联围檩的最大设计剪力 V，支撑截面用于确定焊长；计算书显示控制构件及工况，无需填写连接构件编号。');note.setWordWrap(True);pl.addWidget(note)
  cards=QHBoxLayout();cards.setSpacing(18);self.connection_controls={}
  for kind,title in [('Strut-Waler','支撑 / 角撑 — 围檩'),('Waler-Sheetpile','围檩 — 钢板桩')]:
   box=QGroupBox(title);bf=QFormLayout(box);bf.setSpacing(12);controls={};enabled=QCheckBox('按关联围檩最大设计剪力检查' if kind=='Strut-Waler' else '自动检查最大设计剪力');controls['enabled']=enabled;bf.addRow(enabled)
   fields=[('leg_mm','焊脚（mm）'),('strength_mpa','焊缝强度（MPa）')]+([('effective_length_mm','每米墙宽总有效焊长（mm/m）'),('distribution_length_m','统一分配长度（m）')] if kind=='Waler-Sheetpile' else [])
   for key,label in fields:
    field=QLineEdit();field.setPlaceholderText('请输入');field.textChanged.connect(lambda *_:self.changed.emit());controls[key]=field;bf.addRow(label,field)
   if kind=='Waler-Sheetpile':
    mode=QComboBox();mode.addItems(['采用各围檩 Span','手动统一分配长度']);controls['distribution_mode']=mode;bf.insertRow(3,'荷载分配方式',mode)
    mode.currentIndexChanged.connect(lambda *_:self.update_connection_controls());controls['distribution_length_m'].setEnabled(False)
   enabled.toggled.connect(lambda *_:self.update_connection_controls());self.connection_controls[kind]=controls;cards.addWidget(box,1)
  pl.addLayout(cards)
  bracket_box=QGroupBox('托架 — 钢板桩焊缝');bracket_form=QFormLayout(bracket_box);bracket_form.setSpacing(12);self.bracket_weld_controls={}
  bracket_form.addRow(QLabel('以下参数统一用于 Bracket 页已录入的托架；无需重复填写托架编号。'))
  for key,label in [('leg_mm','焊脚（mm）'),('strength_mpa','焊缝强度（MPa）'),('effective_length_mm','总有效焊长（mm）')]:
   field=QLineEdit();field.setPlaceholderText('请输入');field.textChanged.connect(lambda *_:self.changed.emit());self.bracket_weld_controls[key]=field;bracket_form.addRow(label,field)
  pl.addWidget(bracket_box);self.legacy_notice=QLabel();self.legacy_notice.setWordWrap(True);pl.addWidget(self.legacy_notice);pl.addStretch();scroll=QScrollArea();scroll.setWidgetResizable(True);scroll.setWidget(page);self.tabs.addTab(scroll,'Weld 连接验算')
  bp=QWidget();bl=QVBoxLayout(bp);bl.setContentsMargins(12,12,12,12);bl.setSpacing(12)
  label=QLabel('悬臂点荷载：托架承担其间距范围内的围檩自重及自身自重；围檩合力作用点固定为托架长度 L/2。焊缝参数在 Weld 页填写，工况沿用页面顶部设置。');label.setWordWrap(True);bl.addWidget(label)
  self.brackets=QTableWidget(0,len(BRACKET_FIELDS));self.brackets.setHorizontalHeaderLabels(['托架编号','工况','关联围檩','围檩工况','截面（输入搜索）','钢材','托架长度\nm','托架间距\nm','受力模型','围檩重心距墙\nm','焊脚\nmm','焊缝强度\nMPa','总有效焊长\nmm']);self.style_table(self.brackets,BRACKET_FIELDS);bl.addWidget(self.brackets,1)
  for key in ('case','waler_case','load_model','eccentricity_m','leg_mm','strength_mpa','effective_length_mm'):self.brackets.setColumnHidden(BRACKET_FIELDS.index(key),True)
  buttons=QHBoxLayout()
  for title,action in [('增加托架',lambda:self.add_bracket()),('删除选中托架',lambda:self.delete_rows(self.brackets))]:
   button=QPushButton(title);button.clicked.connect(action);buttons.addWidget(button)
  buttons.addStretch();bl.addLayout(buttons);self.brackets.itemChanged.connect(self.emit_changed);self.tabs.addTab(bp,'Bracket 托架')
  cp=QWidget();cl=QVBoxLayout(cp);cl.setContentsMargins(16,16,16,16);cg=QGroupBox('Blinding concrete');cf=QFormLayout(cg);cf.setSpacing(12)
  self.include_blinding=QCheckBox('生成 Blinding 验算（使用钢板桩页的未分项剪力）');cf.addRow(self.include_blinding);self.thickness=QLineEdit();self.thickness.setPlaceholderText('输入厚度，单位 mm');cf.addRow('混凝土厚度（mm）',self.thickness);cf.addRow(QLabel('fcu 沿用参考方法 20 MPa；有效长度取基坑短边。'));cl.addWidget(cg);cl.addStretch();self.tabs.addTab(cp,'Blinding 垫层')
  self.result=QTextEdit();self.result.setReadOnly(True);self.result.setMaximumHeight(90);layout.addWidget(self.result)
  self.result.hide();self.result.textChanged.connect(lambda:self.result.setVisible(bool(self.result.toPlainText().strip())))
  self.thickness.textEdited.connect(lambda *_:self.include_blinding.setChecked(True))
  for field in (self.source,self.thickness,self.case_name):field.textChanged.connect(lambda *_:self.changed.emit())
  self.case_name.editingFinished.connect(self.apply_case)
  self.combined.toggled.connect(lambda *_:self.changed.emit());self.include_blinding.toggled.connect(lambda *_:self.changed.emit())
  self.update_connection_controls()
  self.apply_editor_style()
 def generation_started(self):
  self.generation_status.setText('准备生成计算书…');self.generation_status.show();self.generation_bar.show()
 def generation_finished(self):
  self.generation_status.hide();self.generation_bar.hide()
 def apply_editor_style(self):
  from .theme import workbench_style
  self.setStyleSheet(workbench_style())
  for index in range(self.tabs.count()):
   page=self.tabs.widget(index);page=page.widget() if isinstance(page,QScrollArea) else page;page.setProperty('surface',True)
 def style_table(self,table,fields):
  table.setAlternatingRowColors(True);table.verticalHeader().setDefaultSectionSize(38);table.horizontalHeader().setMinimumHeight(58);table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive);table.horizontalHeader().setDefaultSectionSize(108);table.horizontalHeader().setMinimumSectionSize(65)
  table.setItemDelegate(EngineeringNumbers(fields,table));table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
  for index,key in enumerate(fields):
   table.setColumnWidth(index,175 if key in ('section','source_strut') else 155 if key=='use_default_k' else 135 if key=='lx_equals_ly' else 95 if key in ('member_id','bracket_id','grade') else 75 if key in ('kx','ky','klt') else 110)
 def update_connection_controls(self):
  for controls in self.connection_controls.values():
   for key,field in controls.items():
    if key=='enabled':continue
    field.setEnabled(key!='distribution_length_m' or controls['distribution_mode'].currentIndex()==1)
    if key=='distribution_length_m':field.setPlaceholderText('采用各围檩 Span 自动填写' if controls['distribution_mode'].currentIndex()==0 else '请输入统一分配长度')
  self.changed.emit()
 def collect_groups(self):
  groups=[]
  for kind,controls in self.connection_controls.items():
   group=dict(kind=kind,enabled=controls['enabled'].isChecked())
   for key,field in controls.items():
    if key=='enabled':continue
    if key=='distribution_mode':group[key]='span' if field.currentIndex()==0 else 'manual';continue
    value=field.text().strip()
    try:group[key]=float(value) if value else None
    except ValueError:
     if group['enabled']:raise ValueError(f'{kind} {key}须为数字')
     group[key]=value
   groups.append(group)
  return groups
 def apply_case(self):
  if self.busy:return
  case=self.case_name.text().strip()
  if not case or case==self._case_synced:return
  self.busy=True
  for table,fields in [(self.struts,STRUT_FIELDS),(self.walers,WALER_FIELDS),(self.brackets,BRACKET_FIELDS),(self.welds,WELD_FIELDS)]:
   for row in range(table.rowCount()):
    table.item(row,fields.index('case')).setText(case)
    if 'waler_case' in fields:table.item(row,fields.index('waler_case')).setText(case)
    if 'waler_id' in fields:
     combo=table.cellWidget(row,fields.index('waler_id'));parts=combo.currentText().split(' / ')
     if len(parts)==2:combo.clear();combo.addItem(parts[0]+' / '+case)
    if 'source_strut' in fields:
     combo=table.cellWidget(row,fields.index('source_strut'));old=combo.currentText();parts=old.split(' / ')
     if len(parts)==2:combo.setProperty('restore_value',parts[0]+' / '+case);combo.setCurrentText('选择关联支撑')
  self._case_synced=case;self._multi_case=False;self.case_notice.hide();self.busy=False;self.refresh();self.changed.emit()
  for table,fields in [(self.struts,STRUT_FIELDS),(self.walers,WALER_FIELDS),(self.brackets,BRACKET_FIELDS)]:table.setColumnHidden(fields.index('case'),True)
 def member_page(self,kind,fields,headers):
  page=QWidget();layout=QVBoxLayout(page)
  note=QLabel('Strut：墙法线力按kN输入；角度为支撑轴线与墙法线的夹角，角撑轴力除以cos(角度)。Strut length用于自重验算，Lx默认跟随该长度。' if kind=='Strut' else 'Waling：选择关联支撑，荷载宽度自动取其间距，分布荷载由支撑力／间距得到；Span手动输入。');note.setWordWrap(True);layout.addWidget(note)
  table=QTableWidget(0,len(fields));table.setHorizontalHeaderLabels(headers);self.style_table(table,fields);table.setColumnHidden(fields.index('case'),True);layout.addWidget(table)
  buttons=QHBoxLayout();layout.addLayout(buttons)
  for label,action in [('增加支撑／角撑' if kind=='Strut' else '增加围檩',lambda checked=False,k=kind:self.add_member(k)),('按选中行批量新增',lambda checked=False,k=kind:self.bulk_dialog(k)),('删除选中行',lambda checked=False,t=table:self.delete_rows(t))]:
   b=QPushButton(label);b.clicked.connect(action);buttons.addWidget(b)
  table.itemChanged.connect(lambda item,k=kind,t=table:self.member_item_changed(k,t,item))
  self.tabs.addTab(page,'Struts 支撑／角撑' if kind=='Strut' else 'Waling 围檩')
  return table
 def emit_changed(self,*args):
  if not self.busy:
   self.refresh();self.changed.emit()
 def section_combo(self,value):
  combo=QComboBox();combo.addItems(['请选择']+self.sections);combo.setEditable(True);combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
  completer=QCompleter(self.sections,combo);completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive);completer.setFilterMode(Qt.MatchFlag.MatchContains);completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion);combo.setCompleter(completer)
  combo.setCurrentText(value or '请选择');combo.currentTextChanged.connect(self.emit_changed);return combo
 def add_member(self,kind,record=None):
  record=dict(record or {});table=self.struts if kind=='Strut' else self.walers;fields=STRUT_FIELDS if kind=='Strut' else WALER_FIELDS
  record.setdefault('case',self.case_name.text().strip() or 'LC1')
  record.setdefault('lx_equals_ly',True)
  record.setdefault('use_default_k',all(record.get(k,1) in (1,1.0,'1','1.0') for k in ('kx','ky','klt')))
  if not record.get('member_id'):record['member_id']=self.next_id(kind)
  row=table.rowCount();self.busy=True;table.insertRow(row)
  for col,key in enumerate(fields):
   value=record.get(key,False if key=='lx_equals_ly' else 1 if key in ('kx','ky','klt') else 0 if key=='angle_deg' else '')
   if key=='section':table.setCellWidget(row,col,self.section_combo(str(value)))
   elif key in ('grade','source_strut'):
    combo=QComboBox();combo.addItems(['请选择','S235','S275','S355','S460'] if key=='grade' else ['选择关联支撑']);combo.setCurrentText(str(value));combo.setProperty('restore_value',str(value));combo.currentTextChanged.connect(self.emit_changed);table.setCellWidget(row,col,combo)
   elif key in ('lx_equals_ly','use_default_k'):
    box=QCheckBox('Lx = Ly' if key=='lx_equals_ly' else '设置默认值');box.setChecked(bool(value));box.setToolTip('勾选后 Ly 跟随 Lx，不能单独修改' if key=='lx_equals_ly' else '勾选：Kx、Ky、KLT 固定为 1；取消勾选后可自行输入');box.toggled.connect(self.emit_changed);table.setCellWidget(row,col,box)
   else:table.setItem(row,col,QTableWidgetItem('' if value is None else str(value)))
  table.item(row,fields.index('lx_m')).setData(Qt.ItemDataRole.UserRole,record.get('lx_is_default',record.get('lx_m') in ('',None)))
  self.busy=False;self.refresh();self.changed.emit()
 def member_item_changed(self,kind,table,item):
  if self.busy:return
  fields=STRUT_FIELDS if kind=='Strut' else WALER_FIELDS
  if item.column()==fields.index('lx_m'):
   item.setData(Qt.ItemDataRole.UserRole,False)
  self.emit_changed()
 def next_id(self,kind):
  prefix='S' if kind=='Strut' else 'W';ids={self.text(t,r,0) for t in (self.struts,self.walers) for r in range(t.rowCount())};n=1
  while prefix+str(n) in ids:n+=1
  return prefix+str(n)
 def text(self,table,row,col):
  widget=table.cellWidget(row,col)
  if isinstance(widget,QComboBox):return widget.currentText().strip()
  item=table.item(row,col);return item.text().strip() if item else ''
 def read_rows(self,table,fields,strict=True):
  records=[]
  for row in range(table.rowCount()):
   record={}
   for col,key in enumerate(fields):
    widget=table.cellWidget(row,col)
    if isinstance(widget,QCheckBox):record[key]=widget.isChecked();continue
    text=self.text(table,row,col)
    if key in TEXT_FIELDS:record[key]=text
    else:
     try:record[key]=float(text) if text else None
     except ValueError:
      if strict:raise ValueError(f'第{row+1}行 {key}须为数字')
      record[key]=None
   records.append(record)
  return records
 def refresh(self):
  if self.busy:return
  self.busy=True
  try:
   for table,fields in ((self.struts,STRUT_FIELDS),(self.walers,WALER_FIELDS)):
    for row in range(table.rowCount()):
     lx_item=table.item(row,fields.index('lx_m'))
     if lx_item.data(Qt.ItemDataRole.UserRole):lx_item.setText(self.text(table,row,fields.index('strut_length_m' if table is self.struts else 'load_span_m')))
     linked=table.cellWidget(row,fields.index('lx_equals_ly')).isChecked();lx=self.text(table,row,fields.index('lx_m'));ly=table.item(row,fields.index('ly_m'))
     if linked:ly.setText(lx)
     ly.setFlags(ly.flags() & ~Qt.ItemFlag.ItemIsEditable if linked else ly.flags() | Qt.ItemFlag.ItemIsEditable)
     use_default=table.cellWidget(row,fields.index('use_default_k')).isChecked()
     for key in ('kx','ky','klt'):
      item=table.item(row,fields.index(key))
      if use_default:item.setText('1')
      item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable if use_default else item.flags() | Qt.ItemFlag.ItemIsEditable)
      item.setToolTip('默认值 1；取消“设置默认值”后可编辑' if use_default else '自定义 K 系数')
   struts={m['member_id']+' / '+m['case']:m for m in self.read_rows(self.struts,STRUT_FIELDS,False) if m['member_id'] and m['case']}
   for row in range(self.walers.rowCount()):
    combo=self.walers.cellWidget(row,WALER_FIELDS.index('source_strut'));chosen=combo.currentText();restored=combo.property('restore_value')
    if chosen=='选择关联支撑' and restored:chosen=restored
    choices=['选择关联支撑']+list(struts)
    if combo.count()!=len(choices) or [combo.itemText(i) for i in range(combo.count())]!=choices:
     combo.clear();combo.addItems(choices)
    if chosen not in choices and chosen not in ('','选择关联支撑'):combo.addItem(chosen if chosen.endswith('（已失效）') else chosen+'（已失效）');combo.setCurrentIndex(combo.count()-1)
    else:combo.setCurrentText(chosen if chosen in choices else choices[0])
    combo.setProperty('restore_value','')
    member=struts.get(combo.currentText());spacing=member.get('spacing_m') if member else None;force=member.get('anchor_kn') if member else None
    q=abs(force)/spacing if isinstance(force,(int,float)) and isinstance(spacing,(int,float)) and spacing>0 else None
    for key,value in [('anchor_kn_per_m',q),('load_width_m',spacing)]:
     item=self.walers.item(row,WALER_FIELDS.index(key));item.setText('' if value is None else f'{value:.12g}');item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
   walers={self.text(self.walers,r,0)+' / '+self.text(self.walers,r,1) for r in range(self.walers.rowCount())}
   for row in range(self.brackets.rowCount()):
    combo=self.brackets.cellWidget(row,BRACKET_FIELDS.index('waler_id'));old=combo.currentText();choices=['选择已定义围檩']+sorted(walers)
    invalid=old not in choices and old not in ('','选择已定义围檩')
    if invalid:choices.append(old)
    if [combo.itemText(i) for i in range(combo.count())]!=choices:combo.clear();combo.addItems(choices)
    combo.setCurrentText(old if old in choices else choices[0])
    if invalid:
     combo.model().item(combo.findText(old)).setEnabled(False);combo.setToolTip('原关联已失效，请选择现有围檩。')
    else:combo.setToolTip('选择已经定义的围檩编号及工况')
  finally:self.busy=False
 def duplicate(self,kind,count):
  table=self.struts if kind=='Strut' else self.walers;fields=STRUT_FIELDS if kind=='Strut' else WALER_FIELDS;row=table.currentRow()
  if row<0:raise ValueError('请先选中一行作为复制模板')
  record=self.read_rows(table,fields)[row]
  record['lx_is_default']=bool(table.item(row,fields.index('lx_m')).data(Qt.ItemDataRole.UserRole))
  welds=self.read_welds();created=[]
  for _ in range(count):
   clone=copy.deepcopy(record);clone['member_id']=self.next_id(kind);created.append(clone['member_id']);self.add_member(kind,clone)
   for weld in welds:
    if (weld['member_id'],weld['case'])==(record['member_id'],record['case']):
     w=copy.deepcopy(weld);w['member_id']=clone['member_id'];self.add_weld(w)
  return created
 def bulk_dialog(self,kind):
  count,ok=QInputDialog.getInt(self,'批量新增','新增数量（沿用选中行参数，编号自动递增）',1,1,100)
  if ok:
   try:self.duplicate(kind,count)
   except ValueError as e:QMessageBox.warning(self,'无法复制',str(e))
 def delete_rows(self,table):
  for row in sorted({i.row() for i in table.selectedIndexes()},reverse=True):table.removeRow(row)
  self.refresh();self.changed.emit()
 def add_weld(self,record=None):
  record=dict(record or {});row=self.welds.rowCount();self.busy=True;self.welds.insertRow(row)
  for col,key in enumerate(WELD_FIELDS):
   if key=='kind':
    combo=QComboBox();combo.addItems(['Strut-Waler','Waler-Sheetpile']);combo.setCurrentText(record.get(key,'Strut-Waler'));combo.currentTextChanged.connect(self.update_welds);self.welds.setCellWidget(row,col,combo)
   else:self.welds.setItem(row,col,QTableWidgetItem('' if record.get(key) is None else str(record.get(key,''))))
  self.busy=False;self.update_welds()
  if record:
   controls=self.connection_controls[record.get('kind','Strut-Waler')];controls['enabled'].setChecked(True)
   for key,field in controls.items():
    if key in ('enabled','distribution_mode'):continue
    value=record.get(key)
    if value is not None:field.setText(str(value))
   if 'distribution_mode' in controls:controls['distribution_mode'].setCurrentIndex(1 if record.get('distribution_length_m') else 0)
 def update_welds(self,*args):
  self.busy=True
  for row in range(self.welds.rowCount()):
   automatic=self.text(self.welds,row,2)=='Strut-Waler'
   for col in (5,6):
    item=self.welds.item(row,col)
    if automatic:item.setText('自动' if col==5 else '不适用');item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    else:
     if item.text() in ('自动','不适用'):item.setText('')
     item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
  self.busy=False;self.changed.emit()
 def read_welds(self):
  records=[]
  for row in range(self.welds.rowCount()):
   record={key:self.text(self.welds,row,col) for col,key in enumerate(WELD_FIELDS)}
   for key in WELD_FIELDS[3:]:
    if record['kind']=='Strut-Waler' and key in WELD_FIELDS[5:]:record[key]=None
    else:
     try:record[key]=float(record[key]) if record[key] else None
     except ValueError:raise ValueError(f'焊缝第{row+1}项 {key}须为数字')
   records.append(record)
  return records
 def add_bracket(self,record=None):
  record=dict(record or {});row=self.brackets.rowCount();self.busy=True;self.brackets.insertRow(row)
  existing={self.text(self.brackets,r,0) for r in range(row)};number=1
  while 'B'+str(number) in existing:number+=1
  record.setdefault('bracket_id','B'+str(number))
  record.setdefault('case',self.case_name.text().strip() or 'LC1');record.setdefault('waler_case',self.case_name.text().strip() or 'LC1')
  record.setdefault('load_model','point')
  for col,key in enumerate(BRACKET_FIELDS):
   value=record.get(key,'')
   if key=='section':self.brackets.setCellWidget(row,col,self.section_combo(value))
   elif key=='waler_id':
    combo=QComboBox();combo.addItems(['选择已定义围檩']+[self.text(self.walers,r,0)+' / '+self.text(self.walers,r,1) for r in range(self.walers.rowCount())]);chosen=str(value)+' / '+str(record.get('waler_case','')) if value else ''
    if chosen and combo.findText(chosen)<0:combo.addItem(chosen)
    combo.setCurrentText(chosen or '选择已定义围檩');combo.currentTextChanged.connect(self.emit_changed);self.brackets.setCellWidget(row,col,combo)
   elif key in ('grade','load_model'):
    combo=QComboBox();combo.addItems(['请选择','S235','S275','S355','S460'] if key=='grade' else ['选择受力模型','point']);combo.setCurrentText(str(value));combo.currentTextChanged.connect(self.emit_changed);self.brackets.setCellWidget(row,col,combo)
   else:self.brackets.setItem(row,col,QTableWidgetItem('' if value is None else str(value)))
  self.busy=False;self.changed.emit()
 def collect(self,length,width):
  self.apply_case()
  if not self.case_name.text().strip() and not self._multi_case:raise ValueError('请在页面顶部填写统一工况')
  self.refresh();members=[]
  for row,m in enumerate(self.read_rows(self.struts,STRUT_FIELDS)):
   m['kind']='Strut';m['lx_is_default']=bool(self.struts.item(row,STRUT_FIELDS.index('lx_m')).data(Qt.ItemDataRole.UserRole));members.append(m)
  for row,m in enumerate(self.read_rows(self.walers,WALER_FIELDS)):
   source=m.pop('source_strut');parts=source.split(' / ');m.update(kind='Waler',angle_deg=0.0,source_strut_id=parts[0] if len(parts)==2 else '',source_strut_case=parts[1] if len(parts)==2 else '');members.append(m)
   m['lx_is_default']=bool(self.walers.item(row,WALER_FIELDS.index('lx_m')).data(Qt.ItemDataRole.UserRole))
  try:thickness=float(self.thickness.text()) if self.thickness.text().strip() else None
  except ValueError:raise ValueError('Blinding厚度须为数字')
  brackets=self.read_rows(self.brackets,BRACKET_FIELDS)
  for b in brackets:
   parts=b['waler_id'].split(' / ');b['waler_id']=parts[0] if len(parts)==2 else '';b['waler_case']=parts[1] if len(parts)==2 else ''
   b['eccentricity_m']=b['length_m']/2 if isinstance(b.get('length_m'),(float,int)) else None
   for key,field in self.bracket_weld_controls.items():
    try:b[key]=float(field.text()) if field.text().strip() else None
    except ValueError:raise ValueError(f'托架焊缝 {key}须为数字')
  return dict(schema_version=2,case_name=self.case_name.text().strip(),connection_groups=self.collect_groups(),connection_template=AA_CONNECTIONS,brackets=brackets,enabled=True,include_sheetpile=self.combined.isChecked(),include_blinding=self.include_blinding.isChecked(),source=self.source.text().strip(),members=members,welds=self.read_welds(),length_m=length,width_m=width,blinding_thickness_mm=thickness)
 def load(self,record):
  self.struts.setRowCount(0);self.walers.setRowCount(0);self.welds.setRowCount(0);self.brackets.setRowCount(0)
  cases=list(dict.fromkeys(m.get('case','LC1') for m in record.get('members',[])))
  self._multi_case=len(cases)>1;case=record.get('case_name') or (cases[0] if len(cases)==1 else '' if cases else 'LC1');self._case_synced=case;self.case_name.setText(case)
  self.case_notice.setText('已保留历史多个工况。顶部输入统一工况后，才会统一替换；构件内力保留。' if self._multi_case else '');self.case_notice.setVisible(self._multi_case)
  for table,fields in [(self.struts,STRUT_FIELDS),(self.walers,WALER_FIELDS),(self.brackets,BRACKET_FIELDS)]:table.setColumnHidden(fields.index('case'),not self._multi_case)
  for controls in self.connection_controls.values():
   controls['enabled'].setChecked(False)
   for key,field in controls.items():
    if isinstance(field,QLineEdit):field.clear()
  self.combined.setChecked(record.get('include_sheetpile',True));self.source.setText(record.get('source',''));self.include_blinding.setChecked(record.get('include_blinding',record.get('blinding_thickness_mm') is not None));self.thickness.setText('' if record.get('blinding_thickness_mm') is None else str(record['blinding_thickness_mm']))
  for m in record.get('members',[]):
   if m['kind']=='Strut':self.add_member('Strut',m)
  for m in record.get('members',[]):
   if m['kind']=='Waler':
    row=dict(m);row['source_strut']=row.get('source_strut_id','')+' / '+row.get('source_strut_case','');self.add_member('Waler',row)
  for w in record.get('welds',[]):self.add_weld(w)
  for b in record.get('brackets',[]):self.add_bracket(b)
  for key,field in self.bracket_weld_controls.items():
   values={b.get(key) for b in record.get('brackets',[])};value=values.pop() if len(values)==1 else None
   field.setText('' if value is None else str(value))
  if any(len({b.get(key) for b in record.get('brackets',[])})>1 for key in self.bracket_weld_controls):
   self.legacy_notice.setText('历史托架的焊缝参数不一致，请在本页填写统一参数后生成。')
  groups=record.get('connection_groups')
  if groups is None:
   groups=[]
   for kind in self.connection_controls:
    rows=[w for w in record.get('welds',[]) if w.get('kind')==kind]
    if not rows:continue
    group=dict(kind=kind,enabled=True,distribution_mode='manual')
    for key in ('leg_mm','strength_mpa','effective_length_mm','distribution_length_m'):
     values={w.get(key) for w in rows};group[key]=values.pop() if len(values)==1 else None
    groups.append(group)
  for group in groups:
   controls=self.connection_controls[group['kind']];controls['enabled'].setChecked(group.get('enabled',False))
   for key,field in controls.items():
    if isinstance(field,QLineEdit):field.setText('' if group.get(key) is None else str(group[key]))
   if 'distribution_mode' in controls:controls['distribution_mode'].setCurrentIndex(0 if group.get('distribution_mode')=='span' else 1)
  self.update_connection_controls()
