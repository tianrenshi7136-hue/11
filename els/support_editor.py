"""Separate strut/waler entry, searchable sections and repeatable member rows."""
import copy
import re
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QWidget,QVBoxLayout,QHBoxLayout,QLabel,QLineEdit,QCheckBox,
 QTableWidget,QTableWidgetItem,QComboBox,QCompleter,QHeaderView,QPushButton,QTabWidget,
 QFormLayout,QTextEdit,QInputDialog,QMessageBox)
from .connections import AA_CONNECTIONS, BRACKET_FIELDS

STRUT_FIELDS=('member_id','case','section','grade','anchor_kn','spacing_m','angle_deg','strut_length_m','lx_m','ly_m','lx_equals_ly','kx','ky','klt')
WALER_FIELDS=('member_id','case','section','grade','source_strut','anchor_kn_per_m','load_width_m','load_span_m','lx_m','ly_m','lx_equals_ly','kx','ky','klt')
WELD_FIELDS=('member_id','case','kind','leg_mm','strength_mpa','effective_length_mm','distribution_length_m')
TEXT_FIELDS={'member_id','case','section','grade','source_strut','kind','bracket_id','waler_id','waler_case','load_model'}

class SupportEditor(QWidget):
 changed=Signal()
 generate_requested=Signal()
 def __init__(self,sections,parent=None):
  super().__init__(parent);self.sections=sections;self.busy=False
  layout=QVBoxLayout(self)
  self.combined=QCheckBox('在钢板桩计算书中同时加入围檩／支撑工作表');layout.addWidget(self.combined)
  self.source=QLineEdit();self.source.setPlaceholderText('PLAXIS／FREW分析文件及版本（必填）');layout.addWidget(self.source)
  self.tabs=QTabWidget();layout.addWidget(self.tabs)
  self.struts=self.member_page('Strut',STRUT_FIELDS,['编号','工况','截面（输入搜索）','钢材','墙法线力\nkN（未分项）','支撑间距\nm','角度\ndeg','Strut length\nm','Lx\nm','Ly\nm','Lx=Ly','Kx','Ky','KLT'])
  self.walers=self.member_page('Waler',WALER_FIELDS,['编号','工况','截面（输入搜索）','钢材','关联支撑／工况','自动分布荷载\nkN/m','自动荷载宽度\nm','Span\nm','Lx\nm','Ly\nm','Lx=Ly','Kx','Ky','KLT'])
  extra=QWidget();el=QVBoxLayout(extra)
  self.welds=QTableWidget(0,len(WELD_FIELDS));self.welds.setHorizontalHeaderLabels(['构件编号','工况','连接','焊脚 mm','焊缝强度 MPa','有效焊长 mm/m','分配长度 m']);self.welds.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents);el.addWidget(self.welds)
  buttons=QHBoxLayout();el.addLayout(buttons)
  for label,action in [('增加焊缝',lambda:self.add_weld()),('删除选中焊缝',lambda:self.delete_rows(self.welds))]:
   b=QPushButton(label);b.clicked.connect(action);buttons.addWidget(b)
  el.addWidget(QLabel('录入焊缝后生成时自动加入；不填写则只生成构件验算。支撑焊长自动计算，围檩焊长按每米墙宽输入。'))
  bracket_label=QLabel('Bracket：承担托架间距范围内的围檩自重及自身自重。围檩重心距墙为力臂（原表0.5m），与托架长度分开输入；采用AA文件的Bracket工作表验算构件。');bracket_label.setWordWrap(True);el.addWidget(bracket_label)
  self.brackets=QTableWidget(0,len(BRACKET_FIELDS));self.brackets.setHorizontalHeaderLabels(['托架编号','工况','围檩编号','围檩工况','截面（输入搜索）','钢材','托架长度\nm','托架间距\nm','受力模型','围檩重心距墙\nm','焊脚\nmm','焊缝强度\nMPa','总有效焊长\nmm'])
  self.brackets.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents);el.addWidget(self.brackets)
  self.brackets.setColumnHidden(BRACKET_FIELDS.index('load_model'),True)
  bracket_buttons=QHBoxLayout();el.addLayout(bracket_buttons)
  for label,action in [('增加托架',lambda:self.add_bracket()),('删除选中托架',lambda:self.delete_rows(self.brackets))]:
   button=QPushButton(label);button.clicked.connect(action);bracket_buttons.addWidget(button)
  self.brackets.itemChanged.connect(self.emit_changed)
  self.include_blinding=QCheckBox('同时生成 Blinding（使用钢板桩页面的未分项剪力）');el.addWidget(self.include_blinding)
  form=QFormLayout();self.thickness=QLineEdit();self.thickness.setPlaceholderText('输入厚度 mm，fcu沿用模板20 MPa');form.addRow('Blinding厚度 mm',self.thickness);el.addLayout(form)
  self.tabs.addTab(extra,'Weld／Bracket／Blinding')
  self.generate=QPushButton('生成围檩／支撑计算书及 PDF');self.generate.setMinimumHeight(38);self.generate.clicked.connect(lambda *_:self.generate_requested.emit());layout.addWidget(self.generate)
  layout.removeWidget(self.generate);layout.insertWidget(2,self.generate)
  self.result=QTextEdit();self.result.setReadOnly(True);self.result.setMaximumHeight(100);layout.addWidget(self.result)
  for field in (self.source,self.thickness):field.textChanged.connect(lambda *_:self.changed.emit())
  self.combined.toggled.connect(lambda *_:self.changed.emit());self.include_blinding.toggled.connect(lambda *_:self.changed.emit())
  self.include_blinding.toggled.connect(self.thickness.setEnabled);self.thickness.setEnabled(False)
  self.welds.itemChanged.connect(self.emit_changed)
 def member_page(self,kind,fields,headers):
  page=QWidget();layout=QVBoxLayout(page)
  layout.addWidget(QLabel('Strut：墙法线力按kN输入；角度为支撑轴线与墙法线的夹角，角撑轴力除以cos(角度)。Strut length用于自重验算，Lx默认跟随该长度。' if kind=='Strut' else 'Waling：选择关联支撑，荷载宽度自动取其间距，分布荷载自动由支撑力／间距得到；Span手动输入，整段按此均布荷载计算。'))
  table=QTableWidget(0,len(fields));table.setHorizontalHeaderLabels(headers);table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents);layout.addWidget(table)
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
  if not record.get('member_id'):record['member_id']=self.next_id(kind)
  row=table.rowCount();self.busy=True;table.insertRow(row)
  for col,key in enumerate(fields):
   value=record.get(key,False if key=='lx_equals_ly' else 1 if key in ('kx','ky','klt') else 0 if key=='angle_deg' else '')
   if key=='section':table.setCellWidget(row,col,self.section_combo(str(value)))
   elif key in ('grade','source_strut'):
    combo=QComboBox();combo.addItems(['请选择','S235','S275','S355','S460'] if key=='grade' else ['选择关联支撑']);combo.setCurrentText(str(value));combo.setProperty('restore_value',str(value));combo.currentTextChanged.connect(self.emit_changed);table.setCellWidget(row,col,combo)
   elif key=='lx_equals_ly':
    box=QCheckBox();box.setChecked(bool(value));box.toggled.connect(self.emit_changed);table.setCellWidget(row,col,box)
   else:table.setItem(row,col,QTableWidgetItem('' if value is None else str(value)))
  if kind=='Strut':table.item(row,fields.index('lx_m')).setData(Qt.ItemDataRole.UserRole,record.get('lx_is_default',record.get('lx_m') in ('',None)))
  self.busy=False;self.refresh();self.changed.emit()
 def member_item_changed(self,kind,table,item):
  if self.busy:return
  if kind=='Strut' and item.column()==STRUT_FIELDS.index('lx_m'):
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
     if table is self.struts:
      lx_item=table.item(row,fields.index('lx_m'))
      if lx_item.data(Qt.ItemDataRole.UserRole):lx_item.setText(self.text(table,row,fields.index('strut_length_m')))
     linked=table.cellWidget(row,fields.index('lx_equals_ly')).isChecked();lx=self.text(table,row,fields.index('lx_m'));ly=table.item(row,fields.index('ly_m'))
     if linked:ly.setText(lx)
     ly.setFlags(ly.flags() & ~Qt.ItemFlag.ItemIsEditable if linked else ly.flags() | Qt.ItemFlag.ItemIsEditable)
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
  finally:self.busy=False
 def duplicate(self,kind,count):
  table=self.struts if kind=='Strut' else self.walers;fields=STRUT_FIELDS if kind=='Strut' else WALER_FIELDS;row=table.currentRow()
  if row<0:raise ValueError('请先选中一行作为复制模板')
  record=self.read_rows(table,fields)[row]
  if kind=='Strut':record['lx_is_default']=bool(table.item(row,fields.index('lx_m')).data(Qt.ItemDataRole.UserRole))
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
  record.setdefault('load_model','point')
  for col,key in enumerate(BRACKET_FIELDS):
   value=record.get(key,'')
   if key=='section':self.brackets.setCellWidget(row,col,self.section_combo(value))
   elif key in ('grade','load_model'):
    combo=QComboBox();combo.addItems(['请选择','S235','S275','S355','S460'] if key=='grade' else ['选择受力模型','point']);combo.setCurrentText(str(value));combo.currentTextChanged.connect(self.emit_changed);self.brackets.setCellWidget(row,col,combo)
   else:self.brackets.setItem(row,col,QTableWidgetItem('' if value is None else str(value)))
  self.busy=False;self.changed.emit()
 def collect(self,length,width):
  self.refresh();members=[]
  for row,m in enumerate(self.read_rows(self.struts,STRUT_FIELDS)):
   m['kind']='Strut';m['lx_is_default']=bool(self.struts.item(row,STRUT_FIELDS.index('lx_m')).data(Qt.ItemDataRole.UserRole));members.append(m)
  for m in self.read_rows(self.walers,WALER_FIELDS):
   source=m.pop('source_strut');parts=source.split(' / ');m.update(kind='Waler',angle_deg=0.0,source_strut_id=parts[0] if len(parts)==2 else '',source_strut_case=parts[1] if len(parts)==2 else '');members.append(m)
  try:thickness=float(self.thickness.text()) if self.thickness.text().strip() else None
  except ValueError:raise ValueError('Blinding厚度须为数字')
  return dict(schema_version=2,connection_template=AA_CONNECTIONS,brackets=self.read_rows(self.brackets,BRACKET_FIELDS),enabled=self.combined.isChecked(),include_blinding=self.include_blinding.isChecked(),source=self.source.text().strip(),members=members,welds=self.read_welds(),length_m=length,width_m=width,blinding_thickness_mm=thickness)
 def load(self,record):
  self.struts.setRowCount(0);self.walers.setRowCount(0);self.welds.setRowCount(0);self.brackets.setRowCount(0)
  self.combined.setChecked(bool(record.get('enabled',False)));self.source.setText(record.get('source',''));self.include_blinding.setChecked(record.get('include_blinding',record.get('blinding_thickness_mm') is not None));self.thickness.setText('' if record.get('blinding_thickness_mm') is None else str(record['blinding_thickness_mm']))
  for m in record.get('members',[]):
   if m['kind']=='Strut':self.add_member('Strut',m)
  for m in record.get('members',[]):
   if m['kind']=='Waler':
    row=dict(m);row['source_strut']=row.get('source_strut_id','')+' / '+row.get('source_strut_case','');self.add_member('Waler',row)
  for w in record.get('welds',[]):self.add_weld(w)
  for b in record.get('brackets',[]):self.add_bracket(b)
