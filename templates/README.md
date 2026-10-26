# 模板接入资料

请将已复核的Excel计算书及对应DWG图纸复制到本目录，保留原始文件。
本目录内的工程模板默认不提交GitHub。

需要确认：Excel版本，宏与外部链接依赖；输入单元格、选型方式、结果单元格；DWG版本、图层、块属性、节点适用条件；至少一个对应的分析内力与已复核成果。

当前Excel检查器只读取xlsx/xlsm结构，不执行计算、不修改模板。
DWG的对象映射与实际出图操作须根据真实模板建立并在AutoCAD中验证。

`supports-structure.local.json`仅保存在本机，记录工程师选用的AA文件、Connection check、Bracket、Blinding及相关源文件指纹。审查过的新版本可运行`python -m tools.register_supports_template --source "完整模板路径.xlsx"`重新登记；登记本身不替代计算方法审查。独立支撑／围檩生成直接从该XLSX开始，不需要先打开旧XLS。模板内容、登记文件、项目数据库和输出仍不提交GitHub。
