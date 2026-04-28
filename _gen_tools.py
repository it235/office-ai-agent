import json
import os

base_dir = 'F:/ai/code/AiHelper/ShareRibbon/Tools'

excel_tools = [
    {'id': 'ApplyFormula', 'name': '应用公式', 'description': '在指定单元格范围应用Excel公式，支持自动向下填充', 'category': '基础操作', 'riskLevel': 'safe', 'params': [
        {'name': 'targetRange', 'type': 'string', 'required': True, 'description': '目标单元格范围，如 C1:C100'},
        {'name': 'formula', 'type': 'string', 'required': True, 'description': 'Excel公式，如 =A1+B1'},
        {'name': 'fillDown', 'type': 'boolean', 'required': False, 'description': '是否自动向下填充', 'defaultValue': True}
    ]},
    {'id': 'WriteData', 'name': '写入数据', 'description': '向指定单元格范围写入数据，支持单值或二维数组', 'category': '基础操作', 'riskLevel': 'safe', 'params': [
        {'name': 'targetRange', 'type': 'string', 'required': True, 'description': '目标单元格范围'},
        {'name': 'data', 'type': 'array', 'required': True, 'description': '要写入的数据（单值或二维数组）'}
    ]},
    {'id': 'FormatRange', 'name': '格式化范围', 'description': '设置单元格的样式格式，包括字体、颜色、边框等', 'category': '基础操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '要格式化的单元格范围'},
        {'name': 'style', 'type': 'string', 'required': False, 'description': '预设样式：header/total/data'},
        {'name': 'bold', 'type': 'boolean', 'required': False, 'description': '是否加粗'},
        {'name': 'italic', 'type': 'boolean', 'required': False, 'description': '是否斜体'},
        {'name': 'fontSize', 'type': 'integer', 'required': False, 'description': '字体大小'},
        {'name': 'backgroundColor', 'type': 'string', 'required': False, 'description': '背景色'},
        {'name': 'fontColor', 'type': 'string', 'required': False, 'description': '字体颜色'},
        {'name': 'borders', 'type': 'string', 'required': False, 'description': '边框样式：true/all/outline/none'}
    ]},
    {'id': 'CreateChart', 'name': '创建图表', 'description': '根据数据范围创建图表', 'category': '基础操作', 'riskLevel': 'safe', 'params': [
        {'name': 'dataRange', 'type': 'string', 'required': True, 'description': '图表数据源范围'},
        {'name': 'type', 'type': 'string', 'required': True, 'description': '图表类型：column/line/pie/bar/scatter/area'},
        {'name': 'title', 'type': 'string', 'required': False, 'description': '图表标题'},
        {'name': 'position', 'type': 'string', 'required': False, 'description': '图表放置位置'},
        {'name': 'seriesNames', 'type': 'array', 'required': False, 'description': '系列名称数组'},
        {'name': 'categoryAxis', 'type': 'string', 'required': False, 'description': '分类轴范围'},
        {'name': 'legendPosition', 'type': 'string', 'required': False, 'description': '图例位置：right/left/top/bottom'}
    ]},
    {'id': 'CleanData', 'name': '数据清洗', 'description': '对指定范围进行数据清洗操作', 'category': '基础操作', 'riskLevel': 'medium', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '要清洗的数据范围'},
        {'name': 'operation', 'type': 'string', 'required': True, 'description': '清洗操作：removeduplicates/fillempty/trim/replace'}
    ]},
    {'id': 'SortData', 'name': '排序数据', 'description': '对指定范围的数据进行排序', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '要排序的数据范围'},
        {'name': 'sortColumn', 'type': 'integer', 'required': True, 'description': '排序列号（从1开始）'},
        {'name': 'order', 'type': 'string', 'required': False, 'description': '排序方向：asc/desc', 'defaultValue': 'asc'},
        {'name': 'hasHeader', 'type': 'boolean', 'required': False, 'description': '是否包含表头', 'defaultValue': True}
    ]},
    {'id': 'FilterData', 'name': '筛选数据', 'description': '对数据范围应用或清除筛选', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '要筛选的数据范围'},
        {'name': 'column', 'type': 'integer', 'required': True, 'description': '筛选列号'},
        {'name': 'criteria', 'type': 'string', 'required': False, 'description': '筛选条件，如 >100'},
        {'name': 'clearFilter', 'type': 'boolean', 'required': False, 'description': '是否清除筛选', 'defaultValue': False}
    ]},
    {'id': 'RemoveDuplicates', 'name': '删除重复项', 'description': '删除指定范围内的重复行', 'category': '数据操作', 'riskLevel': 'medium', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '数据范围'},
        {'name': 'columns', 'type': 'array', 'required': False, 'description': '检查重复的列号数组'},
        {'name': 'hasHeader', 'type': 'boolean', 'required': False, 'description': '是否包含表头', 'defaultValue': True}
    ]},
    {'id': 'ConditionalFormat', 'name': '条件格式', 'description': '为单元格范围添加条件格式', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '目标范围'},
        {'name': 'rule', 'type': 'string', 'required': True, 'description': '规则类型：highlight/databar/colorscale/iconset'},
        {'name': 'condition', 'type': 'string', 'required': False, 'description': '条件表达式'},
        {'name': 'color', 'type': 'string', 'required': False, 'description': '颜色'}
    ]},
    {'id': 'MergeCells', 'name': '合并单元格', 'description': '合并或取消合并指定范围的单元格', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '单元格范围'},
        {'name': 'unmerge', 'type': 'boolean', 'required': False, 'description': '是否取消合并', 'defaultValue': False}
    ]},
    {'id': 'AutoFit', 'name': '自动调整', 'description': '自动调整列宽或行高', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '目标范围'},
        {'name': 'type', 'type': 'string', 'required': True, 'description': '调整类型：columns/rows/both'}
    ]},
    {'id': 'FindReplace', 'name': '查找替换', 'description': '在范围内查找并替换文本', 'category': '数据操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '查找范围，all 表示全部'},
        {'name': 'find', 'type': 'string', 'required': True, 'description': '查找内容'},
        {'name': 'replace', 'type': 'string', 'required': True, 'description': '替换内容'},
        {'name': 'matchCase', 'type': 'boolean', 'required': False, 'description': '是否区分大小写'},
        {'name': 'matchEntireCell', 'type': 'boolean', 'required': False, 'description': '是否全单元格匹配'}
    ]},
    {'id': 'CreatePivotTable', 'name': '创建透视表', 'description': '从数据源创建数据透视表', 'category': '数据操作', 'riskLevel': 'medium', 'params': [
        {'name': 'sourceRange', 'type': 'string', 'required': True, 'description': '数据源范围'},
        {'name': 'targetCell', 'type': 'string', 'required': True, 'description': '透视表放置位置'},
        {'name': 'rowFields', 'type': 'array', 'required': True, 'description': '行字段数组'},
        {'name': 'valueFields', 'type': 'array', 'required': True, 'description': '值字段数组'},
        {'name': 'columnFields', 'type': 'array', 'required': False, 'description': '列字段数组'}
    ]},
    {'id': 'CreateSheet', 'name': '创建工作表', 'description': '创建新的工作表', 'category': '工作表操作', 'riskLevel': 'safe', 'params': [
        {'name': 'name', 'type': 'string', 'required': True, 'description': '工作表名称'},
        {'name': 'position', 'type': 'string', 'required': False, 'description': '插入位置：before/after'},
        {'name': 'referenceSheet', 'type': 'string', 'required': False, 'description': '参考工作表名称'}
    ]},
    {'id': 'DeleteSheet', 'name': '删除工作表', 'description': '删除指定工作表', 'category': '工作表操作', 'riskLevel': 'risky', 'params': [
        {'name': 'name', 'type': 'string', 'required': True, 'description': '要删除的工作表名称'}
    ]},
    {'id': 'RenameSheet', 'name': '重命名工作表', 'description': '重命名工作表', 'category': '工作表操作', 'riskLevel': 'safe', 'params': [
        {'name': 'oldName', 'type': 'string', 'required': True, 'description': '原名称'},
        {'name': 'newName', 'type': 'string', 'required': True, 'description': '新名称'}
    ]},
    {'id': 'CopySheet', 'name': '复制工作表', 'description': '复制工作表', 'category': '工作表操作', 'riskLevel': 'safe', 'params': [
        {'name': 'sourceName', 'type': 'string', 'required': True, 'description': '源工作表名称'},
        {'name': 'newName', 'type': 'string', 'required': True, 'description': '新工作表名称'}
    ]},
    {'id': 'InsertRowCol', 'name': '插入行列', 'description': '插入行或列', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'type', 'type': 'string', 'required': True, 'description': '插入类型：row/column'},
        {'name': 'position', 'type': 'string', 'required': True, 'description': '插入位置（行号或列字母）'},
        {'name': 'count', 'type': 'integer', 'required': False, 'description': '插入数量', 'defaultValue': 1}
    ]},
    {'id': 'DeleteRowCol', 'name': '删除行列', 'description': '删除指定行或列', 'category': '高级功能', 'riskLevel': 'medium', 'params': [
        {'name': 'type', 'type': 'string', 'required': True, 'description': '删除类型：row/column'},
        {'name': 'position', 'type': 'string', 'required': True, 'description': '位置（行号或列字母）'},
        {'name': 'count', 'type': 'integer', 'required': False, 'description': '删除数量', 'defaultValue': 1}
    ]},
    {'id': 'HideRowCol', 'name': '隐藏行列', 'description': '隐藏或取消隐藏行或列', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'type', 'type': 'string', 'required': True, 'description': '类型：row/column'},
        {'name': 'position', 'type': 'string', 'required': True, 'description': '位置'},
        {'name': 'unhide', 'type': 'boolean', 'required': False, 'description': '是否取消隐藏', 'defaultValue': False}
    ]},
    {'id': 'ProtectSheet', 'name': '保护工作表', 'description': '保护或取消保护工作表', 'category': '高级功能', 'riskLevel': 'medium', 'params': [
        {'name': 'sheetName', 'type': 'string', 'required': False, 'description': '工作表名称'},
        {'name': 'password', 'type': 'string', 'required': False, 'description': '保护密码'},
        {'name': 'unprotect', 'type': 'boolean', 'required': False, 'description': '是否取消保护', 'defaultValue': False}
    ]},
    {'id': 'ExecuteVBA', 'name': '执行VBA', 'description': '执行VBA代码作为回退方案，当注册命令无法满足需求时使用', 'category': '高级功能', 'riskLevel': 'risky', 'isVbaFallback': True, 'params': [
        {'name': 'code', 'type': 'string', 'required': True, 'description': '完整的VBA Sub或Function代码'}
    ]},
]

word_tools = [
    {'id': 'InsertText', 'name': '插入文本', 'description': '在文档中插入文本内容', 'category': '基础文本操作', 'riskLevel': 'safe', 'params': [
        {'name': 'content', 'type': 'string', 'required': True, 'description': '要插入的文本内容'},
        {'name': 'position', 'type': 'string', 'required': False, 'description': '插入位置：cursor/start/end', 'defaultValue': 'cursor'}
    ]},
    {'id': 'FormatText', 'name': '格式化文本', 'description': '设置文本格式样式', 'category': '基础文本操作', 'riskLevel': 'safe', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '目标范围：selection/all'},
        {'name': 'bold', 'type': 'boolean', 'required': False, 'description': '是否加粗'},
        {'name': 'italic', 'type': 'boolean', 'required': False, 'description': '是否斜体'},
        {'name': 'fontSize', 'type': 'integer', 'required': False, 'description': '字体大小'},
        {'name': 'fontName', 'type': 'string', 'required': False, 'description': '字体名称'},
        {'name': 'underline', 'type': 'boolean', 'required': False, 'description': '是否下划线'},
        {'name': 'color', 'type': 'string', 'required': False, 'description': '字体颜色'}
    ]},
    {'id': 'ReplaceText', 'name': '查找替换', 'description': '查找并替换文本内容', 'category': '基础文本操作', 'riskLevel': 'safe', 'params': [
        {'name': 'find', 'type': 'string', 'required': True, 'description': '查找内容'},
        {'name': 'replace', 'type': 'string', 'required': True, 'description': '替换内容'},
        {'name': 'matchCase', 'type': 'boolean', 'required': False, 'description': '是否区分大小写'}
    ]},
    {'id': 'DeleteText', 'name': '删除文本', 'description': '删除指定范围的文本', 'category': '基础文本操作', 'riskLevel': 'medium', 'params': [
        {'name': 'range', 'type': 'string', 'required': True, 'description': '范围：selection/all'}
    ]},
    {'id': 'CopyPasteText', 'name': '复制粘贴文本', 'description': '复制粘贴文本内容', 'category': '基础文本操作', 'riskLevel': 'safe', 'params': [
        {'name': 'sourceRange', 'type': 'string', 'required': True, 'description': '源范围'},
        {'name': 'targetPosition', 'type': 'string', 'required': False, 'description': '目标位置'}
    ]},
    {'id': 'ApplyStyle', 'name': '应用样式', 'description': '应用Word预设样式', 'category': '段落和样式', 'riskLevel': 'safe', 'params': [
        {'name': 'styleName', 'type': 'string', 'required': True, 'description': '样式名称，如 "标题 1"'},
        {'name': 'range', 'type': 'string', 'required': False, 'description': '范围：selection/paragraph', 'defaultValue': 'selection'}
    ]},
    {'id': 'SetParagraphFormat', 'name': '设置段落格式', 'description': '设置段落对齐、缩进、间距等', 'category': '段落和样式', 'riskLevel': 'safe', 'params': [
        {'name': 'alignment', 'type': 'string', 'required': False, 'description': '对齐方式：left/center/right/justify'},
        {'name': 'firstLineIndent', 'type': 'number', 'required': False, 'description': '首行缩进'},
        {'name': 'beforeSpacing', 'type': 'number', 'required': False, 'description': '段前间距'},
        {'name': 'afterSpacing', 'type': 'number', 'required': False, 'description': '段后间距'}
    ]},
    {'id': 'InsertParagraph', 'name': '插入段落', 'description': '插入新段落', 'category': '段落和样式', 'riskLevel': 'safe', 'params': [
        {'name': 'count', 'type': 'integer', 'required': False, 'description': '插入段落数', 'defaultValue': 1},
        {'name': 'pageBreak', 'type': 'boolean', 'required': False, 'description': '是否分页', 'defaultValue': False}
    ]},
    {'id': 'SetLineSpacing', 'name': '设置行距', 'description': '设置行间距', 'category': '段落和样式', 'riskLevel': 'safe', 'params': [
        {'name': 'spacing', 'type': 'number', 'required': True, 'description': '行距：1/1.5/2'},
        {'name': 'range', 'type': 'string', 'required': False, 'description': '范围：selection/all', 'defaultValue': 'selection'}
    ]},
    {'id': 'SetIndent', 'name': '设置缩进', 'description': '设置段落缩进', 'category': '段落和样式', 'riskLevel': 'safe', 'params': [
        {'name': 'left', 'type': 'number', 'required': False, 'description': '左缩进(cm)'},
        {'name': 'right', 'type': 'number', 'required': False, 'description': '右缩进(cm)'},
        {'name': 'firstLine', 'type': 'number', 'required': False, 'description': '首行缩进(cm)'}
    ]},
    {'id': 'InsertTable', 'name': '插入表格', 'description': '插入新表格', 'category': '表格操作', 'riskLevel': 'safe', 'params': [
        {'name': 'rows', 'type': 'integer', 'required': True, 'description': '行数'},
        {'name': 'cols', 'type': 'integer', 'required': True, 'description': '列数'},
        {'name': 'data', 'type': 'array', 'required': False, 'description': '表格数据（可选）'}
    ]},
    {'id': 'FormatTable', 'name': '格式化表格', 'description': '设置表格样式', 'category': '表格操作', 'riskLevel': 'safe', 'params': [
        {'name': 'tableIndex', 'type': 'integer', 'required': True, 'description': '表格索引（从1开始）'},
        {'name': 'style', 'type': 'string', 'required': False, 'description': '表格样式'},
        {'name': 'borders', 'type': 'boolean', 'required': False, 'description': '是否显示边框'},
        {'name': 'headerRow', 'type': 'boolean', 'required': False, 'description': '是否有标题行'}
    ]},
    {'id': 'InsertTableRow', 'name': '插入表格行', 'description': '在表格中插入行', 'category': '表格操作', 'riskLevel': 'safe', 'params': [
        {'name': 'tableIndex', 'type': 'integer', 'required': True, 'description': '表格索引'},
        {'name': 'position', 'type': 'string', 'required': True, 'description': '插入位置：after/before'}
    ]},
    {'id': 'DeleteTableRow', 'name': '删除表格行', 'description': '删除表格中的行', 'category': '表格操作', 'riskLevel': 'medium', 'params': [
        {'name': 'tableIndex', 'type': 'integer', 'required': True, 'description': '表格索引'},
        {'name': 'rowIndex', 'type': 'integer', 'required': True, 'description': '要删除的行号'}
    ]},
    {'id': 'GenerateTOC', 'name': '生成目录', 'description': '自动生成文档目录', 'category': '文档结构', 'riskLevel': 'safe', 'params': [
        {'name': 'position', 'type': 'string', 'required': False, 'description': '目录位置：start/cursor', 'defaultValue': 'cursor'},
        {'name': 'levels', 'type': 'integer', 'required': False, 'description': '目录层级（1-9）', 'defaultValue': 3}
    ]},
    {'id': 'InsertHeader', 'name': '插入页眉', 'description': '插入或修改页眉', 'category': '文档结构', 'riskLevel': 'safe', 'params': [
        {'name': 'content', 'type': 'string', 'required': True, 'description': '页眉内容'},
        {'name': 'alignment', 'type': 'string', 'required': False, 'description': '对齐方式：left/center/right', 'defaultValue': 'center'}
    ]},
    {'id': 'InsertFooter', 'name': '插入页脚', 'description': '插入或修改页脚', 'category': '文档结构', 'riskLevel': 'safe', 'params': [
        {'name': 'content', 'type': 'string', 'required': True, 'description': '页脚内容'},
        {'name': 'alignment', 'type': 'string', 'required': False, 'description': '对齐方式：left/center/right', 'defaultValue': 'center'}
    ]},
    {'id': 'InsertPageNumber', 'name': '插入页码', 'description': '插入页码到页眉或页脚', 'category': '文档结构', 'riskLevel': 'safe', 'params': [
        {'name': 'position', 'type': 'string', 'required': True, 'description': '位置：header/footer'},
        {'name': 'alignment', 'type': 'string', 'required': False, 'description': '对齐方式：left/center/right', 'defaultValue': 'center'}
    ]},
    {'id': 'BeautifyDocument', 'name': '美化文档', 'description': '一键美化文档整体样式', 'category': '文档美化', 'riskLevel': 'safe', 'params': [
        {'name': 'theme', 'type': 'object', 'required': False, 'description': '主题设置：{h1, h2, body}'},
        {'name': 'margins', 'type': 'object', 'required': False, 'description': '页边距：{top, bottom, left, right}'}
    ]},
    {'id': 'SetPageMargins', 'name': '设置页边距', 'description': '设置页面边距', 'category': '文档美化', 'riskLevel': 'safe', 'params': [
        {'name': 'top', 'type': 'number', 'required': False, 'description': '上边距(cm)'},
        {'name': 'bottom', 'type': 'number', 'required': False, 'description': '下边距(cm)'},
        {'name': 'left', 'type': 'number', 'required': False, 'description': '左边距(cm)'},
        {'name': 'right', 'type': 'number', 'required': False, 'description': '右边距(cm)'}
    ]},
    {'id': 'InsertImage', 'name': '插入图片', 'description': '在文档中插入图片', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'imagePath', 'type': 'string', 'required': True, 'description': '图片文件路径'},
        {'name': 'width', 'type': 'number', 'required': False, 'description': '图片宽度'},
        {'name': 'height', 'type': 'number', 'required': False, 'description': '图片高度'}
    ]},
    {'id': 'ExecuteVBA', 'name': '执行VBA', 'description': '执行VBA代码作为回退方案', 'category': '高级功能', 'riskLevel': 'risky', 'isVbaFallback': True, 'params': [
        {'name': 'code', 'type': 'string', 'required': True, 'description': '完整的VBA Sub或Function代码'}
    ]},
]

ppt_tools = [
    {'id': 'InsertSlide', 'name': '插入幻灯片', 'description': '在指定位置插入新幻灯片', 'category': '幻灯片操作', 'riskLevel': 'safe', 'params': [
        {'name': 'position', 'type': 'string', 'required': False, 'description': '插入位置：current/end', 'defaultValue': 'end'},
        {'name': 'layout', 'type': 'string', 'required': False, 'description': '幻灯片布局'},
        {'name': 'title', 'type': 'string', 'required': False, 'description': '幻灯片标题'},
        {'name': 'content', 'type': 'string', 'required': False, 'description': '幻灯片内容'}
    ]},
    {'id': 'DeleteSlide', 'name': '删除幻灯片', 'description': '删除指定幻灯片', 'category': '幻灯片操作', 'riskLevel': 'medium', 'params': [
        {'name': 'slideIndex', 'type': 'integer', 'required': True, 'description': '幻灯片索引，-1表示当前'}
    ]},
    {'id': 'DuplicateSlide', 'name': '复制幻灯片', 'description': '复制指定幻灯片', 'category': '幻灯片操作', 'riskLevel': 'safe', 'params': [
        {'name': 'slideIndex', 'type': 'integer', 'required': True, 'description': '要复制的幻灯片索引'}
    ]},
    {'id': 'MoveSlide', 'name': '移动幻灯片', 'description': '移动幻灯片到指定位置', 'category': '幻灯片操作', 'riskLevel': 'safe', 'params': [
        {'name': 'fromIndex', 'type': 'integer', 'required': True, 'description': '源位置'},
        {'name': 'toIndex', 'type': 'integer', 'required': True, 'description': '目标位置'}
    ]},
    {'id': 'CreateSlides', 'name': '批量创建幻灯片', 'description': '批量创建多个幻灯片', 'category': '幻灯片操作', 'riskLevel': 'safe', 'params': [
        {'name': 'slides', 'type': 'array', 'required': True, 'description': '幻灯片数组，每项含title/content/layout'}
    ]},
    {'id': 'InsertText', 'name': '插入文本', 'description': '在幻灯片中插入文本', 'category': '内容操作', 'riskLevel': 'safe', 'params': [
        {'name': 'content', 'type': 'string', 'required': True, 'description': '文本内容'},
        {'name': 'slideIndex', 'type': 'integer', 'required': False, 'description': '幻灯片索引，-1表示当前', 'defaultValue': -1},
        {'name': 'x', 'type': 'number', 'required': False, 'description': 'X坐标'},
        {'name': 'y', 'type': 'number', 'required': False, 'description': 'Y坐标'}
    ]},
    {'id': 'FormatText', 'name': '格式化文本', 'description': '设置幻灯片文本格式', 'category': '内容操作', 'riskLevel': 'safe', 'params': [
        {'name': 'bold', 'type': 'boolean', 'required': False, 'description': '是否加粗'},
        {'name': 'italic', 'type': 'boolean', 'required': False, 'description': '是否斜体'},
        {'name': 'fontSize', 'type': 'integer', 'required': False, 'description': '字体大小'},
        {'name': 'fontName', 'type': 'string', 'required': False, 'description': '字体名称'},
        {'name': 'color', 'type': 'string', 'required': False, 'description': '字体颜色'}
    ]},
    {'id': 'InsertShape', 'name': '插入形状', 'description': '在幻灯片中插入形状', 'category': '内容操作', 'riskLevel': 'safe', 'params': [
        {'name': 'shapeType', 'type': 'string', 'required': True, 'description': '形状类型'},
        {'name': 'x', 'type': 'number', 'required': True, 'description': 'X坐标'},
        {'name': 'y', 'type': 'number', 'required': True, 'description': 'Y坐标'}
    ]},
    {'id': 'InsertImage', 'name': '插入图片', 'description': '在幻灯片中插入图片', 'category': '内容操作', 'riskLevel': 'safe', 'params': [
        {'name': 'imagePath', 'type': 'string', 'required': True, 'description': '图片路径'},
        {'name': 'x', 'type': 'number', 'required': False, 'description': 'X坐标'},
        {'name': 'y', 'type': 'number', 'required': False, 'description': 'Y坐标'},
        {'name': 'width', 'type': 'number', 'required': False, 'description': '宽度'},
        {'name': 'height', 'type': 'number', 'required': False, 'description': '高度'}
    ]},
    {'id': 'InsertTable', 'name': '插入表格', 'description': '在幻灯片中插入表格', 'category': '内容操作', 'riskLevel': 'safe', 'params': [
        {'name': 'rows', 'type': 'integer', 'required': True, 'description': '行数'},
        {'name': 'cols', 'type': 'integer', 'required': True, 'description': '列数'},
        {'name': 'data', 'type': 'array', 'required': False, 'description': '表格数据'}
    ]},
    {'id': 'FormatSlide', 'name': '格式化幻灯片', 'description': '设置幻灯片背景、布局等', 'category': '样式和动画', 'riskLevel': 'safe', 'params': [
        {'name': 'background', 'type': 'string', 'required': False, 'description': '背景设置'},
        {'name': 'layout', 'type': 'string', 'required': False, 'description': '布局类型'}
    ]},
    {'id': 'AddAnimation', 'name': '添加动画', 'description': '为幻灯片元素添加动画', 'category': '样式和动画', 'riskLevel': 'safe', 'params': [
        {'name': 'effect', 'type': 'string', 'required': True, 'description': '动画效果：fadeIn/flyIn/zoom/wipe'},
        {'name': 'targetShapes', 'type': 'string', 'required': False, 'description': '目标：all/title', 'defaultValue': 'all'}
    ]},
    {'id': 'ApplyTransition', 'name': '应用切换效果', 'description': '设置幻灯片切换动画', 'category': '样式和动画', 'riskLevel': 'safe', 'params': [
        {'name': 'transitionType', 'type': 'string', 'required': True, 'description': '切换类型：fade/push/wipe'},
        {'name': 'scope', 'type': 'string', 'required': False, 'description': '范围：all/current', 'defaultValue': 'all'}
    ]},
    {'id': 'BeautifySlides', 'name': '美化幻灯片', 'description': '一键美化幻灯片样式', 'category': '样式和动画', 'riskLevel': 'safe', 'params': [
        {'name': 'scope', 'type': 'string', 'required': False, 'description': '范围：all/current', 'defaultValue': 'all'},
        {'name': 'theme', 'type': 'object', 'required': False, 'description': '主题：{background, titleFont, bodyFont}'}
    ]},
    {'id': 'SetSlideLayout', 'name': '设置幻灯片布局', 'description': '更改幻灯片布局', 'category': '样式和动画', 'riskLevel': 'safe', 'params': [
        {'name': 'layout', 'type': 'string', 'required': True, 'description': '布局：title/titleAndContent/blank'}
    ]},
    {'id': 'InsertChart', 'name': '插入图表', 'description': '在幻灯片中插入图表', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'chartType', 'type': 'string', 'required': True, 'description': '图表类型：column/line/pie'},
        {'name': 'data', 'type': 'array', 'required': True, 'description': '数据（二维数组）'}
    ]},
    {'id': 'InsertVideo', 'name': '插入视频', 'description': '在幻灯片中插入视频', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'videoPath', 'type': 'string', 'required': True, 'description': '视频文件路径'},
        {'name': 'autoPlay', 'type': 'boolean', 'required': False, 'description': '是否自动播放'}
    ]},
    {'id': 'AddSpeakerNotes', 'name': '添加演讲备注', 'description': '为幻灯片添加演讲者备注', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'notes', 'type': 'string', 'required': True, 'description': '备注内容'},
        {'name': 'slideIndex', 'type': 'integer', 'required': False, 'description': '幻灯片索引'}
    ]},
    {'id': 'SetSlideShow', 'name': '设置放映', 'description': '配置幻灯片放映设置', 'category': '高级功能', 'riskLevel': 'safe', 'params': [
        {'name': 'loopUntilEsc', 'type': 'boolean', 'required': False, 'description': '是否循环放映'},
        {'name': 'advanceMode', 'type': 'string', 'required': False, 'description': '翻页模式'}
    ]},
    {'id': 'ApplyTheme', 'name': '应用主题', 'description': '应用预设主题或自定义主题', 'category': '母版和主题', 'riskLevel': 'safe', 'params': [
        {'name': 'themeName', 'type': 'string', 'required': False, 'description': '主题名称'},
        {'name': 'themeFile', 'type': 'string', 'required': False, 'description': '主题文件路径'}
    ]},
    {'id': 'EditSlideMaster', 'name': '编辑母版', 'description': '编辑幻灯片母版', 'category': '母版和主题', 'riskLevel': 'medium', 'params': [
        {'name': 'background', 'type': 'string', 'required': False, 'description': '背景设置'},
        {'name': 'titleFont', 'type': 'string', 'required': False, 'description': '标题字体'},
        {'name': 'bodyFont', 'type': 'string', 'required': False, 'description': '正文字体'}
    ]},
    {'id': 'ExecuteVBA', 'name': '执行VBA', 'description': '执行VBA代码作为回退方案', 'category': '高级功能', 'riskLevel': 'risky', 'isVbaFallback': True, 'params': [
        {'name': 'code', 'type': 'string', 'required': True, 'description': '完整的VBA Sub代码，换行用\\n转义'}
    ]},
]

def write_tools(tools, app_type):
    for tool in tools:
        path = os.path.join(base_dir, app_type, tool['id'] + '.json')
        data = {
            'id': tool['id'], 'name': tool['name'], 'description': tool['description'],
            'appType': app_type if app_type != 'common' else 'common',
            'category': tool['category'], 'riskLevel': tool['riskLevel'],
            'isVbaFallback': tool.get('isVbaFallback', False), 'parameters': tool['params']
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'{app_type}: {len(tools)} tools written')

write_tools(excel_tools, 'excel')
write_tools(word_tools, 'word')
write_tools(ppt_tools, 'ppt')
