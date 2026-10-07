# 手语歌智能工坊：工具说明

供智能体在调用工具、配置依赖和处理文档技术问题时查阅。备课流程与翻译要求见根目录`AGENTS.md`。

`scripts/requirements.txt`记录工具所需的Python库：python-docx用于Word，pdfplumber用于PDF。由智能体先查找客户端自带或系统已有的兼容Python，环境已有这些库时无需重复安装。缺少依赖时由智能体在项目根目录自动执行`python -m pip install -r scripts/requirements.txt`；缺少Python时优先自动配置运行时。脚本失败后先由模型诊断并调整方案，确实无法完成才请老师执行最少的手动步骤。查MDX和检索Markdown的脚本不依赖这两个库。依赖清单不属于教学资料，无需手语老师手动维护。

转写检索示例：`python scripts/search_cases.py 名字 --source 教材 --limit 3`。来源筛选匹配文件名、章节或“教材／优秀讲义”；检索读取Excel完整记录及辨析Markdown，包含原句、转写与前后文，返回原件位置。

新增资料可用`python scripts/extract_examples.py '<原件>' --out '<工作台/临时提取.md>'`生成临时定位稿。核对PDF格式、同段多个转写和跨行括号后，将完整记录写入`资料/提取与检索/原句与手语转写.xlsx`，临时稿不作为第二份长期资料库。修改辨析全文维护Markdown，其目录表只记录检索摘要。

既有Word局部编辑时，不能给图片所在run直接赋`text`，否则会丢图。保留绘图、裁剪和尺寸，不从旧词典整图重建老师截图。



一般只需告诉智能体任务，无需老师维护脚本或清单。工具均使用当前文件和相对路径；在本文件夹执行下列命令，`python`换成本机可用运行时。

```powershell
python scripts/mdict_reader.py query dictionaries/国家通用手语词典/国家通用手语词典.mdx 跳舞 --summary
python scripts/search_cases.py 否定
python scripts/read_document.py '资料/优秀讲义/《暮色回响》手语讲义定稿（已校）.docx'
python scripts/audit_docx.py '<老师版.docx>' '<打印版.docx>'
```

查词可加`--contains`寻找候选。提取候选图：

```powershell
python scripts/mdict_reader.py images dictionaries/国家通用手语词典/国家通用手语词典.mdx dictionaries/国家通用手语词典/国家通用手语词典.mdd 跳舞 '<工作台/歌题/候选图>'
```

该命令提取全部引用图供选择，不能自动全贴。读取器支持本包使用的未加密MDict2、zlib或未压缩块；不支持的加密、LZO或缺失资源须报告，可改用合法可用的词典软件查询。

`read_document.py`按Word正文段落、表格及批注，或PDF页码输出可检索文本；原文有删除线或斜体时保留标记。默认只输出，不另建资料副本；PDF的格式与图片仍需看原页。

新图解可将已确认内容写成一份Markdown，`# 歌题`、普通歌词转写行、下一行的`![词](图片路径)`依次排版。`*文字*`生成斜体，`~~文字~~`生成删除线，单个`~`原样保留，`++`为正体；图片标题`"参考"`表示原生1pt点线框，例：`![桥](候选图/桥.jpg "参考")`。这些是工具支持的格式，具体含义和使用方式按本次讲义格式要求确定。每个词选哪幅图由老师或助手先确定。

```powershell
python scripts/make_handout.py '<工作台/歌题/内容.md>' --out '<工作台/歌题/讲义.docx>' --height 0.79
python scripts/audit_docx.py '<工作台/歌题/讲义.docx>' --height 0.79
```

图解工具生成新文档，不覆盖已有文件，不修改原讲义。需要精细沿用既有讲义时用局部Word编辑，而不是重建。`audit_docx.py`检查两版转写、斜体++、字体字号分布、嵌入/浮动图、固定行距、图片是否超宽；指定图高时才核对高度，文化照片等例外人工判断。

新电脑复制整个文件夹，由智能体按上述顺序自动配置Python及`scripts/requirements.txt`中的依赖。Word渲染可用获许可的Word或环境提供的渲染器；Windows可隐藏、只读打开副本导出PDF，不保存、不关闭用户的Word进程。宋体缺失时提示老师确认替代，不能声称版式一致。



