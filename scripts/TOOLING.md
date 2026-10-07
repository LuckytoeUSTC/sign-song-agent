# 手语歌智能工坊：工具说明

编号速查：`R01`公共转写规则，`E0001`转写例句，`D001`教材会话；`太阳-R01`和`太阳-M001`是太阳歌曲的本曲转写规则与修改案例。

供智能体在调用工具、配置依赖和处理文档技术问题时查阅。备课流程与翻译要求见根目录`AGENTS.md`。

新歌先按AGENTS.md完整阅读参照讲义与小班原件，再检索数据库。`read_document.py`读取完整文本、Word格式和批注；它不能显示图示与完整版面，必须同时查看原件页面。查词使用`mdict_reader.py query`，整词失败后按AGENTS.md拆分对象、动作和关系，继续别名与近义词检索。查询现在自动展开别名指向，循环和断链会明确报告；原始记录与图片地址可用于本曲查询依据，不替代语义判断。

内置指拼沿用同一工具：`python scripts/mdict_reader.py query dictionaries/指拼 "X H" --summary`；`python scripts/mdict_reader.py images dictionaries/指拼 dictionaries/指拼 "M C T Z" "工作台/歌题/指拼原图"`。本地资源无MDD，第二个路径仍填指拼目录，便于沿用现有命令结构。`ZH`、`CH`、`SH`、`NG`是单个原图条目，首字母序列必须用空格分开；字母顺序及重复保留。将输出图片路径插入已确认的Markdown，交给`make_handout.py`按正常图片流程制作。程序不根据汉字擅自决定名字读音。

`scripts/requirements.txt`记录工具所需的Python库：python-docx用于Word，pdfplumber用于PDF。由智能体先查找客户端自带或系统已有的兼容Python，环境已有这些库时无需重复安装。缺少依赖时由智能体在项目根目录自动执行`python -m pip install -r scripts/requirements.txt`；缺少Python时优先自动配置运行时。脚本失败后先由模型诊断并调整方案，确实无法完成才请老师执行最少的手动步骤。查MDX和检索Markdown的脚本不依赖这两个库。依赖清单不属于教学资料，无需手语老师手动维护。

转写检索示例：`python scripts/search_cases.py 名字 --source 教材 --limit 3`。来源筛选匹配文件名或“教材／手语歌讲义／小班／修改”。`python scripts/search_cases.py --rule R04`查规则说明及已关联案例；加`--context`展开教材完整会话。读取按工作表和列名定位，不依赖列顺序；返回编号、原件位置和链接。修改案例默认只显示引导，命中后加`--detail`读取正文。

新增资料可用`python scripts/extract_examples.py '<原件>' --out '<工作台/临时提取.md>'`生成临时定位稿。核对PDF格式、同段多个转写和跨行括号后，将完整记录写入`资料/规则与例句数据库/转写例句.xlsx`，临时稿不作为第二份长期资料库。修改辨析每例一个Markdown，保存在对应歌曲的“转写修改与辨析案例”内；修改正文或开头引导后，运行`python scripts/build_case_catalog.py`重新生成目录表。

既有Word局部编辑时，不能给图片所在run直接赋`text`，否则会丢图。保留绘图、裁剪和尺寸，不从旧词典整图重建老师截图。



一般只需告诉智能体任务，无需老师维护脚本或清单。工具均使用当前文件和相对路径；在本文件夹执行下列命令，`python`换成本机可用运行时。

```powershell
python scripts/mdict_reader.py query dictionaries/国家通用手语词典/国家通用手语词典.mdx 跳舞 --summary
python scripts/search_cases.py 否定
python scripts/read_document.py '资料/手语歌讲义/暮色回响/《暮色回响》手语讲义定稿（已校）.docx'
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




## 云端词典下载

接管时运行 `python scripts/download_dictionaries.py --default`。查看目录运行 `python scripts/download_dictionaries.py --list`；用户选择后运行 `python scripts/download_dictionaries.py --dictionary "美术常用词通用手语"`。

脚本仅使用标准库，读取 `dictionaries/resources.json` 的公开地址，不需要云端凭证。已有完整文件跳过，下载中断保留临时文件供重试；同名但不同长度的已有文件保留并要求确认。

## 更新修改案例目录

每个案例的“编号、手语歌、何时参考、核心辨析、关联规则”是目录的数据来源，正文单独保存。新增或修改后执行`python scripts/build_case_catalog.py`，自动重建`资料/规则与例句数据库/转写辨析案例.xlsx`，无需手动改两份。脚本只使用Python标准库。检索本曲转写规则示例：`python scripts/search_cases.py --rule 太阳-R01`。

更新资料后运行 `python scripts/build_rule_index.py`，自动生成规则到案例的 JSON 索引；只依赖标准库。索引为程序数据；老师在查询窗口双向检索。修改案例目录生成命令同时更新各歌曲的 `.content`，不再手动维护目录链接。

运行 `python scripts/build_query_page.py` 可更新内部界面 `scripts/query_page.html`；规则索引生成命令也会同步更新此页。页面由 Excel、规则正文和案例生成，不作为新的资料来源，不手工修改页面数据。

桌面查询使用 `python scripts/desktop_query.py`，或双击 `资料/规则与例句查询程序.pyw`。需 pywebview 与 Windows WebView2，由智能体自动检查和安装；窗口只查询，所有原件交给系统默认应用，运行时不调用模型。

桌面依赖自动安装命令：`python -m pip install -r scripts/desktop-requirements.txt`。Python 3.13 必须使用兼容的 typing_extensions；窗口使用系统 WebView2，缓存放在本机 LocalAppData/SignSongAgent，不写入教学资料。

桌面入口为 `资料/规则与例句查询程序.pyw`，内部 HTML 位于 `scripts/query_page.html`，由脚本生成，老师无需打开。
