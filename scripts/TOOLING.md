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

## 批量查词与参照版式

同一次查词可在一个命令中传入多个词，例如：`python scripts/mdict_reader.py query <词典.mdx> 天空 高 住 --summary`。同一读取器复用词条索引与最近8个解压块，避免逐词重复打开和解析词典。含空格的指拼序列仍作为一个带引号的参数传入。

例句与辨析案例按关键词出现次数分别排序、交替展示，避免一种结果占满；这只是文本匹配排序，智能体仍须判断实际适用性。规则反查与索引共用关联逻辑，包括条目标注和规则正文提及。

制作前按选定参照填写本曲版式JSON，再传入 `make_handout.py --layout <版式.json>`。可配置：`font_name`（字体）、`font_size`（正文字号）、`title_size`（标题字号）、`heading_size`（小标题字号）、`image_height`（英寸）、`margin_horizontal`与`margin_vertical`（厘米）、`line_spacing`（正文行距倍数）。例如：`{"font_name":"SimSun","font_size":12,"image_height":0.79,"margin_horizontal":2}`。未填项使用默认值；`--height`可覆盖图高。支持一级至三级Markdown标题。段落样式、复杂表格等需要精确沿用时，直接局部编辑原Word，不为每首歌复制一套生成脚本。

## 制作输入可选扩展 3.1.1

旧Markdown命令、标记、图片语法、--height、--layout与全部原版式字段不变，默认仍按Markdown读取与原样排版。试排与正式稿共用make_handout.py；文本读取、数据检查、add_picture等排版函数可复用，但不要另写一套Word生成器。

只有显式加`--input-format json`才读取制作JSON。顶层`version: 1`、`blocks`数组；标题块使用`heading`和可选`level`（1—3）；正文使用`text`，或`lyrics`与`transcription`。每块可有唯一`id`。图片`images`数组的`path`相对制作JSON所在目录，直接指向老师已有截图或裁图；程序不追查词典、替换图或改写图片。`entry`、`sense`、`source`是内部追溯信息，`action`记录实际采用的动作部分，`alt`是无障碍替代文本，均不自动印成标签。图旁`note`是已确认的必要短说明，排在对应图示后；`record`只保留长解释，永不打印。不要把长记录填入note，或把数据中的候选冒充老师确认。

```json
{"version":1,"blocks":[
 {"heading":"合成示例"},
 {"id":"a","lyrics":"示例句","transcription":"*动作①++*/~~省略②~~","images":[
  {"path":"已有截图.png","entry":"内部词条名","sense":"①","source":"原件出处","action":"实际采用的动作部分","note":"方向调整","record":"完整解释留在备课记录","reference":true,"crop":[0,0,100,100]}
 ]},
 {"repeat_of":"a","omit_images":true,"note":"重复段沿用前述图示"}
]}
```

示例裁图坐标须按实际原图改写：crop为原图像素范围[left, top, right, bottom]，右、下边界不包含；没有crop就保留整张已选截图。Word用原生裁剪，不改磁盘原图，保持所取区域比例。坐标合法不代表未裁掉动作，制作前必须看原图；reference:true保留原生参考图虚线框。

repeat_of只引用前面已定义的块id；自动复用原句、转写和选图，不能再存一份正文。omit_images:true明确省图；不设或false则复用图片。需要刻意不同的方案时另写独立块，不强行同步。块级note用于已要求的重复段说明或短注，完整解释放record。图片名不影响标签；要显示短说明必须明确提供note，不自动生成栏目。

```powershell
python scripts/make_handout.py '<内容.json>' --input-format json --out '<新稿.docx>' --layout '<原版式.json>'
python scripts/audit_docx.py '<新稿.docx>' --source '<内容.json>'
python scripts/audit_docx.py '<旧稿.docx>'
```

检查入口旧调用与audit(path,height)返回值保留。错误报告已确认的结构或指定输入不一致，警告报告可能的版式问题，人工核对列出机器不能确认的动作与页面问题；旧项目无新数据字段仍可检查，不要求补字段。只有明确传--source才按指定制作输入核对正文、短注释、图片数和裁图；旧Markdown用--input-format markdown。所有检查只读，不自动修改Word、PDF或歌曲。图片统计只输出到内部检查。

## 覆盖更新接管工具 3.1.1

更新说明见根目录AGENTS.md标记区块。`make_handout.py --update-action check`不写教学资料，以临时合成稿核对兼容，读取本机记录；`--update-action fail --update-note "已完成；未完成及原因"`保存未完成状态；`--update-action complete --update-note "实际调整与验证结果"`由智能体在全部接管验证成功后调用。complete再次检查兼容，保存本机记录及清理前AGENTS备份，且仅清理本版本区块；本机结果不会随更新包分发。失败保留说明，修复后可重试；重新覆盖带回同版说明也会核对文件指纹。

这不是安装器，也不自动推断任务阶段或下载词典。接管智能体负责先核对任务、已确认选择及所需环境，必要配置修改先备份；完成机器兼容检查、任务与必要环境核对后才调用complete；若本机进行试排，记录实际页面检查结果，缺渲染器不得虚报。制作交付仍必须逐页检查。交付源码与更新包不得提前清理区块。测试仅在`python -m unittest discover -s scripts/tests -v`下用临时合成数据运行，不触碰歌曲工作台。

参考图虚线框留边：旧Markdown默认保持原排版；遇到上边截断，可在本曲版式JSON中明确加`"reference_frame_padding": true`。该选项把图片行设置为至少“图高＋6磅”的行高并内收图框，不改图高、比例或原图像素；它可能影响分页，须按完整页面试排确认。默认false，不会在覆盖更新时自动修改旧版式文件。新旧内容输入均可选择启用。

制作代码职责：make_handout.py负责制作输入与可复用Word排版，audit_docx.py负责只读检查，handout_update.py仅封装本机接管状态与兼容检查，无独立命令入口。后续先复用现有函数；新增选项须显式启用，不能复制生成器或硬编码具体歌曲。更新模块也纳入版本指纹，模块变更会要求重新检查。
