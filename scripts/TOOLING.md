# 手语歌智能工坊：工具说明

编号速查：`R01`公共转写规则，`E0001`转写例句，`D001`教材会话；`太阳-R01`和`太阳-M001`是太阳歌曲的本曲转写规则与修改案例。

供智能体在调用工具、配置依赖和处理文档技术问题时查阅。备课流程与翻译要求见根目录`AGENTS.md`。

新歌先按AGENTS.md完整阅读参照讲义与小班原件，再检索数据库。`read_document.py`读取完整文本、Word格式和批注；它不能显示图示与完整版面，必须同时查看原件页面。查词使用`mdict_reader.py query`，整词失败后按AGENTS.md拆分对象、动作和关系，继续别名与近义词检索。查询现在自动展开别名指向，循环和断链会明确报告；原始记录与图片地址可用于本曲查询依据，不替代语义判断。

内置指拼沿用同一工具：`python scripts/mdict_reader.py query dictionaries/指拼 "X H" --summary`；`python scripts/mdict_reader.py images dictionaries/指拼 dictionaries/指拼 "M C T Z" "工作台/歌题/指拼原图"`。本地资源无MDD，第二个路径仍填指拼目录，便于沿用现有命令结构。`ZH`、`CH`、`SH`、`NG`是单个原图条目，首字母序列必须用空格分开；字母顺序及重复保留。将输出图片路径写入已确认的制作JSON，交给`make_handout.py`按正常图片流程制作。程序不根据汉字擅自决定名字读音。

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

新图解统一将已确认内容写成version: 1的制作JSON，结构与示例见下文。转写字符串中的`*文字*`生成斜体，`~~文字~~`生成删除线，单个`~`原样保留，`++`为正体；这些标记仍是当前转写格式的一部分。图片放入images数组，每个词选哪幅图由老师或助手先确定。

```powershell
python scripts/make_handout.py '<工作台/歌题/内容.json>' --out '<工作台/歌题/讲义.docx>' --height 0.79
python scripts/audit_docx.py '<工作台/歌题/讲义.docx>' --source '<工作台/歌题/内容.json>'
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

制作前按选定参照填写本曲版式JSON，再传入 `make_handout.py --layout <版式.json>`。可配置：`font_name`（字体）、`font_size`（正文字号）、`title_size`（标题字号）、`heading_size`（小标题字号）、`image_height`（英寸）、`margin_horizontal`与`margin_vertical`（厘米）、`line_spacing`（正文行距倍数）。例如：`{"font_name":"SimSun","font_size":12,"image_height":0.79,"margin_horizontal":2}`。未填项使用默认值；`--height`可覆盖图高。制作JSON标题level支持1至3级。段落样式、复杂表格等需要精确沿用时，直接局部编辑原Word，不为每首歌复制一套生成脚本。

## 当前制作输入 3.1.1

完整版默认读取制作JSON，亦可显式加`--input-format json`。试排与正式稿共用make_handout.py；不要另写一套Word生成器。顶层`version: 1`、`blocks`数组；标题块使用`heading`和可选`level`（1—3）；正文使用`text`，或`lyrics`与`transcription`。每块可有唯一`id`。图片`images`数组的`path`相对制作JSON所在目录，直接指向老师已有截图或裁图；程序不追查词典、替换图或改写图片。`entry`、`sense`、`source`是内部追溯信息，`action`记录实际采用的动作部分，`alt`是无障碍替代文本，均不自动印成标签。图旁`note`是已确认的必要短说明，排在对应图示后；`record`只保留长解释，永不打印。不要把长记录填入note，或把数据中的候选冒充老师确认。

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

检查入口使用inspect_docx与audit_docx.py命令。错误报告结构或指定输入不一致，警告报告可能的版式问题，人工核对列出机器不能确认的动作与页面问题。不传--source时可只读检查任意既有Word，不要求补制作数据；传--source时只接受当前制作JSON，核对正文、短注释、图片数、裁图与选图。检查不自动修改文件。

完整版不提供--update-action、本机接管状态记录和旧Markdown读取；已有项目请使用对应局部更新包，它保留旧格式兼容与接管检查。

参考图虚线框留边可在版式JSON中设置reference_frame_padding:true；只给包含参考图框的图片行增加至少“图高＋6磅”的行高并内收图框，普通图片行不变，不改原图像素或图高比例。默认false；可能影响分页，须逐页核对。

## 按内容配置格式与原件对照副本

以下版式字段按本次要求选用，未指定项使用默认值。image_space_before、image_space_after以磅设置图片段前段后，允许0；image_line_spacing为图片行距倍数。content_styles按角色设置font_name、font_size、space_before、space_after、line_spacing；JSON正文块role指定角色，未指定用body，短注释用note。未提供的字段保留已有默认值。例如content_styles中的intro设置楷体，body设置宋体，note设置9磅；不要按歌曲名写代码分支。图旁note仅应用字体字号，不能给同一行的不同run设置不同段间距。

精细沿用已有Word，显式使用make_handout.py 原件.docx --input-format docx --out 新副本.docx --layout 调整.json。副本模式保留全部原图、尺寸、裁剪、间隔、页面宽度、样式和页眉页脚，只修改word/document.xml中明确指定的段落。paragraph_overrides以从1开始的正文段落序号为键，使用上述五个格式字段；段落序号与原件绑定，原件结构变化后先重新核对。不能给此模式传全局图高、边距或字体，也不推测段落角色；没有覆盖项即原样复制。

reference_frame_rows指定确实需要修正的参考图段落序号；reference_frame_padding_pt控制额外行高（磅，最小可用留边须逐页验证）。只处理指定行中已有虚线图框，保留图片宽高及横向间隔。新稿reference_frame_padding:true也只增加含参考图框行的行高，普通图行保持原有间距。未指定额外留边时沿用6磅，可明确减小。不要把某一份参照的段落序号或字体要求写入通用默认值。
