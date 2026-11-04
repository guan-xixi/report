本实验旨在将开源代码审查工具 Co-Review 集成到 iCode 主 Agent 执行环境中，构建一套可重复、可追踪、能够比较不同记忆条件下代码安全审查表现的自动化实验流程。

实验的核心问题是：

在审查相同仓库、相同 Commit 和相同 review.patch 的条件下，Agent 在引入审查记忆之前、通过交互式问答补充知识之后，以及能够召回历史记忆之后，其漏洞识别与风险判断表现是否发生变化？

整个实验不是单纯比较三个不同 Prompt 的输出，而是尽可能控制代码输入、审查范围和运行环境，通过三种模式构造不同的知识条件，并验证记忆是否真实写入、保存和召回。

实验需要区分三个层次：
1. 流程自动化： 自动读取实验配置、准备代码仓库、执行 F/G/H 三种模式、保存运行记录。
2. 记忆机制验证： 验证初始记忆、交互后的知识新增，以及后续运行中的记忆召回。
3. 审查效果评估： 只有当 iCode 主 Agent 实际完成审查并生成有效的正式报告后，才能将审查结论用于效果对比。
第三点尤其重要：记忆写入成功、记忆召回成功，不等于代码审查已经成功，更不等于漏洞检测效果已经得到验证。
二、整体技术架构
实验配置与样本管理

bench_info.xlsx：仓库名、Commit、补充线索、预期审查目标


样本准备与环境初始化

读取压缩包 → 解压仓库 → 定位 review.patch → 快照记忆状态


F
初始基线

silent

初始记忆状态


G
交互补充

ask

补充线索并验证知识


H
记忆后审查

silent

在更新后记忆下审查




iCode 主 Agent + Co-Review Skill

真实执行审查工作流、分析补丁、调用相关工具并生成审查结论


结果校验与实验归档

正式报告校验 → F/G/H 结果回填 → JSONL、日志及记忆证据归档


整体采用“实验配置驱动—样本准备—分模式执行—记忆状态控制—正式报告校验—结果回填”的流水线设计。

这里需要明确：Python 脚本承担实验编排与结果管理；iCode 主 Agent 承担实际审查任务；Co-Review Skill 提供代码审查工作流及其相关能力。不能把 Python 自动准备材料、写入记忆或生成日志，直接等同于 Agent 已经完成了审查。
三、实验环境与目录组织
3.1 关键文件和目录
根据目前的实验配置，核心路径如下。
组成	路径	作用
iCode 项目	/data/cyjia/iCode/	Agent 执行环境
iCode CLI	/data/cyjia/iCode/icode	启动主 Agent
Python 环境	/data/cyjia/iCode/.venv/bin/python	运行自动化实验脚本
实验 Excel	/data/cyjia/iCode/bench_info.xlsx	保存样本信息和 F/G/H 结果
Co-Review 项目	/data/cyjia/co-review/	审查工作流代码及资源
Co-Review Skill	/data/cyjia/co-review/SKILL.md	向 Agent 指定审查技能
批量实验脚本	/data/cyjia/co-review_bench/bench_fill_full.py	自动化实验主入口
样本压缩包	/tmp/wazuh-cwe78-44bf114.zip	当前配置的 Wazuh 样本
工作目录	/tmp/co-review-bench/	解压工作区及实验运行材料
运行记录	/tmp/co-review-bench/bench_runs.jsonl	保存逐次实验记录
日志目录	/tmp/co-review-bench/logs/	保存标准输出、错误输出等日志

以上是配置中使用的路径，不代表当前服务器上所有文件都已被实际检查或验证。
3.2 Excel 实验配置
Excel 是样本管理和实验结果的主要入口。目前约定的列含义如下。
列	含义	用途
A	仓库名	确定被测项目
B	Commit	标识待审查的代码版本
C	Missing 标记	记录样本是否缺少相关信息
D	Missing Info / 补充线索	为交互模式提供待核查的审查线索
E	Expected Issue	记录预期检查的安全问题或审查目标
F	Initial Silent	初始记忆下的 silent 模式结果
G	Ask	交互式 ask 模式结果
H	After-Memory Silent	记忆更新后的 silent 模式结果

当前示例为 Wazuh 的 CWE-78 命令注入样本，Excel 第 2 行此前配置的信息包括：
- Commit：e6b41db5cdf815540cf27f327a536b5986648516
- D 列线索：CWE-78 命令注入触发条件及外部数据来源
- E 列目标：判断是否存在可利用的命令注入路径
这里的 D 列内容是需要 Agent 核查的线索，不是已经证实的漏洞事实；E 列是审查目标，也不能直接替代最终的代码证据。
3.3 当前样本映射
当前脚本的样本包映射只配置了一个样本：
PACKAGE_MAP = {    "wazuh-cwe78-44bf114.zip": 2}



含义是：将 wazuh-cwe78-44bf114.zip 与 Excel 第 2 行关联。

因此，即使命令指定 --rows 2,3，第 3 行也不会自动获得对应的样本压缩包。要扩展到其他样本，需要补充映射，并确保每个压缩包、Excel 行、Commit 和 review.patch 的对应关系正确。
四、实验输入与代码审查对象的准备
4.1 为什么使用 review.patch
实验以指定 Commit 对应的代码变更作为审查对象，而不是让 Agent 随意浏览整个项目后自行决定检查范围。

核心输入包括：
- 仓库代码：提供函数调用、上下文、依赖关系等必要信息。
- Commit：确定代码变更版本。
- review.patch：记录此次 Commit 的代码差异。
- D 列补充线索：在 G 模式中提供待核查的信息。
- E 列预期审查目标：用于明确需要检查的安全问题。
review.patch 是主要审查入口，但不一定包含判断漏洞所需的全部上下文。因此，Agent 仍可能需要查看仓库中的相关源代码、调用链和保护逻辑，不能简单地将补丁文本视为全部证据。
4.2 样本准备流程
1. 读取 Excel，获取指定行的仓库名、Commit、补充线索和预期审查目标。
2. 根据 PACKAGE_MAP 找到对应的 ZIP 或 TAR 压缩包。
3. 将压缩包解压到实验工作区，并尝试识别仓库根目录。
4. 在仓库中定位 review.patch，确认实际审查输入存在。
5. 记录初始记忆文件状态，为 F/G/H 运行前的状态恢复提供依据。
6. 创建当前实验的运行目录、任务 Prompt 和日志文件。
其中，仓库根目录的自动识别与 review.patch 的定位是重要检查点。如果压缩包内部层级与脚本假设不一致，或者没有找到补丁文件，就不应继续把该样本当成有效实验。
4.3 解压与复用规则
脚本支持通过 --keep-extracted 复用已有的解压目录。若不使用该选项，按照当前脚本逻辑，会重新解压相应样本。

这意味着：
- 复用目录可以减少重复解压时间。
- 复用时必须确认代码版本、补丁和仓库内容没有被上次实验修改。
- 不能只因为目录存在，就认定其对应正确的 Commit。
- 如果 Co-Review 或 Agent 会修改工作区，必须额外检查工作区是否受到污染。
五、F/G/H 三种实验模式的设计
这是整个实验最核心的部分。三种模式使用同一个样本，区别在于交互方式和记忆状态。
F

Initial Silent：初始基线
silent
在初始记忆状态下执行代码审查，不通过交互式问答补充线索。记录 Agent 原始审查能力、初始记忆条件下的风险判断以及实际使用的记忆。

主要回答：没有本轮交互补充时，Agent 能否发现目标问题？



G

Ask：交互补充与记忆演化
ask
通过 ask 模式执行审查。Excel D 列提供预先准备的补充线索，Agent 应结合实际代码进行核查；只有经过验证、满足记忆写入规则的信息，才应作为候选知识进入记忆库。

主要回答：交互式补充能否帮助 Agent 发现问题，并形成可供后续使用的知识？



H

After-Memory Silent：记忆后审查
silent
使用 G 阶段形成的记忆状态重新执行审查，不再依赖本轮 ask 交互。需要验证相关记忆确实被召回，并且对最终判断产生了可追踪的作用。

主要回答：交互形成的知识能否被后续审查复用？



5.1 F：建立基线
F 模式的目标是获得一个可比较的起点。

具体要求：
1. 将记忆恢复到实验约定的初始状态。
2. 以 silent 模式启动 iCode 主 Agent。
3. 让 Co-Review 对指定的 review.patch 开展真实审查。
4. 记录 Agent 的审查过程、正式报告、风险分类及记忆调用情况。
5. 不将 D 列线索直接注入成已经确认的结论。
F 不是“完全没有记忆”，而是不包含 G 阶段本轮新增知识的初始记忆条件。如果初始知识库本身有内容，就应记录其状态。
5.2 G：通过 ask 模式补充与验证知识
G 模式的目标是让 Agent 在交互式审查条件下处理补充线索，并检验知识能否进入记忆库。

具体要求：
1. 将记忆恢复到与 F 相同的初始基线。
2. 以 ask 模式执行审查。
3. 提供 Excel D 列的补充线索和 E 列的审查目标。
4. Agent 检查线索是否与实际代码、调用路径和保护逻辑相符。
5. 根据 Co-Review 的真实工作流处理知识写入。
6. 记录记忆条目写入前后的变化，以及新条目的具体内容和索引。
7. 保存正式审查报告及相应证据。
这里有两个容易混淆的概念：
- 补充线索：用户或实验配置提供的信息，可能正确，也可能不完整。
- 已验证知识：经过代码证据核查后，满足写入条件、可以供后续审查参考的信息。
不能因为 D 列写着 CWE-78，就直接把“存在命令注入”作为 confirmed 记忆写入。否则实验测到的可能是把预期答案提前提供给 Agent 的效果，而不是 Agent 自主验证和演化知识的能力。

另外，ask 表示交互模式，并不意味着在无人工参与的 CLI 运行中，系统一定会自动停下来等待真人回答。对于 headless 自动化执行，必须确认 Co-Review 的工作流实际如何处理交互请求。
5.3 H：检验记忆能否被后续任务复用
H 模式主要检验 G 阶段产生的记忆是否具有可复用性。

具体要求：
1. 加载 G 阶段结束后的有效记忆状态。
2. 不再通过 ask 模式提供同一轮交互补充。
3. 使用 silent 模式启动 iCode 主 Agent。
4. 记录召回是否命中，以及命中的记忆条目索引。
5. 检查召回的记忆是否与当前代码审查任务相关。
6. 检查 Agent 是否真正使用记忆中的知识分析代码，而非仅仅发生一次检索调用。
7. 保存正式报告，并与 F、G 的审查结果进行比较。
H 阶段必须同时关注三件事：
- Recall Hit：相关记忆有没有被召回？
- Evidence Use：Agent 有没有将记忆用于实际代码分析？
- Outcome Change：最终审查结论是否发生变化，变化是否有代码证据支持？
即使召回命中，也不代表记忆必然有帮助；即使最终结论发生变化，也不能直接认定是记忆导致的，需要结合审查过程和证据分析。
5.4 三种模式的对照关系
对比维度	F：初始 silent	G：ask	H：记忆后 silent
初始记忆	基线	同一基线	G 阶段的演化后状态
交互方式	silent	ask	silent
本轮补充线索	不主动提供 D 列线索	提供并核查 D 列线索	不依赖本轮 ask 交互
知识写入	不应引入 G 的新增知识	验证后允许新增	读取 G 的有效结果
记忆召回	记录初始召回	记录相关召回与写入	重点验证新增记忆的召回
正式审查报告	必须	必须	必须
主要用途	建立基线	观察交互和知识演化	检验记忆复用

需要注意：F 与 G 的差异不只有记忆，交互方式也不同；H 与 F 的差异主要在记忆状态，但实际实验还需要控制执行顺序、代码工作区和工具环境等因素。因此，三模式对照能够提供证据，但不能仅凭一次结果就证明记忆是性能变化的唯一原因。
六、自动化实验脚本的内部执行流程
当前主脚本为：

/data/cyjia/co-review_bench/bench_fill_full.py

它的职责是组织实验，不是取代 iCode 主 Agent 执行审查。
6.1 主流程
#chatgpt-mermaid-_r_653_{font-family:-apple-system-body,ui-sans-serif,-apple-system,system-ui,"Segoe UI",Helvetica,"Apple Color Emoji",Arial,sans-serif,"Segoe UI Emoji","Segoe UI Symbol";font-size:16px;fill:rgb(13, 13, 13);}@keyframes edge-animation-frame{from{stroke-dashoffset:0;}}@keyframes dash{to{stroke-dashoffset:0;}}#chatgpt-mermaid-_r_653_ .edge-animation-slow{stroke-dasharray:9,5!important;stroke-dashoffset:900;animation:dash 50s linear infinite;stroke-linecap:round;}#chatgpt-mermaid-_r_653_ .edge-animation-fast{stroke-dasharray:9,5!important;stroke-dashoffset:900;animation:dash 20s linear infinite;stroke-linecap:round;}#chatgpt-mermaid-_r_653_ .error-icon{fill:rgb(243, 243, 243);}#chatgpt-mermaid-_r_653_ .error-text{fill:rgb(13, 13, 13);stroke:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .edge-thickness-normal{stroke-width:1px;}#chatgpt-mermaid-_r_653_ .edge-thickness-thick{stroke-width:3.5px;}#chatgpt-mermaid-_r_653_ .edge-pattern-solid{stroke-dasharray:0;}#chatgpt-mermaid-_r_653_ .edge-thickness-invisible{stroke-width:0;fill:none;}#chatgpt-mermaid-_r_653_ .edge-pattern-dashed{stroke-dasharray:3;}#chatgpt-mermaid-_r_653_ .edge-pattern-dotted{stroke-dasharray:2;}#chatgpt-mermaid-_r_653_ .marker{fill:rgb(143, 143, 143);stroke:rgb(143, 143, 143);}#chatgpt-mermaid-_r_653_ .marker.cross{stroke:rgb(143, 143, 143);}#chatgpt-mermaid-_r_653_ svg{font-family:-apple-system-body,ui-sans-serif,-apple-system,system-ui,"Segoe UI",Helvetica,"Apple Color Emoji",Arial,sans-serif,"Segoe UI Emoji","Segoe UI Symbol";font-size:16px;}#chatgpt-mermaid-_r_653_ p{margin:0;}#chatgpt-mermaid-_r_653_ .label{font-family:-apple-system-body,ui-sans-serif,-apple-system,system-ui,"Segoe UI",Helvetica,"Apple Color Emoji",Arial,sans-serif,"Segoe UI Emoji","Segoe UI Symbol";color:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .cluster-label text{fill:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .cluster-label span{color:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .cluster-label span p{background-color:transparent;}#chatgpt-mermaid-_r_653_ .label text,#chatgpt-mermaid-_r_653_ span{fill:rgb(13, 13, 13);color:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .node rect,#chatgpt-mermaid-_r_653_ .node circle,#chatgpt-mermaid-_r_653_ .node ellipse,#chatgpt-mermaid-_r_653_ .node polygon,#chatgpt-mermaid-_r_653_ .node path{fill:rgb(222, 234, 251);stroke:rgb(83, 154, 248);stroke-width:1px;}#chatgpt-mermaid-_r_653_ .rough-node .label text,#chatgpt-mermaid-_r_653_ .node .label text,#chatgpt-mermaid-_r_653_ .image-shape .label,#chatgpt-mermaid-_r_653_ .icon-shape .label{text-anchor:middle;}#chatgpt-mermaid-_r_653_ .node .katex path{fill:#000;stroke:#000;stroke-width:1px;}#chatgpt-mermaid-_r_653_ .rough-node .label,#chatgpt-mermaid-_r_653_ .node .label,#chatgpt-mermaid-_r_653_ .image-shape .label,#chatgpt-mermaid-_r_653_ .icon-shape .label{text-align:center;}#chatgpt-mermaid-_r_653_ .node.clickable{cursor:pointer;}#chatgpt-mermaid-_r_653_ .root .anchor path{fill:rgb(143, 143, 143)!important;stroke-width:0;stroke:rgb(143, 143, 143);}#chatgpt-mermaid-_r_653_ .arrowheadPath{fill:rgb(143, 143, 143);}#chatgpt-mermaid-_r_653_ .edgePath .path{stroke:rgb(143, 143, 143);stroke-width:1px;}#chatgpt-mermaid-_r_653_ .flowchart-link{stroke:rgb(143, 143, 143);fill:none;}#chatgpt-mermaid-_r_653_ .edgeLabel{background-color:rgb(252, 252, 252);text-align:center;}#chatgpt-mermaid-_r_653_ .edgeLabel p{background-color:rgb(252, 252, 252);}#chatgpt-mermaid-_r_653_ .edgeLabel rect{opacity:0.5;background-color:rgb(252, 252, 252);fill:rgb(252, 252, 252);}#chatgpt-mermaid-_r_653_ .labelBkg{background-color:rgba(252, 252, 252, 0.5);}#chatgpt-mermaid-_r_653_ .cluster rect{fill:rgb(243, 243, 243);stroke:rgba(0, 0, 0, 0.1);stroke-width:1px;}#chatgpt-mermaid-_r_653_ .cluster text{fill:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ .cluster span{color:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ div.mermaidTooltip{position:absolute;text-align:center;max-width:200px;padding:2px;font-family:-apple-system-body,ui-sans-serif,-apple-system,system-ui,"Segoe UI",Helvetica,"Apple Color Emoji",Arial,sans-serif,"Segoe UI Emoji","Segoe UI Symbol";font-size:12px;background:rgb(243, 243, 243);border:1px solid rgba(0, 0, 0, 0.1);border-radius:2px;pointer-events:none;z-index:100;}#chatgpt-mermaid-_r_653_ .flowchartTitleText{text-anchor:middle;font-size:18px;fill:rgb(13, 13, 13);}#chatgpt-mermaid-_r_653_ rect.text{fill:none;stroke-width:0;}#chatgpt-mermaid-_r_653_ .icon-shape,#chatgpt-mermaid-_r_653_ .image-shape{background-color:rgb(252, 252, 252);text-align:center;}#chatgpt-mermaid-_r_653_ .icon-shape p,#chatgpt-mermaid-_r_653_ .image-shape p{background-color:rgb(252, 252, 252);padding:2px;}#chatgpt-mermaid-_r_653_ .icon-shape .label rect,#chatgpt-mermaid-_r_653_ .image-shape .label rect{opacity:0.5;background-color:rgb(252, 252, 252);fill:rgb(252, 252, 252);}#chatgpt-mermaid-_r_653_ .label-icon{display:inline-block;height:1em;overflow:visible;vertical-align:-0.125em;}#chatgpt-mermaid-_r_653_ .node .label-icon path{fill:currentColor;stroke:revert;stroke-width:revert;}#chatgpt-mermaid-_r_653_ .node .neo-node{stroke:rgb(83, 154, 248);}#chatgpt-mermaid-_r_653_ [data-look="neo"].node rect,#chatgpt-mermaid-_r_653_ [data-look="neo"].cluster rect,#chatgpt-mermaid-_r_653_ [data-look="neo"].node polygon{stroke:url(#chatgpt-mermaid-_r_653_-gradient);filter:drop-shadow( 1px 2px 2px rgba(185,185,185,1));}#chatgpt-mermaid-_r_653_ [data-look="neo"].swimlane.cluster rect{filter:none;}#chatgpt-mermaid-_r_653_ [data-look="neo"].node path{stroke:url(#chatgpt-mermaid-_r_653_-gradient);stroke-width:1px;}#chatgpt-mermaid-_r_653_ [data-look="neo"].node .outer-path{filter:drop-shadow( 1px 2px 2px rgba(185,185,185,1));}#chatgpt-mermaid-_r_653_ [data-look="neo"].node .neo-line path{stroke:rgb(83, 154, 248);filter:none;}#chatgpt-mermaid-_r_653_ [data-look="neo"].node circle{stroke:url(#chatgpt-mermaid-_r_653_-gradient);filter:drop-shadow( 1px 2px 2px rgba(185,185,185,1));}#chatgpt-mermaid-_r_653_ [data-look="neo"].node circle .state-start{fill:#000000;}#chatgpt-mermaid-_r_653_ [data-look="neo"].icon-shape .icon{fill:url(#chatgpt-mermaid-_r_653_-gradient);filter:drop-shadow( 1px 2px 2px rgba(185,185,185,1));}#chatgpt-mermaid-_r_653_ [data-look="neo"].icon-shape .icon-neo path{stroke:url(#chatgpt-mermaid-_r_653_-gradient);filter:drop-shadow( 1px 2px 2px rgba(185,185,185,1));}#chatgpt-mermaid-_r_653_ .node text{font-size:14px;font-weight:600;letter-spacing:normal;fill:rgb(0, 79, 153);}#chatgpt-mermaid-_r_653_ .edgeLabels text{font-size:13px;font-weight:600;letter-spacing:-0.08px;fill:rgb(0, 79, 153);}#chatgpt-mermaid-_r_653_ .node tspan[font-weight="normal"],#chatgpt-mermaid-_r_653_ .edgeLabels tspan[font-weight="normal"]{font-weight:600;}#chatgpt-mermaid-_r_653_ .edgeLabel .label rect{opacity:1;rx:13px;ry:13px;fill:rgb(245, 250, 255);stroke:rgb(206, 219, 229);stroke-width:1px;}#chatgpt-mermaid-_r_653_ .node rect,#chatgpt-mermaid-_r_653_ .node circle,#chatgpt-mermaid-_r_653_ .node ellipse,#chatgpt-mermaid-_r_653_ .node polygon,#chatgpt-mermaid-_r_653_ .node path{fill:rgb(229, 243, 255);stroke:rgba(0, 0, 0, 0.1);stroke-width:1px;}#chatgpt-mermaid-_r_653_ .node rect{rx:16px;ry:16px;}#chatgpt-mermaid-_r_653_ .mermaid-decision .label-container{fill:rgb(245, 250, 255);stroke:rgb(206, 219, 229);stroke-dasharray:2,2;}#chatgpt-mermaid-_r_653_ .edgePaths .flowchart-link{stroke:rgb(143, 143, 143);stroke-width:1px;stroke-linecap:round;stroke-linejoin:round;}#chatgpt-mermaid-_r_653_ .marker{fill:rgb(143, 143, 143);stroke:rgb(143, 143, 143);}#chatgpt-mermaid-_r_653_ :root{--mermaid-font-family:-apple-system-body,ui-sans-serif,-apple-system,system-ui,"Segoe UI",Helvetica,"Apple Color Emoji",Arial,sans-serif,"Segoe UI Emoji","Segoe UI Symbol";}读取命令行参数检查 Excel、CLI、Skill和样本包读取 Excel 指定行解压样本并定位 review.patch记录初始记忆状态恢复基线记忆F：silent 审查恢复基线记忆G：ask 审查与记忆演化保存 G 阶段后的记忆状态恢复 G 阶段结束状态H：silent 记忆后审查校验正式报告写入 JSONL、日志和 Excel恢复实验前基线记忆





6.2 参数读取和环境检查
脚本通过环境变量和命令行参数控制路径、运行模式和超时设置。

主要默认值如下：
BENCH_XLSX = "/data/cyjia/iCode/bench_info.xlsx"SHEET_NAME = "Sheet1"PKG_DIR = "/tmp"WORK_BASE = "/tmp/co-review-bench"ICODE_EXE = "/data/cyjia/iCode/icode"ICODE_AGENT = "Code"ICODE_TIMEOUT = 600COREVIEW_ROOT = "/data/cyjia/co-review"COREVIEW_SKILL = "/data/cyjia/co-review/SKILL.md"



运行 --check 时，脚本会检查以下内容：
- iCode CLI 是否存在。
- Excel 文件是否存在。
- 样本包目录是否存在。
- Co-Review 仓库与 SKILL.md 是否存在。
- icode run --help 是否包含 --agent、--json 和 --workdir 等必要参数。
但这些检查只能证明基本文件和 CLI 参数满足要求，不能证明 Co-Review 的 Hook、MCP、记忆调用和正式报告生成已经端到端跑通。
6.3 启动 iCode 主 Agent
脚本采用类似以下命令启动每次审查：
/data/cyjia/iCode/icode run \
  --task /path/to/task.md \
  --agent Code \
  --workdir /path/to/repo \
  --json


其中：
- run：启动一次 Agent 任务。
- --task：读取文件中的任务 Prompt。
- --agent Code：指定 Code Agent。
- --workdir：指定被审查的仓库工作目录。
- --json：要求 CLI 以 JSON 形式输出结构化执行结果。
脚本通过 Python 的子进程机制启动 CLI，并分别保存标准输出与标准错误输出。成功解析 JSON 后，可以提取 session_id、result 和 duration 等字段。

要特别区分两种输出：
1. CLI 执行结果：表示 iCode 返回了什么。
2. 正式审查报告：表示 Co-Review 工作流实际产生了什么审查结论。
CLI 返回成功，不一定意味着报告存在；报告文件存在，也不一定意味着其内容符合正式结果格式。
6.4 正式报告的定位与校验
当前脚本会尝试从以下位置查找正式报告：
- Agent JSON 结果中提供的报告路径。
- 当前运行目录中的报告文件。
- execution-state.json 中记录的相关路径或状态。
找到报告后，还需要检查报告内容是否包含可识别的 risk_type，以及该值是否符合脚本当前允许的枚举：
- confirmed_vulnerability
- potential
- potential_risk
- not_vulnerable
这只是脚本当前使用的校验规则，并不意味着已经确认这些值与 Co-Review 实际版本的报告 Schema 完全一致。正式执行前应以实际 Skill、工作流和报告样例为准。

推荐的报告校验至少分为四层：
校验层	检查内容	不通过时的处理
文件层	报告是否真实存在、是否可读取	标记报告缺失
格式层	JSON 或正式报告格式是否有效	标记格式错误
字段层	风险分类、必要字段是否存在且合法	标记字段无效
证据层	是否存在具体代码位置、数据流和判断依据	标记证据不足，不能直接计为有效审查

当前脚本明确实现了报告定位和部分 risk_type 校验；证据层的完整验证程度需要根据实际代码和输出报告进一步确认，不能假定已经全部实现。
6.5 结果回填
当运行状态为 SUCCESS 且正式报告通过当前校验时，脚本才会将风险分类等结果写回 Excel 对应的 F、G 或 H 列。

失败情况可能包括：
- CLI 运行失败。
- 子进程超时。
- 没有找到报告。
- 报告格式不正确。
- risk_type 缺失或不符合预设枚举。
这类情况应该保留为失败状态或错误信息，不能自动填成 not_vulnerable，也不能用预期答案替代实际输出。

Excel 是最终结果的汇总入口；JSONL 和日志则用于追溯每一次运行的过程与失败原因。
七、记忆隔离、状态传递与实验可重复性
7.1 为什么要隔离记忆
如果 F 运行后产生的记忆直接留给 G，G 的结果就可能受到 F 的影响；如果 G 的记忆没有正确传递给 H，H 就无法验证记忆复用。

因此，脚本试图构造如下状态关系：
基线状态 M₀

实验开始前记录的记忆状态


F：M₀ → F 的审查结果

F 的执行不应把新增记忆污染到 G 的起点


恢复 M₀ → G：M₀ → M₁

G 在基线记忆下运行，并形成候选演化状态 M₁


H：M₁ → H 的审查结果

使用 G 阶段结束后的记忆验证召回和实际使用


恢复 M₀

尽量避免当前样本污染后续样本的实验环境


这里的 M₀ 和 M₁ 是对实验状态的抽象表示，并不是某个具体 JSON 文件名。
7.2 当前实现的边界
当前脚本会扫描文件名包含 knowledge 或 memory 的 JSON 文件，并尝试进行快照与恢复。

这是一种初步的隔离策略，但存在明确限制：
- 可能漏掉不包含这些关键词的记忆文件。
- 可能漏掉 SQLite 数据库、其他格式文件或独立服务端存储。
- 如果运行期间创建了新的记忆文件，而恢复逻辑只恢复原先快照中已有的文件，新文件可能残留。
- 如果记忆由独立 MCP 服务或持久化服务管理，单纯恢复本地文件不一定能恢复真实状态。
- 如果 Agent 修改了仓库代码或配置文件，只恢复记忆并不能保证整个实验环境完全一致。
所以，汇报时可以说“脚本实现了对指定记忆文件的快照与恢复”，但在没有进一步验证前，不宜宣称“已经实现所有记忆后端的完全隔离”。
7.3 如何判断记忆确实被传递
建议为每次运行保存以下证据：
证据	用途
运行前记忆快照	确定初始状态
G 运行前后条目数	观察知识条目变化
新增条目的内容及索引	确认写入了什么
H 的召回命中标志	确认是否发生召回
H 命中的记忆索引	确认是否命中新产生的知识
H 的实际推理与代码证据	判断记忆是否被使用
最终正式报告	判断审查结论及其变化

此前已有的前置自动化记录中，曾观察到 F 使用记忆索引 [1]，G 的记忆条目从 1 条增至 2 条，H 的召回命中为 true 且命中索引 [1, 2]。这些记录可以支持“曾观察到记忆状态变化和召回”的表述，但不能自动证明当前这版完整脚本已经端到端执行成功，也不能证明正式审查报告已生成。
