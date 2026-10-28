# 项目维护说明（MAINTAINING）

> 给接手维护这个插件的人（或新的 DSH 工作区）看的。日常使用看 [README](README.md)，这里只讲**怎么改、怎么测、怎么发**，以及**踩过的坑**。
>
> 最后更新：2026-10-29 · 版本 1.4.0

---

## 1. 这是什么

`@andrewveraside/dsh-wps-sidebar` —— DSH 的 WPS 办公插件（Windows-only）。让 AI 直接在用户**打开的 WPS 文字 / 表格 / 演示窗口里**改文档：改动实时出现在窗口里，用户可以同时编辑，AI 也能渲染成图自己核对。

它不是从零写的，是 **fork 自 `@copylee/dsh-office` 1.2.0**（MIT），做了两件事：

1. **改名 + 钉死驱动 WPS**（本机 WPS 抢注了微软 Office 的 ProgID，见 §6）
2. **补齐原作者没做的功能**：Word 1–5、Excel 6–9、版本存档

## 2. 目录与关键文件

| 路径 | 作用 | 改不改 |
|---|---|---|
| `lib/index.js` | 插件主体（19,641 行 ESM，未混淆）：注册工具、调 helper、拼提示词 | 加工具 / 改 op 时改 |
| `lib/client.js` | 侧栏前端（640 行：设置页、状态卡片） | 改 UI 时改 |
| `helper/OfficeHelper.cs` | **核心**（11,212 行 C# 5）：COM 后期绑定操作 WPS/Office | 所有文档操作的真实实现在这 |
| `cordis.patch.yml` | 插件默认配置（`suite: wps` 硬钉） | 改默认设置时改 |
| `package.json` | 版本号、`files` 白名单、`dsh` 元数据 | 发版时改版本号 |
| `README.md` | 用户文档 | 加功能时补 |
| `scripts/build-helper.ps1` | 编译 helper（**入库**，见 §3.1） | 改编译参数时改 |
| `scripts/office-smoke.mjs` | 冒烟测试（**入库**，见 §3.3） | 加 op / 改协议时补 |
| `scripts/lib/helper.cjs` | 冒烟测试用的最小 helper 客户端 | 改协议（stdio）时改 |

### `helper/OfficeHelper.cs` 内部结构（加 op 时按这个找位置）

| 区域 | 行号（1.4.0 时，会随代码增长偏移） | 说明 |
|---|---|---|
| 协议分发 `Handle(...)` | ~870–930 | `cmd` 路由：`open` / `read` / `edit` / `render` / `save` / `versions` / `revert` / `close` / `status` |
| 应用与文档定位 | ~760–830 | `App()` `Docs()` `Find()` `Doc()` —— **`Doc()` 认的是 `doc` 字段，不是 `path`** |
| 打开 / 保存 | ~981–1120 | `Open()` / `Save()`（存档钩子在这） |
| 版本存档 | ~1300–1420 | `VersionDir` / `Snapshot` / `ListVersions` / `RevertVersion` / `SnapshotNow` |
| Word 操作 | **4038 起** | `insert_paragraphs` 起，`set_text` 在 4134 |
| Excel 操作 | **5331 起** | `write_range` 起，`refresh_pivot` 6567、`conditional_format` 6748 |
| PowerPoint 操作 | **7663 起** | `add_slide` 起，`theme` 7676 |

加一个新 `office_edit` op 的套路：
1. 在 `helper` 的 `Edit()` 分派链里加一个 `if (type == "你的op") return YourOp(...);`
2. 写 `static string YourOp(dynamic doc, Bag a)`，用 `a.Need("x")` 取必填、`a.Str("x", 默认)` 取可选
3. 在 `lib/index.js` 的 `office_help` 文案和 README 里补上
4. 编译 → 用 `scripts/office-smoke.mjs` 实测（或照它的写法加一步）

加一个**新工具**（不只是 op）的套路：在 `lib/index.js` 里 `tools.push(defineTool({...}))`，照抄 `office_versions` 的写法（参数用 `DOC` / `APP` 复用、`output`、`timeoutMs`、`execute`、`presentCall`），然后调 `helper.call("<cmd>", ...)`，并在 `helper` 的 `Handle()` 里加对应 `cmd` 分支。

## 3. 开发循环（每次改代码都走这条）

```
改源码 → 编译 helper → 更新测试脚本里的 hash → 实测 → 打包 → 安装 → 重启 DSH
```

### 3.1 编译 helper

```powershell
cd E:\DshData\dsh-office-wps
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\build-helper.ps1
```

脚本做的事：按源码 SHA256 前 12 位算 hash → kill 掉占用 exe 的 `office-helper.exe` → 用系统 `csc.exe` 编到 `%LOCALAPPDATA%\dsh-office\helper-<hash>\office-helper.exe`。缓存目录写不进去时自动改在 `%TEMP%\dsh-helper-build-<PID>\` 编好再拷回来。

> ⚠️ **坑 1：脚本必须指向工作副本**。曾指向 `node_modules` 里的安装副本，改了源码 exe 毫无变化，白排查半小时。现在默认 `-RepoRoot` 就是仓库根，指向 `helper\OfficeHelper.cs`；用 `-RepoRoot` 参数可指向别的 checkout。
>
> ⚠️ **坑 2：exe 被占用会编译失败**（`CS0016: 未能写入输出文件`）。脚本会自动 kill，但如果 DSH 正在跑，kill 后 DSH 下次调用会重新拉起——所以改完记得重启 DSH。

### 3.2 验证 exe 真的编进了新代码

```bash
"C:/Users/23171/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe" -c "
import os
p=os.path.expandvars(r'%LOCALAPPDATA%\dsh-office\helper-<hash>\office-helper.exe')
data=open(p,'rb').read()
for k in ['你的新字符串','VersionsFolder']:
    print(k, 'UTF16:', 'FOUND' if k.encode('utf-16-le') in data else 'absent')
"
```

> ⚠️ **坑 3：必须用 UTF-16-LE 扫**。.NET 的字符串常量是 UTF-16 存的，用 ASCII/`strings` 扫**永远 absent**，会让你误以为代码没编进去。

### 3.3 实测

```bash
cd /e/DshData/dsh-office-wps
node scripts/office-smoke.mjs            # = npm run smoke；退出码 0 = 全过
node scripts/office-smoke.mjs --keep     # 留下 %TEMP%\dsh-smoke-* 里的 docx / png / 版本，便于人眼看
```

它在一个临时目录造 docx，然后**在同一个 helper 进程里**跑完整生命周期并逐条断言：

`status → open(how=opened) → read(读到夹具文字) → edit(insert_paragraphs) → read(改动可见) → render(必须给 out，校验 png >1KB) → save → versions(有快照) → revert(最新那份) → read(改动已回滚) → close`，最后删临时目录。

约 4 秒跑完。**改 helper 协议 / 加 op 后先跑它**，比在 DSH 里手点一遍快得多。`_test/` 下那些老探针是本地残留、不入库（`.gitignore` 已排除），需要一次性深挖时照写。

> ⚠️ **坑 4（最坑）**：heredoc 里写 Windows 路径的**反斜杠会被吃掉**。`'E:\\DshData\\x\\_vtest\\v.docx'` 里的 `\v`、`\t` 会被 Node 字符串字面量当转义序列，wps 拿到 `E:DshDatadsh-office-wps_vtest\test.docx`，报中文 COM 错误「**路径中具有非法字符**」。看起来像插件 bug，其实是测试脚本的问题。**一律用正斜杠 `E:/xxx/y.docx`**。
>
> ⚠️ **坑 4b：stdio 只有一种摆法能跑（2026-10-29 逐个探过）**。测试客户端把 helper 的输出**重定向到文件**、自己轮询**同一个读写句柄**；命令**走真管道**（`stdio:['pipe', outFd, errFd]` + `proc.stdin.write`）。踩过的三种摆法全部失败：
> - stdin 给**文件**句柄：helper 打完 `{"event":"ready"}` 立刻 EOF、exit 0，一条命令都不回（`cmd /c "helper < in > out"` 同样）；句柄换成可写则 `ReadLine` 抛 `UnauthorizedAccessException: 对路径的访问被拒绝`；
> - stdout 先 `writeFileSync` 再**另开只读句柄**读：读不到子进程写的字节；
> - 完全不重定向（继承控制台）：在 DSH 的伪控制台里 helper 拿不到控制台，起来就退。
>
> 细节写在 `scripts/lib/helper.cjs` 文件头，改这段前先读。
>
> ⚠️ **坑 5**：helper 协议里定位「已打开的文档」用 **`doc`** 字段；`path` 只在 `open`（要打开的文件）/ `save`（另存为）/ `versions`、`revert`（目标文件）里用。`{cmd:'open', app:'word', ...D}` 里 D 是 `{doc:P}` 的话，open 拿不到 path 会**静默新建一个空白文档**（返回 `how:"created"`，不报错）。
>
> ⚠️ **坑 6**：每个 helper 进程是独立状态。上一个脚本打开的文档，新脚本里看不到，会报 `NO_DOCUMENT`。要在**同一个进程**里跑完整流程。
>
> ⚠️ **坑 7**：测试前先关掉 WPS 里占着测试文件的文档，否则 `rm` 报 `Device or resource busy`。冒烟测试用独立临时目录，一般不受影响；残留下来的 `probe-*` 目录若被某个 WPS 窗口占着，先在 WPS 里关掉那个窗口。

### 3.4 打包安装

```bash
cd /e/DshData/dsh-office-wps
npm pack --ignore-scripts           # 必须加 --ignore-scripts，prepare 要 pnpm，PATH 里没有
cd /e/DshData/.dsh/profiles/desktop
/d/dsh/resources/runtime/cli/bin/dsh.cmd plugin --profile desktop add "E:/DshData/dsh-office-wps/andrewveraside-dsh-wps-sidebar-<ver>.tgz"
```

需要 `danger-full-access`（写 profile 的 `node_modules`）。

安装时 `@nanmicoder/dsh-agent-teams` 的版本不兼容警告是**既有问题**，跟本插件无关，忽略。

装完**重启 DSH**（关掉重开 `DeepSeek Harness.exe`）才生效。

### 3.5 发版

1. 改 `package.json` 的 `version`（语义化：加功能 minor、修 bug patch）
2. 更新 `README.md` 的工具表 / 功能说明
3. 更新本文件的「变更记录」
4. `git add -A && git commit && git push origin HEAD`

## 4. 环境事实（换机器要重新确认）

| 项 | 值 |
|---|---|
| 工作副本 | `E:\DshData\dsh-office-wps` |
| 仓库 | https://github.com/andrewveraside/dsh-wps-sidebar |
| DSH profile | `E:\DshData\.dsh\profiles\desktop` |
| helper 缓存 | `%LOCALAPPDATA%\dsh-office\helper-<hash>\office-helper.exe` |
| 编译器 | `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`（C# 5） |
| WPS | `E:\Apps\WPS Office\12.1.0.28488\office6` |
| 微软 Office | `C:\Program Files\Microsoft Office\root\Office16`（16.0.20527.20088，家庭学生版 2021 x64） |
| python（跑脚本） | `C:\Users\23171\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe` |
| bash | Git Bash；`pwsh.exe` 不存在，只有 `powershell.exe` |

> ⚠️ `csc.exe` 的 `/nologo` 等 MSVC 风格参数在 Git Bash 下会被当成路径。**必须写 .ps1 用 `powershell.exe -File` 跑**，且 `.ps1` 里别写中文（GBK 会截断），测试脚本一律纯 ASCII。

> ⚠️ `git push` 直连常报 `Recv failure: Connection was reset`（本机走 127.0.0.1:7897 的代理，git 没配）。临时用 `git -c http.proxy=http://127.0.0.1:7897 push origin main`，不要为此改全局 git 配置。

## 5. 坑：WPS 抢注了微软 Office 的 ProgID

本机 `Word.Application` / `Excel.Application` / `PowerPoint.Application` 三个 ProgID 的 `CLSID` 都指向 `wps.exe`（64 位和 32 位注册表都被改，`.ksobak` 备份里存的也是 WPS 路径，原始值没保留）。

**后果**：`suite: office` 和 `suite: auto` 在本机都会起 WPS。所以 `cordis.patch.yml` 里**硬钉 `suite: wps`**，走 WPS 专属的 `KWPS.Application` / `KET.Application` / `KWPP.Application`，不依赖被抢占的那三个。

想切回微软 Office 需要管理员改注册表（当前进程非管理员）。**不要**为了让 `suite: auto` 工作去改这个——配置已经绕过去了。

备份在 `C:\Users\23171\dsh-office-regbackup\`。

## 6. 改动记录（相对上游 copylee 1.2.0）

### 1.3.0 —— 改名 + Word / Excel 功能补齐

> 下列名字是从 `helper/OfficeHelper.cs` 的 `type == "..."` 分派里**实际提取**的，不是凭记忆写的。

**Word**
- 题注与图表目录：`caption`
- 交叉引用：`cross_reference`
- 域：`field` / `update_fields`
- 分节：`section` / `section_setup`
- 修订：`track_changes` / `revisions` / `compare_documents`
- 批注：`insert_comment` / `edit_comment` / `set_comment` / `delete_comment`
- 链接与书签：`add_hyperlink` / `edit_hyperlink` / `delete_hyperlink` / `add_bookmark`
- 水印：`watermark`（`WatermarkOp`，`helper` 3352 行）—— 代码在、基本功能可用，但 WPS 不支持 `Shape.Anchor`，页眉形状经 `SaveAs` 重开会掉到正文，复杂场景不可靠；README 未收录，用户确认基本不用，**不再投入**
- 其他：`insert_diagram` / `insert_flowchart` / `format_table` / `empty` / `fill_blanks` / `fill_gaps`

**Excel**
- 透视表：`refresh_pivot`
- 切片器：`add_slicer`
- 表格转区间：`table_to_range`（反向：`table`）
- 数字格式：`number_format`
- 其他：`conditional_format` / `validation` / `protect_sheet` / `unprotect_sheet` / `fill_formula` / `filter` / `sort` / `freeze` / `merge_cells` / `unmerge_cells` / `copy_range` / `remove_duplicates`

**PowerPoint**
- `theme` / `slide` / `add_slide` / `add_smartart` / `add_shape` / `set_shape` / `set_background` / `set_layout` / `move_slide` / `duplicate_slide` / `reuse_slide` / `delete_slide` / `set_notes`

**其他**
- `office_help` 可按应用列出全部 op

### 1.3.1 —— 修 bug

### 1.4.0 —— 版本存档（类似 git）

- `helper`：**每次 `office_save` 前**把改动前的文件复制到 `<文档目录>\.dsh-versions\<文档名>\<时间戳>_<原因>.<ext>`，每文档保留最近 50 份（`KeepVersions` 常量），目录设为隐藏
- `helper`：`VersionDir` / `VersionFiles` / `Snapshot` / `PruneVersions` / `ListVersions` / `RevertVersion` / `SnapshotNow`
- `helper`：`RevertVersion` 回退后**关文档再从磁盘重开**（否则窗口里还是旧内容，且下一命令报 `NO_DOCUMENT`）
- `helper`：版本号解析改用 `LastIndexOf('_')`（时间戳 `2026-10-28_20-39-48-752` 本身含下划线，用 `IndexOf` 会截成 `2026-10-28`）
- `lib`：`office_save` 新增 `no_version` / `version_note`
- `lib`：新增工具 `office_versions`（列版本）、`office_revert`（回退）

**版本语义（重要）**：版本存的是「这次保存**之前**文档的样子」，所以回退到某版本 = 回到那次改动之前 —— 和 git 一致。实测：`EDIT→SAVE1` 存下原始 956B；`EDIT→SAVE2` 存下第一版 11285B；`REVERT latest` 回到 11285B，读出来是 `SECOND VERSION CONTENT`。

## 7. 已知限制

- **`suite: wps` 是硬编码的**，换到没装 WPS 的机器需要改成 `auto` 或 `office`
- WPS 公式在演示里是**矢量图叠加**，不是可编辑公式；移动/缩放文字形状后要重写文字才会重新定位
- WPS 文字窗口**不跟随滚动**
- Excel / PPT 的外部修改**不能用 Ctrl+Z 撤回**（Word 可以）
- 不能写原生 Excel 公式（WPS 限制）
- 兼容模式（Word 97-2003）文档不能建原生公式
- PowerPoint 表格单元格和 SmartArt 节点里放不了排好的公式
- **侧栏不能预览文档**：插件只注册了设置页，没有文件预览组件。侧栏白屏是 DSH 通用预览器找不到 docx 渲染器，不是 bug
- 版本存档**每次保存都存一份**，即使内容没变（用户选的「每次保存前自动存档」策略，稳妥优先）

## 8. 接手 checklist

新工作区第一次打开这个项目时：

1. `git log --oneline -5` 看最近改了什么
2. 读本文件 §3 确认开发循环还能跑（编译 → 测试 → 打包 → 安装）
3. 确认 `scripts/build-helper.ps1` 编译的是**工作副本**（`E:\DshData\dsh-office-wps\helper\OfficeHelper.cs`），不是 node_modules；跑 `node scripts/office-smoke.mjs` 确认 14/14 绿
4. 确认 `package.json` 版本号 / `git status` 干净
5. 想加功能前先读 §2 的表，找到对应文件
