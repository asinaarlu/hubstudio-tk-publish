# hubstudio-tk-publish

把 Hubstudio 客户端「TK运营助手 · 批量发视频」的手工流程（新建任务 → 选账号环境 → 传视频 → 填文案/小黄车/首评/发布时间 → 发布 → 盯失败重试 → 回写台账）改成一个**完全后台、无界面**的命令行脚本：只调 Hubstudio 本机开放接口，不点鼠标、不抢焦点、不截图。

> 这是个人运营用的自动化脚本，不是 Hubstudio 或 TikTok 的官方工具。用之前先确认你的账号和用法符合平台规则，风险自负。仓库里所有账号环境名、serialNumber、商品 ID、首评文案都换成了占位符，真实配置请放本地、不要提交。

## 前置条件

- Windows（脚本用了 `kernel32` 判进程存活，未在 Linux/macOS 验证）
- 本机装好并运行 Hubstudio（实测版本 3.61.2），开放接口在 `http://127.0.0.1:6873`
- Python 3.8+，依赖只有 `openpyxl`（读写台账 xlsx）：`pip install openpyxl`
- 云手机处于**已开机**状态。脚本只读状态，不会替你开机（见下方红线 P1）
- 一份台账 xlsx + 每条视频对应的文案 txt

## 快速开始

```bash
# 1. 配置：复制示例，把 <...> 占位换成你自己的
cp tk_publish.config.example.json tk_publish.config.json
#   必须改的 5 项：ledger.path、video_source.root、outputs_dir、
#                 phones.list[].serial、products.list[].productId

# 2. 看一眼环境和台账读取是否正常
python -X utf8 tk_publish.py phones
python -X utf8 tk_publish.py today

# 3. 出计划（只算不建，产物里带一次性确认令牌）
python -X utf8 tk_publish.py plan
python -X utf8 tk_publish.py verify --token <plan 输出的令牌>   # 核对 + 自检

# 4. 确认无误后才真正排期发（到点才提交，脚本常驻后台等）
powershell -ExecutionPolicy Bypass -File launch_schedule.ps1 -Token <令牌>

# 5. 另起一个自检守护：每条发完 15 分钟查一次、下一条发前 15 分钟再查一次
powershell -ExecutionPolicy Bypass -File launch_checkd.ps1 -Token <同一个令牌>
```

`plan` 不落令牌就没法提交：只有**落盘的**计划文件里的 `confirm_token` 才有效，`plan --no-out` 是纯只读预演、刻意不生成令牌。这样"看一眼"和"真的发"永远是两个动作，中间必须过一次 `verify`。

## 排期规则

- 起始时间 = 运行时刻的真实时间，不传 `--start-time` 就从下一整分钟开始
- `auto_spacing=true` 时间隔自动摊：`间隔 = max(round_spacing_minutes, (22:00 - 起始) // (单环境条数 - 1))`
- 两条硬下限：同一台云手机相邻两条 ≥ `min_spacing_minutes`（默认 30）；轮次间隔默认 ≥ `round_spacing_minutes`（默认 60，留出自检重发要占的两个 30 分钟空档）
- 当天最后一条不越过 `day_last_deadline`（默认 22:00），越界是 BLOCK 不是提醒
- 一个任务里**每台云手机只放一条**：同任务出现重复 `mobileId` 会被接口直接打回（`E010008`）；确实要一条一任务时显式加 `--one-per-task`
- 按权重分配（`publish.alloc.enabled=true`）时用最大余数法拆，并先统计今天各号已发几条，超 `max_day_per_mobile` 就往权重次高的号挪

## 子命令

| 命令 | 作用 | 对外有没有写 |
|---|---|---|
| `today` | 今天已发/待发一览 | 只读 |
| `phones` | 列云手机状态 | 只读 |
| `ledger` | 台账读取与体检 | 只读 |
| `plan` | 算计划、出校验报告与令牌 | 只读接口 + 本地写 |
| `verify` | 按令牌重跑核对与自检 | 只读 |
| `submit` | 立刻提交（一次性） | **写：建任务** |
| `schedule` | 常驻守护，到点批量提交 | **写：建任务** |
| `check` | 查执行明细、判成功/失败/结果不明 | **写：明确失败当场重发**（只看不动加 `--no-retry`） |
| `checkd` | 自检 + 失败重发守护 | **写：可建重发任务** |
| `retry` | 一行命令重发某个任务的失败条目，等价在任务详情页点「重试」 | **写：可建重发任务** |
| `watch` / `status` | 盯任务进度 | 只读 |
| `watchdog` | 巡检守护进程与状态是否打架 | 只读 |
| `cancel` | 取消任务（仅限本脚本 state 里记下的） | **写：取消** |
| `pack` / `paste` / `retime` | 出客户端批量粘贴包、写剪贴板、改时间 | 本地文件 |

常用参数：`--config` `--token` `--only` `--exclude` `--task` `--start-time` `--spacing` `--schedule` `--one-per-task` `--max-day` `--check-after` `--check-before` `--check-interval` `--max-retry` `--no-retry` `--retry-wait-gap` `--retry-now` `--retry-ambiguous` `--force-check` `--force` `--no-out` `--skip-url-check` `--no-upload` `--write-back` `--in-place` `--dry-run`。

退出码只有一个非零值：**成功 0，其余一律 2**（内部统一 `die(msg, code=2)`），适合直接拿来做批处理判断。

## 常见错误码

| 码 | 含义 | 处理 |
|---|---|---|
| `E010008` | 同一任务里 `mobileId` 重复，或 `queueWhenBusy` 非布尔 | 拆任务 / 检查参数 |
| `E114208` | 任务名重复 | 改 `task_name_prefix` 或让脚本自动加后缀 |
| `E010006` | 接口不支持该操作（比如想改发布时间、想重试） | 开放接口给不了，只能重建任务或回客户端点 |
| `E114269` | 云手机未开机 | 开机是人工动作，脚本不代劳 |
| `E010205` | 请求太密，被限流 | 调大 `hubstudio.min_interval_sec` |
| `4041` / `E114204` | 结果不明：可能已发也可能没发 | **不自动重发**，去账号页确认后再决定 |
| `26` | 套餐权益不足 | 续费或减量 |

## 安全红线

这几条是脚本硬编码的判断，不是文档建议：

1. **P1** 不调用任何云手机开机/关机/重启/唤起接口，状态只读并告警。
2. **P2** 不修改、暂停、删除用户已有任务；`cancel` 只允许取消本脚本 state 里记下的 `taskId`。
3. **P3** 不往素材目录写任何文件；台账回写默认只出副本，改原表要显式 `--write-back`。
4. **P4** 不生成截图/抽帧/音频等中间文件。
5. **P5** 不做前台界面自动化：不抢焦点、不点击、不粘贴、不敲键盘。
6. **P6** 核对 + 自检没过不发：先确认视频与商品信息准确、完整、一致，再完整自检一次，发现问题先修再重新自检。

## 已知边界（不是 bug，是接口给不了）

- `video-task/create` 不解析任何时间字段，试过 21 个候选字段名全部无效。"定时发布"靠的是脚本自己等到点再提交，或者用客户端向导。
- **没有重试接口**。客户端那两个按钮走的是内部路由（顶部"批量重试"= `task-main/retry {taskId}`，行内"重试"= `task-detail/retry {内部行 id}`），都不在本地开放接口的 11 条 tk-automation 白名单里：10 种路径写法加假 taskId 直连一律 `E010006`（对照组用白名单内的 `task/cancel` 回 `E010008`，证明不是拼错）。行内重试还缺参数——`task/execution/page` 不回那个内部行 id。所以所谓自动重发，是脚本拿原计划载荷**重建一个任务**：文案、商品 ID、首评、已上传好的视频地址逐字复用。
- 界面点重试是**原地重跑**：`executionId` 不变，只把 `startedAt` 刷新。脚本比每一圈的执行快照认得这个指纹（`externally_rerun`），发现"这条已经被点过了"就把自己的重发撤掉，不会和界面撞成同账号双发。
- 自检守护的退出条件不看 `4041/E114204` 这类"结果不明"，日志说"全部落定"时不代表结果不明的那几条已经查清，那部分必须人工确认。

## 目录结构

```
tk_publish.py                      主脚本（配置缺失时兜底用内置默认值）
tk_publish.config.example.json     配置示例，复制成 tk_publish.config.json 用
launch_schedule.ps1                排期守护的分离式启动器（无窗口、不抢焦点）
launch_checkd.ps1                  自检守护的分离式启动器
batch_watch.py                     只读盯梢：把今天本脚本建的任务汇总成一行
_set_clip.py                       把粘贴包按 Excel 剪贴板约定写进剪贴板（向导路线用）
```

跑起来后 `outputs_dir` 里会出现：计划 json、`*_state.json` 状态文件、`.lock` 单实例锁、`*.log` 日志。这些都属于运行时产物，别提交。
