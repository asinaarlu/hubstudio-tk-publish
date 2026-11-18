#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Hubstudio「TK运营助手 · 批量发视频」后台自动上架脚本
====================================================

把录屏里的 7 步手工操作（导出运行结果 → 找失败/未发布 → 新建任务 → 选账号环境
→ 传视频 → 填文案/小黄车/评论/定位/AI-Label/发布时间/忙时等待 → 发布 → 监控重试
→ 回写台账）改成走 Hubstudio 本地 API 的无界面脚本。

已验证的接口（本机 Hubstudio 3.61.2，http://127.0.0.1:6873）
    POST /api/v1/cloud-mobile/mobile-page                                  列云手机
    POST /api/v1/cloud-mobile/tk-automation/video-task/create              建发视频任务
    POST /api/v1/cloud-mobile/tk-automation/task/page                       任务列表
    POST /api/v1/cloud-mobile/tk-automation/task/detail                     任务详情
    POST /api/v1/cloud-mobile/tk-automation/task/execution/page             运行明细(失败原因/视频链接)
    POST /api/v1/cloud-mobile/tk-automation/task/cancel                     取消任务(仅限本脚本创建的)

video-task/create 请求体（2026-11-16 用「必然被校验打回」的请求逐字段反查得到，未建成功任务）：
    {
      "taskName": "新建任务NNN",
      "queueWhenBusy": true,                        # 顶层布尔，非布尔报 E010008
      "items": [{
          "mobileId":     100000001,                # 必填整数；缺失/非法 -> E010008
          "videoUrl":     "https://.../a.mp4",      # 必填非空串
          "videoName":    "a.mp4",
          "caption":      "视频文案",                # 限 1000 字符（不是 videoText！）
          "productId":    "1740000000000000001",     # ^[0-9]{6,50}$ 或空串
          "productTitle": "Sample Product Title",    # 限 30 字符，禁 %、反斜杠、+、竖线、双引号、「」
          "comment":      "首评置顶",                # 限 200 字符
          "enableAIGC":   false,                     # 布尔
          "getLeads":     false,
          "inQueue":      true,
          "videoSize"/"videoWidth"/"videoHeight"/"videoDuration": 客户端 config 用，开放接口不校验
      }]
    }
    判据（逐条实测）：items 缺 mobileId/videoUrl 报错；caption 超 1000 报错；videoText 传超长不报错
    （= 不认）；comment 超 200 报错；productTitle 超 30 或含 % 报错；productId 非 6~50 位数字报错；
    queueWhenBusy 非布尔报错；enableAIGC 非布尔报错。

关于发布时间（重要，别再看错）：
    app.asar 里客户端内部发的是 {taskName, details:[{mobileId, triggerTime, config:{videoText…}}]}，
    那是渲染层 -> 内网 v2 接口的结构，和本机开放接口不是同一套；用 details 提交开放接口会被打成
    「items 至少提供一个条目」（2026-11-16 实发被拒的正是这个）。
    开放接口这边试过 21 个候选时间字段名（triggerTime/executeAt/scheduledAt/…），非法值也不报错，
    说明校验层不解析时间；实测一个只带 mobileId+videoUrl 的任务创建后 40 秒就 RUNNING，
    所以「定时」能否生效必须以创建后 execution/page 的 scheduledAt 为准（本脚本仍在 item 里带
    triggerTime 供报表核对）。要真正定时发布有两条路：客户端向导，或本脚本的 schedule 子命令
    ——schedule 由进程自己等到点再调 video-task/create，发出去的时刻就等于 plan 的时刻，
    全程只走本地接口，不碰客户端界面（2026-11-17 定下的做法：一律后台，不动前台窗口）。

客户端里的其它硬约束（同一次核对读到的常量）
    MAX_VIDEO_LENGTH = 15            单账号环境单任务上限
    TK_DEFAULT_VIDEO_DIFF_TIME = 10  同一台手机默认间隔（分钟），是下限不是推荐值
    TK_DEFAULT_MOBILE_DIFF_TIME = 5  不同手机之间默认间隔
    isBeforeTime()                   渲染层禁止 triggerTime 早于当前时间（UTC+8）
    文案限 1000 字符，productTitle 限 30 字符，comment 限 200 字符

已实测确认做不到的两件事（其中第 1 条已在 2026-11-14 解决，见下）
    1) 本地 mp4 -> videoUrl：/get-upload-url 与 /upload-file-local 要求云手机开机（P1 禁止），
       但 /api/v1/cloud-mobile/get-s3-url {channelId:1,scene:"taskVideo",bucketCode:"public",
       fileNames:[...]} 返回预签名 PUT 地址，直接 PUT 上去、去掉 query 就是公网 URL。
       这条路由不碰云手机电源，脚本已按此实现（video_source.mode = "s3"）。
    2) 运行明细里没有视频名：execution/page 只有 mobileId/status/errorCode/publishedUrl，
       所以「哪条对应台账哪一行」无法自动判定，台账回写只能靠 --mark 或人工核对。

安全边界（与《Hubstudio批量发视频_工作流程规范_v1》一致）
    P1 不调用开机/关机/重启等云手机控制接口
    P2 不修改/暂停/删除已有任务；cancel 只允许操作本脚本 state 里记录的 taskId
    P3 默认不向素材目录写任何文件；台账回写需显式 --write-back
    P4 不产生截图/抽帧等中间图片文件
    P6 默认只出计划（plan），必须拿 plan 给出的 token 才能 submit

用法
    python tk_publish.py plan
    python tk_publish.py plan --retry-export "E:\发视频新建任务202（hubstudio）.xlsx"
    python tk_publish.py submit --token <plan 打印的 token>
    python tk_publish.py pack --only 蘑菇灯 --seq 22 --spacing 60     # 出向导粘贴包（不提交）
    python tk_publish.py retime --file 蘑菇灯_粘贴包_连续60_第1批.xlsx \\
            --start-time "11-18 20:00" --spacing 60 --mobile-spacing 5   # 批量修改发布时间
    python tk_publish.py status
    python tk_publish.py watch --interval 120
    python tk_publish.py phones
    python tk_publish.py ledger --write-back

排期（--spacing / --mobile-spacing / --start-time / --schedule）
    这两个间隔就是客户端向导第 4 步「批量设置发布时间」对话框里的两个字段，边界也照抄源码：
      spacing        单手机多视频间隔，向导允许 10~1440 分钟，默认 60（源码默认 10）
      mobile-spacing 多手机间隔，    向导允许 5~1440 分钟，默认 5
    start-time     排期起点 = 对话框的「起始时间」；支持 20:00 / 11-17 20:00 / 2026-11-17 20:00，
                   只给时刻且已过则顺延到明天。wizard 模式留空则取下一整分钟。
    schedule       wizard         = 逐字复刻客户端算法（推荐，见下）
                   caption_window = 从起点一路 +spacing（22 条 * 60 分钟 = 跨 21 小时，会排进凌晨）
                   window_daily   = 只在窗口内排（20:00-23:00 一天 4 条），排满顺延次日（22 条 = 6 天）
                   fixed          = 用 publish.fixed_time；immediate = 建完即发

批量修改发布时间（wizard 模式与 retime 共用同一条公式）
    客户端源码 app.asar module 13504 里那个对话框的 onOk 只有一行：
        triggerTime = 起始时间 + 组内视频序号 * 单手机多视频间隔 + 手机分组序号 * 多手机间隔
    外层 n 是账号环境在表格里首次出现的顺序，内层 l 是该环境下视频的顺序，单位分钟，
    对话框顶部还写着「设置后将覆盖单独设定的时间，时区为 UTC+8 北京时间，可能略有延迟」。
    2026-11-17 录屏《屏幕录制 2026-11-17 084833.mp4》实测：1 台手机、起始 08:48、间隔填 50
    -> 三条变成 08:48 / 09:38 / 10:28，与本公式一致。
    所以脚本算出来的 发布时间 列，和你在向导里点一次「批量设置」敲同样三个数得到的是同一张表；
    pack/plan 会顺手打印这三个数该填多少。
    retime 用来事后改时间：只改 xlsx 的 发布时间 列（「核对」表同步），内容列一律不动，
    默认另存 <原名>_改时.xlsx，--in-place 才覆盖，且只允许覆盖 outputs_dir 里的文件（P3）。
    改已建任务的时间用的是 task-main/batch-update-trigger-time，那条路由不在本地开放 API
    白名单里（白名单只有 video-task/create、task/page、task/detail、task/execution/page、
    task/cancel 等），只能在客户端点，脚本碰不到。
    注意这套时间只对客户端向导有用：开放接口 video-task/create 不解析时间字段（见上）。
    不想碰界面的话用 schedule 子命令——它按这些算好的时刻到点再建任务，效果就是定时发布。
"""

import argparse
import csv
import glob
import hashlib
import io
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# ----------------------------------------------------------------------------
# 常量：来自官方「批量发视频」文档与客户端源码 (module 6145) 的硬阈值
# ----------------------------------------------------------------------------
MAX_VIDEO_MB = 200                # 单视频 ≤200MB
MAX_VIDEOS_PER_MOBILE = 15        # 源码 MAX_VIDEO_LENGTH=15：单台云手机单任务≤15条
VIDEO_EXTS = (".mp4", ".mov")
COVER_EXTS = (".png", ".jpg", ".jpeg")
COVER_MAX_MB = 5
PRODUCT_ID_RE = re.compile(r"^\d{6,50}$")     # 客户端 rules: ^\d{6,50}$
MAX_PRODUCT_TITLE = 30                         # 小黄车商品标题 ≤30 字符
TITLE_BAD_CHARS = set("%\\+|\"「」")             # 客户端表单对商品标题禁的字符（原样照抄，别改）
MAX_COMMENT = 200                              # 首评 ≤200 字符
MAX_CAPTION_DEFAULT = 1000                     # 客户端 tkTextRules() 默认限 1000 字符（不是 TikTok 的 2200）
DEFAULT_VIDEO_SPACING_MIN = 10                 # 源码 TK_DEFAULT_VIDEO_DIFF_TIME=10
DEFAULT_MOBILE_SPACING_MIN = 5                 # 源码 TK_DEFAULT_MOBILE_DIFF_TIME=5
MIN_SPACING_FLOOR = 30                         # 2026-11-17 你定的「每条间隔不低于 30 分钟」
ROUND_SPACING_DEFAULT = 60                     # 加了自检之后的排期下限：一轮 + 一次重发要占两个 30 分钟空档
CATCHUP_DEFAULT_MIN = 10                       # schedule 的迟到宽限，watchdog 判「过点了还没动」用同一个数（别两处各写一个 10）
DEFAULT_DAY_DEADLINE = "22:00"                 # 2026-11-17 你定的「当天最晚一条不越过 22:00」
MIN_VIDEO_SPACING_MIN = 10                     # 向导「单手机多视频间隔」InputNumber min=10
MIN_MOBILE_SPACING_MIN = 5                     # 向导「多手机间隔」InputNumber min=5
MAX_SPACING_MIN = 1440                         # 两个间隔字段的共同上限（源码 max:1440）
RATE_LIMIT_CODE = "E010205"
NO_VIP_CODE = "26"
TIME_FMT = "%Y-%m-%d %H:%M:%S"
TRIGGER_FMT = "%Y-%m-%d %H:%M"      # 客户端 TKDateFormat = "YYYY-MM-DD HH:mm"（无秒）


def video_meta(path):
    """接口 config 里带 videoSize/Width/Height/Duration，用 OpenCV 就地读，不产生任何中间文件（P4）。"""
    meta = {"videoSize": os.path.getsize(path)}
    try:
        import cv2
        cap = cv2.VideoCapture(path)
        try:
            w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
            h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
            fps = cap.get(cv2.CAP_PROP_FPS) or 0
            n = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
            if w:
                meta["videoWidth"] = w
            if h:
                meta["videoHeight"] = h
            if fps and n:
                meta["videoDuration"] = round(float(n) / float(fps), 2)
        finally:
            cap.release()
    except Exception:
        pass
    return meta



def log(msg):
    print(msg, flush=True)


def die(msg, code=2):
    log("[停止] " + msg)
    sys.exit(code)


# ----------------------------------------------------------------------------
# 本地 API 客户端
# ----------------------------------------------------------------------------
class ApiError(RuntimeError):
    def __init__(self, res):
        self.res = res
        RuntimeError.__init__(self, "%s %s" % (res.get("code"), res.get("msg")))


class Hubstudio:
    def __init__(self, host="127.0.0.1", port=6873, min_interval=0.8, timeout=120):
        self.base = "http://%s:%d" % (host, port)
        self.min_interval = min_interval
        self.timeout = timeout
        self._last = 0.0

    def post(self, path, payload, retries=4):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        for attempt in range(retries):
            wait = self.min_interval - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            req = urllib.request.Request(
                self.base + path, data=body,
                headers={"Content-Type": "application/json"}, method="POST")
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as r:
                    res = json.loads(r.read().decode("utf-8", "ignore"))
            except urllib.error.HTTPError as e:
                res = safe_json(e.read()) or {"code": "HTTP%s" % e.code, "msg": str(e)}
            except Exception as e:                      # 连接失败等
                res = {"code": "LOCAL", "msg": "%r" % e}
            self._last = time.time()
            if str(res.get("code")) == RATE_LIMIT_CODE and attempt < retries - 1:
                time.sleep(2 + 3 * attempt)
                continue
            return res

    # --- 只读 ---
    def mobiles(self):
        res = self.post("/api/v1/cloud-mobile/mobile-page", {"current": 1, "size": 200})
        if str(res.get("code")) != "0":
            raise ApiError(res)
        return (res.get("data") or {}).get("list") or []

    def tasks(self, size=50):
        res = self.post("/api/v1/cloud-mobile/tk-automation/task/page",
                        {"pageNum": 1, "pageSize": size})
        if str(res.get("code")) != "0":
            raise ApiError(res)
        return (res.get("data") or {}).get("list") or []

    def task_detail(self, task_id):
        res = self.post("/api/v1/cloud-mobile/tk-automation/task/detail", {"taskId": task_id})
        if str(res.get("code")) != "0":
            raise ApiError(res)
        return res.get("data") or {}

    def executions(self, task_id, size=100):
        out, page = [], 1
        while True:
            res = self.post("/api/v1/cloud-mobile/tk-automation/task/execution/page",
                            {"taskId": task_id, "pageNum": page, "pageSize": size})
            if str(res.get("code")) != "0":
                raise ApiError(res)
            d = res.get("data") or {}
            out += d.get("list") or []
            if len(out) >= int(d.get("total") or 0) or not (d.get("list") or []):
                break
            page += 1
        return out

    # --- 写入 ---
    def create_video_task(self, payload):
        res = self.post("/api/v1/cloud-mobile/tk-automation/video-task/create", payload)
        if str(res.get("code")) != "0":
            raise ApiError(res)
        return res.get("data") or {}

    def cancel_task(self, task_id):
        res = self.post("/api/v1/cloud-mobile/tk-automation/task/cancel", {"taskId": task_id})
        if str(res.get("code")) != "0":
            raise ApiError(res)
        return res.get("data")

    # --- 素材上传（PUT 预签名，不需要云手机开机）---
    def s3_put_urls(self, file_names, channel_id=1, scene="taskVideo", bucket="public"):
        """和客户端 module 22492 完全一致的调用：{channelId, scene, bucketCode, fileNames[]}。
        返回 [(fileName, 预签名PUT地址, cdnHost)]；缺 channelId 会被 "Channel cannot be empty" 挡掉。"""
        res = self.post("/api/v1/cloud-mobile/get-s3-url",
                        {"channelId": int(channel_id), "scene": scene, "bucketCode": bucket,
                         "fileNames": list(file_names)})
        if str(res.get("code")) != "0":
            raise ApiError(res)
        out = []
        for d in (res.get("data") or []):
            out.append((d.get("fileName") or "", d.get("url") or "", d.get("cdnHost")))
        return out


def s3_public_url(signed_url, cdn_host=None):
    """预签名地址去掉 query 就是公网地址；有 cdnHost 时按客户端做法替换域名。"""
    if not signed_url:
        return ""
    base = signed_url.split("?", 1)[0]
    if cdn_host:
        parts = urllib.parse.urlsplit(base)
        base = urllib.parse.urlunsplit((parts.scheme, cdn_host, parts.path, "", ""))
    return base


def http_put_file(signed_url, local_path, timeout=900, chunk=1024 * 256):
    """流式 PUT 整个文件，返回 HTTP 状态码（200 才算成功）。"""
    size = os.path.getsize(local_path)

    def reader():
        with io.open(local_path, "rb") as f:
            while True:
                b = f.read(chunk)
                if not b:
                    break
                yield b

    req = urllib.request.Request(signed_url, data=reader(), method="PUT",
                                 headers={"Content-Length": str(size),
                                          "Content-Type": "multipart/form-data;charset=UTF-8"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status
    except urllib.error.HTTPError as e:
        return e.code
    except Exception as e:
        log("[上传] PUT 异常：%r" % (e,))
        return 0


def safe_json(raw):
    try:
        return json.loads((raw or b"").decode("utf-8", "ignore"))
    except Exception:
        return None


# ----------------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "hubstudio": {"host": "127.0.0.1", "port": 6873, "min_interval_sec": 0.8},
    "outputs_dir": "./outputs",
    "ledger": {
        "path": "./data/发布跟踪表.xlsx",
        "sheet": "发布跟踪",
        "columns": {"seq": "序号", "name": "视频名称", "video": "视频文件路径",
                    "caption": "发布文案的路径", "status": "发布状态"},
        "unpublished_values": ["未发布", ""],
        "published_value": "已发布",
    },
    "products": [
        {"match": ["示例产品A"], "productId": "1740000000000000001", "productTitle": "",
         "title_by_video": {"示例产品A-01": "Sample Product Title One",
                            "示例产品A-02": "Sample Product Title Two"}},
        {"match": ["示例产品B"], "productId": "1740000000000000002", "productTitle": ""},
    ],
    "comment_by_video": {
        "示例产品A-01": "（示例首评文案，实际文案写在 tk_publish.config.json 里）",
    },
    "phones": [
        {"name": "Hubstudio账号环境名", "serial": 100000001},
    ],
    "publish": {
        "task_name_prefix": "新建任务",
        "schedule_mode": "caption_window",   # immediate | fixed | caption_window
        "fixed_time": "",                    # "2026-11-15 20:30"
        "spacing_minutes": 60,               # 固定间隔（auto_spacing=false 或显式 --spacing 时用）
        "min_spacing_minutes": 30,           # 硬规则下限：同台相邻两条不得更短
        "round_spacing_minutes": 60,         # 排期下限：一轮 + 一次自检重发要占两个 30 分钟空档
        "queue_when_busy": True,             # 忙时等待
        "enable_aigc": False,                # AI-Label
        "get_leads": False,
        "location": "",
        "chunk_size": 15,                    # 单任务提交条数上限（≤15/台）
    },
    "video_source": {"mode": "s3",          # s3 | url_prefix | map | unset
                     "channel_id": 1, "scene": "taskVideo", "bucket_code": "public",
                     "url_prefix": "", "root": "D:/videos", "map_file": ""},
    "caption": {"include_title": True,
                "include_hashtags": True, "use_pinned_comment": True,
                "max_caption_chars": MAX_CAPTION_DEFAULT},
}


def deep_merge(base, extra):
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            deep_merge(base[k], v)
        elif isinstance(v, list) and isinstance(base.get(k), list) and all(isinstance(x, dict) for x in v):
            ids = {x.get("match") or x.get("name") for x in base[k]}
            base[k] += [x for x in v if (x.get("match") or x.get("name")) not in ids]
        else:
            base[k] = v
    return base


LIST_KEYS = ("products", "phones")


def unwrap_commented_lists(cfg):
    """配置里 products/phones 可以写成 {"_说明": "...", "list": [...]}，统一摊平成数组。"""
    for k in LIST_KEYS:
        v = cfg.get(k)
        if isinstance(v, dict):
            cfg[k] = list(v.get("list") or [])
    return cfg


def load_config(path):
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if path and os.path.exists(path):
        with io.open(path, encoding="utf-8") as f:
            deep_merge(cfg, json.load(f))
    return unwrap_commented_lists(cfg)


# ----------------------------------------------------------------------------
# 文案解析
# ----------------------------------------------------------------------------
SECTION_RE = re.compile(r"^\s*【([一二三四五六七八九十])】\s*([^\n]*)$", re.M)


def parse_caption_file(path):
    """返回 {title, body, hashtags, pinned_comment, window, spec}；缺段返回空串。"""
    out = {"title": "", "body": "", "hashtags": "", "pinned_comment": "", "window": "", "spec": ""}
    if not path or not os.path.exists(path):
        return out, False
    text = io.open(path, encoding="utf-8", errors="replace").read()
    hits = list(SECTION_RE.finditer(text))
    for i, m in enumerate(hits):
        name = m.group(2)
        end = hits[i + 1].start() if i + 1 < len(hits) else len(text)
        content = text[m.end():end].strip("=\n-— \t")
        content = re.sub(r"^[=\s]*", "", content).strip()
        if "标题" in name and "Title" in name:
            out["title"] = content.splitlines()[0].strip() if content else ""
        elif "正文" in name or "Caption" in name:
            out["body"] = content
        elif "话题" in name or "Hashtag" in name:
            out["hashtags"] = content
        elif "发布建议" in name:
            for line in content.splitlines():
                if "首评" in line:
                    out["pinned_comment"] = re.sub(r"^.*首评[^:：]*[:：]\s*", "", line).strip()
                if "发布时间" in line:
                    mm = re.search(r"(\d{1,2}:\d{2})\s*[-~–至]\s*(\d{1,2}:\d{2})", line)
                    out["window"] = mm.group(0) if mm else line.strip()
                    out["window_start"] = mm.group(1) if mm else ""
        elif "视频规格" in name or "1080" in name:
            out["spec"] = content
    if not out["body"] and hits:                       # 没识别到【二】就退回首行正文
        out["body"] = "\n".join(text.splitlines()[2:6]).strip()
    return out, True


def pick_window_start(window_info):
    txt = window_info.get("window_start") or ""
    m = re.search(r"(\d{1,2}):(\d{2})", txt)
    if not m:
        m = re.search(r"(\d{1,2}):(\d{2})", window_info.get("window", ""))
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def pick_window_range(window_info):
    """文案【四】里的 20:00-23:00 解析成 (起时, 起分, 止时, 止分)；没有窗口就 (None,None,None,None)。"""
    txt = (window_info or {}).get("window") or ""
    m = re.search(r"(\d{1,2}):(\d{2})\s*[-~–至]\s*(\d{1,2}):(\d{2})", txt)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
    sh, sm = pick_window_start(window_info or {})
    return (sh, sm, None, None) if sh is not None else (None, None, None, None)


# ----------------------------------------------------------------------------
# 台账与导出表
# ----------------------------------------------------------------------------
def read_ledger(cfg):
    import openpyxl
    lc = cfg["ledger"]
    path = lc["path"]
    if not os.path.exists(path):
        die("台账不存在：%s" % path)
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = lc.get("sheet") or wb.sheetnames[0]
    if sheet not in wb.sheetnames:
        die("台账缺少工作表：%s（现有：%s）" % (sheet, ",".join(wb.sheetnames)))
    ws = wb[sheet]
    header = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    col = {}
    for key, title in lc["columns"].items():
        if title not in header:
            die("台账缺少列「%s」（表头：%s）" % (title, " | ".join(header)))
        col[key] = header.index(title)
    rows = []
    for r, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if values is None or all(v in (None, "") for v in values):
            continue
        def g(k):
            i = col[k]
            return "" if i >= len(values) or values[i] is None else str(values[i]).strip()
        rows.append({"row": r, "seq": g("seq"), "name": g("name"), "video": g("video"),
                     "caption_path": g("caption"), "status": g("status")})
    return rows


def read_export(cfg, path):
    """读 Hubstudio 导出的运行结果表，返回失败/取消的执行行。"""
    import openpyxl
    if not os.path.exists(path):
        die("导出表不存在：%s" % path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    for need in ("视频名称", "运行结果"):
        if need not in idx:
            die("导出表缺少列「%s」（表头：%s）" % (need, " | ".join(header)))
    out = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or all(v in (None, "") for v in values):
            continue
        res = str(values[idx["运行结果"]] or "")
        if res.startswith("失败") or "失败" in res:
            reason = res.split("：", 1)[1] if "：" in res else res
            out.append({"name": str(values[idx["视频名称"]] or "").strip(),
                        "reason": reason.strip(),
                        "env": str(values[idx["账号环境"]] or "").strip() if "账号环境" in idx else "",
                        "plan_time": str(values[idx["计划时间"]] or "").strip() if "计划时间" in idx else ""})
    return out


def read_export_all(cfg, path):
    """读导出表全部行 -> [(视频名, 运行结果原文)]。用于反查「其实已经发过」的视频。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    if "视频名称" not in idx or "运行结果" not in idx:
        return []
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or all(v in (None, "") for v in values):
            continue
        nm = str(values[idx["视频名称"]] or "").strip()
        rs = str(values[idx["运行结果"]] or "").strip()
        if nm:
            rows.append((nm, rs))
    return rows


def published_names_from_exports(cfg, extra_paths=None):
    """扫描 outputs/exports/*.xlsx（+ 显式给的路径），返回运行结果以「成功」开头的视频名集合。

    台账「发布状态」是人填的，会滞后；导出表是 Hubstudio 的事实。两者冲突时以导出表为准，
    否则会把已经发过的视频再发一遍（重复发布在 TK 侧是会掉权重的）。
    """
    import glob as _glob
    cand = list(_glob.glob(os.path.join(cfg["outputs_dir"], "exports", "*.xlsx")))
    cand += list(extra_paths or [])
    done, seen = set(), set()
    for p in cand:
        if not os.path.exists(p) or os.path.normpath(p).lower() in seen:
            continue
        seen.add(os.path.normpath(p).lower())
        try:
            for nm, rs in read_export_all(cfg, p):
                if rs.startswith("成功"):
                    done.add(nm.lower())
        except Exception as e:
            log("[警告] 导出表读失败，忽略：%s（%s）" % (p, e))
    return done, sorted(seen)


# ----------------------------------------------------------------------------
# 解析器：视频 URL
# ----------------------------------------------------------------------------
def resolve_video_url(cfg, local_path, api=None, uploader=None):
    vs = cfg["video_source"]
    mode = vs.get("mode", "unset")
    if mode == "unset":
        return "", "未配置 video_source.mode（可选 s3 / url_prefix / map）"
    if not local_path or not os.path.exists(local_path):
        return "", "本地视频文件不存在"
    if mode == "s3":
        if uploader is None:
            return "", "s3 模式本次没有上传（--no-upload），加上就能取到 URL"
        return uploader(local_path)
    if mode == "url_prefix":
        prefix = (vs.get("url_prefix") or "").rstrip("/")
        root = (vs.get("root") or "").rstrip("/")
        if not prefix:
            return "", "video_source.url_prefix 为空"
        rel = local_path.replace("\\", "/")
        if root and rel.lower().startswith(root.lower()):
            rel = rel[len(root):].lstrip("/")
        else:
            rel = os.path.basename(local_path)
        url = prefix + "/" + urllib.parse.quote(rel.replace("/", "/"), safe="/-_.~")
        return url, ""
    if mode == "map":
        mf = vs.get("map_file") or ""
        if not mf or not os.path.exists(mf):
            return "", "video_source.map_file 不存在"
        table = load_url_map(mf)
        key = os.path.basename(local_path).lower()
        return table.get(key) or table.get(local_path.replace("\\", "/").lower()) or "", \
               ("" if table.get(key) or table.get(local_path.replace("\\", "/").lower())
                else "URL 映射表里没有这条视频")
    return "", "未知 video_source.mode：%s" % mode


_map_cache = {}


class Uploader(object):
    """本地 mp4 -> 公网 URL，并把结果缓存在 outputs 下（不往素材目录写任何东西，P3）。

    缓存键是 绝对路径+字节数+修改时间，改过文件会自动重传。
    """

    def __init__(self, cfg, api):
        self.cfg = cfg
        self.api = api
        self.path = os.path.join(cfg["outputs_dir"], "tk_publish_upload_cache.json")
        self.lock = threading.Lock()          # 并发上传时缓存文件不能同时写
        self.cache = {}
        try:
            if os.path.exists(self.path):
                self.cache = json.load(io.open(self.path, encoding="utf-8"))
        except Exception:
            self.cache = {}

    def _save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        json.dump(self.cache, io.open(self.path, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)

    def __call__(self, local_path):
        key = os.path.abspath(local_path).replace("\\", "/")
        st = os.stat(local_path)
        ent = self.cache.get(key)
        if ent and ent.get("size") == st.st_size and abs(float(ent.get("mtime", 0)) - st.st_mtime) < 2 \
                and (ent.get("url") or "").lower().startswith("http"):
            return ent["url"], ""
        vs = self.cfg["video_source"]
        try:
            got = self.api.s3_put_urls([os.path.basename(local_path)],
                                       channel_id=int(vs.get("channel_id", 1)),
                                       scene=vs.get("scene", "taskVideo"),
                                       bucket=vs.get("bucket_code", "public"))
        except ApiError as e:
            return "", "get-s3-url 失败：%s" % e
        if not got:
            return "", "get-s3-url 返回空 data"
        _, signed, cdn = got[0]
        if not signed:
            return "", "get-s3-url 没给预签名地址"
        log("[上传] %s（%.1fMB）-> S3 ..." % (os.path.basename(local_path), st.st_size / 1048576.0))
        code = http_put_file(signed, local_path)
        if code not in (200, 201, 204):
            return "", "PUT 预签名地址返回 HTTP %s" % code
        url = s3_public_url(signed, cdn)
        with self.lock:
            self.cache[key] = {"url": url, "size": st.st_size, "mtime": st.st_mtime,
                               "uploaded_at": datetime.now().strftime(TIME_FMT)}
            self._save()
        log("[上传] 完成 -> %s" % url)
        return url, ""


def preupload(uploader, paths, workers=4):
    """「批量上传」这一步是纯网络 IO：串行传 22 个文件要等十几分钟，这里并发传。

    每个文件独立成败；失败的不中断整批，后面 resolve_video_url 还会再试一次并带出原因。
    返回 (成功数, [(文件名, 原因)])。
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    total, done, fails, t0 = len(paths), 0, [], time.time()
    with ThreadPoolExecutor(max_workers=max(1, min(int(workers), 8))) as ex:
        futs = {ex.submit(uploader, p): p for p in paths}
        for f in as_completed(futs):
            p = futs[f]
            try:
                url, err = f.result()
            except Exception as e:                        # 单个文件炸了不拖垮整批
                url, err = "", repr(e)[:120]
            done += 1
            if err:
                fails.append((os.path.basename(p), err))
            if done % 5 == 0 or done == total:
                log("[批量上传] %d/%d 完成，用时 %.0f 秒" % (done, total, time.time() - t0))
    for nm, err in fails:
        log("[批量上传][失败] %s：%s" % (nm, err))
    return total - len(fails), fails



def load_url_map(path):
    if path in _map_cache:
        return _map_cache[path]
    table = {}
    if path.lower().endswith(".json"):
        raw = json.load(io.open(path, encoding="utf-8"))
        if isinstance(raw, dict):
            table = {k.lower(): v for k, v in raw.items()}
        else:
            for it in raw:
                k = (it.get("name") or it.get("path") or "").lower()
                table[k] = it.get("url") or it.get("videoUrl") or ""
    else:
        with io.open(path, encoding="utf-8-sig", newline="") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and row[0].strip():
                    table[os.path.basename(row[0].strip()).lower()] = row[1].strip()
    _map_cache[path] = table
    return table


def head_check(url, timeout=15):
    try:
        req = urllib.request.Request(url, method="HEAD")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            size = r.headers.get("Content-Length")
            return True, int(size) if size and size.isdigit() else None, ""
    except Exception as e:
        try:                                        # 有些站不支持 HEAD，退化为 Range GET
            req = urllib.request.Request(url, headers={"Range": "bytes=0-1"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                cr = r.headers.get("Content-Range") or ""
                total = cr.rsplit("/", 1)[-1] if "/" in cr else r.headers.get("Content-Length")
                return True, int(total) if total and total.isdigit() else None, ""
        except Exception as e2:
            return False, None, "%r" % e2


# ----------------------------------------------------------------------------
# 计划构建
# ----------------------------------------------------------------------------
def match_product(cfg, video_name):
    for p in cfg["products"]:
        if any(k and k in video_name for k in p.get("match", [])):
            return p
    return None


def by_video(mapping, video_name):
    """按视频名从字典里取值。台账的 视频名称 带 .mp4，手写的配置一般不带，所以三种写法都认。"""
    if not isinstance(mapping, dict) or not video_name:
        return ""
    stem = re.sub(r"\.(mp4|mov|m4v|avi|mkv|webm)$", "", video_name, flags=re.I)
    for key in (video_name, stem):
        for k, v in mapping.items():
            if str(k).strip() == key:
                return v or ""
    for k, v in mapping.items():
        ks = re.sub(r"\.(mp4|mov|m4v|avi|mkv|webm)$", "", str(k).strip(), flags=re.I)
        if ks and ks == stem:
            return v or ""
    return ""


def _root_index(cfg):
    """把 video_source.root 下面所有 mp4/mov 和 *_TK发布文案.txt 按文件名建一次索引。"""
    if cfg.get("_idx"):
        return cfg["_idx"]
    root = (cfg["video_source"].get("root") or "").replace("\\", "/")
    idx = {"videos": {}, "captions": {}}
    if root and os.path.isdir(root):
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                full = os.path.join(dirpath, fn).replace("\\", "/")
                ext = os.path.splitext(fn)[1].lower()
                if ext in VIDEO_EXTS:
                    idx["videos"].setdefault(fn, full)
                elif "TK发布文案" in full and ext == ".txt":
                    idx["captions"].setdefault(fn, full)
    cfg["_idx"] = idx
    return idx


def repair_row_paths(cfg, rows, log_func=None):
    """台账里的路径写漏了一层（例如 D:/videos/产品A/... 实际在 D:/videos/东南亚产品A/...）。
    只读不改台账（P3）：文件真不存在时按文件名重新定位，找到就用，找不到留给后面的 BLOCK 说清楚。"""
    idx = _root_index(cfg)
    n = 0
    for r in rows:
        for field, key in (("video", "videos"), ("caption_path", "captions")):
            p = (r.get(field) or "").replace("\\", "/")
            if not p or os.path.exists(p):
                continue
            cand = idx[key].get(os.path.basename(p)) or \
                idx[key].get((r.get("name") or "") + ".mp4") or \
                idx[key].get((r.get("name") or "") + "_TK发布文案.txt")
            if cand:
                r[field] = cand
                n += 1
                if log_func:
                    log_func("  %s：%s -> %s" % (r.get("name"), p, cand))
    return n


def build_phone_index(cfg, api, seqs=None):
    try:
        ms = api.mobiles()
    except ApiError as e:
        die("读取云手机列表失败：%s" % e)
    if seqs:
        seqs = [int(s) for s in seqs]
        want = set(seqs)
        chosen = [m for m in ms if m.get("number") in want]
        missing = sorted(want - {m.get("number") for m in chosen})
        if missing:
            die("找不到序列号 %s 对应的账号环境（UI 的「序列」= 接口 number 字段）" % missing)
        chosen.sort(key=lambda m: seqs.index(int(m["number"])))
    else:
        by_serial = {int(m["serialNumber"]): m for m in ms if m.get("serialNumber") is not None}
        by_name = {m.get("name", ""): m for m in ms}
        chosen = []
        for ph in cfg["phones"]:
            m = None
            if ph.get("serial") and int(ph["serial"]) in by_serial:
                m = by_serial[int(ph["serial"])]
            elif ph.get("name") in by_name:
                m = by_name[ph["name"]]
            if m is None:
                log("[警告] 配置的账号环境在 Hubstudio 里找不到：%s -> 跳过" % ph)
                continue
            chosen.append(m)
        if not chosen:
            die("没有可用的账号环境（config.phones 全部匹配失败）")
    # weight 是配置里的档位权重，Hubstudio 的 mobile 记录没有这个字段，摊到记录上才能被分配逻辑读到
    wmap = {}
    for ph in cfg["phones"]:
        if ph.get("weight") is None:
            continue
        for k in (str(ph.get("serial", "")), str(ph.get("name", ""))):
            if k:
                wmap[k] = int(ph["weight"])
    for m in chosen:
        w = wmap.get(str(m.get("serialNumber"))) or wmap.get(str(m.get("name", "")))
        if w:
            m["weight"] = w
    for m in chosen:
        if m.get("status") not in (0, None):
            log("[警告] %s 当前 status=%s，官方要求任务开始时云手机处于关机状态；"
                "脚本不会去开机/关机（P1），但你得知道这台在忙。" % (m.get("name"), m.get("status")))
    return chosen


def window_dt(window_info, base_date, now):
    hh, mm = pick_window_start(window_info)
    if hh is None:
        return None
    dt = base_date.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if dt <= now:
        dt += timedelta(days=1)
    return dt


def parse_when(txt):
    """把 --start-time 的 2026-11-17 20:00 / 20:00 / 11-17 20:00 解析成 datetime。"""
    txt = (txt or "").strip()
    if not txt:
        return None
    now = datetime.now()
    m = re.match(r"^(\d{1,2}):(\d{2})$", txt)            # 只给时刻：今天这个点已过就顺延明天
    if m:
        d = now.replace(hour=int(m.group(1)), minute=int(m.group(2)), second=0, microsecond=0)
        return d + timedelta(days=1) if d <= now else d
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            d = datetime.strptime(txt, fmt)
            return d.replace(second=0, microsecond=0)
        except ValueError:
            continue
    m = re.match(r"^(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})$", txt)   # 11-18 20:00：不写年份，手动补，避开 strptime 的 DeprecationWarning
    if m:
        mo, dd, hh, mi = (int(x) for x in m.groups())
        try:
            d = datetime(now.year, mo, dd, hh, mi)
        except ValueError:
            die("--start-time 里的月日不合法：%s" % txt)
        if d <= now:
            try:
                d = d.replace(year=now.year + 1)
            except ValueError:
                die("--start-time 看不懂：%s" % txt)
        return d
    die("--start-time 看不懂：%s（可用 2026-11-17 20:00 或 20:00）" % txt)


def check_spacing(kind, value):
    """按向导字段的硬边界校验间隔：单手机多视频间隔 10~1440，多手机间隔 5~1440（整数分钟）。"""
    lo = MIN_VIDEO_SPACING_MIN if kind == "video" else MIN_MOBILE_SPACING_MIN
    label = "单手机多视频间隔" if kind == "video" else "多手机间隔"
    if value is None:
        return None
    v = int(value)
    if not (lo <= v <= MAX_SPACING_MIN):
        die("%s=%d 超出向导允许范围 %d~%d 分钟" % (label, v, lo, MAX_SPACING_MIN))
    return v


def wizard_time(t0, mobile_idx, video_idx, mobile_spacing, video_spacing):
    """客户端「批量设置发布时间」的原式（app.asar module 13504/6145 实测反解）：

        triggerTime = 起始时间 + 组内视频序号 * 单手机多视频间隔
                              + 手机分组序号 * 多手机间隔      （单位：分钟）

    外层 map 的 n 是账号环境在表格里首次出现的顺序，内层 l 是该环境下视频的顺序；
    对话框上还写着「设置后将覆盖单独设定的时间，时区为 UTC+8 北京时间，可能略有延迟」。
    """
    return t0 + timedelta(minutes=video_idx * int(video_spacing) + mobile_idx * int(mobile_spacing))


def parse_deadline(txt):
    """「最晚不超过当天 HH:MM」的那个 HH:MM；解析不了就退回默认 22:00。"""
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            d = datetime.strptime(str(txt or DEFAULT_DAY_DEADLINE), fmt)
            return d.hour, d.minute
        except ValueError:
            continue
    die("publish.day_last_deadline 应为 HH:MM，现在是 %r" % (txt,))


def day_capacity(anchor, deadline_hm, step):
    """从 anchor 起、每 step 分钟一条、当天最晚一条不越过 deadline_hm ⇒ 这天这个环境能排几条。"""
    hour, minute = deadline_hm
    dl = anchor.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if dl <= anchor:
        dl += timedelta(days=1)
    if step <= 0:
        return 1
    return max(int((dl - anchor).total_seconds() // 60) // int(step) + 1, 1)


def round_floor_minutes(pb):
    """排期用的间隔下限 = max(硬规则下限, 轮次间隔)。

    自检上线之后，一轮里若有明确失败，重发最早落在「上次发布 + min_spacing_minutes」，
    它自己又要和下一条间隔 min_spacing_minutes，所以轮次间隔取 2×下限才不用把下一轮往后推。
    硬规则（min_spacing_minutes）仍然是底线，重发那条走的是它，不受这个数影响。
    """
    return max(int(pb.get("min_spacing_minutes", MIN_SPACING_FLOOR) or 0),
               int(pb.get("round_spacing_minutes", ROUND_SPACING_DEFAULT) or 0))


def spread_spacing(anchor, deadline_hm, count, min_spacing):
    """2026-11-17 你定的规则：间隔按条数自动摊，但每条不少于 min_spacing，且当天最晚一条不越过 deadline。

        可用分钟 = 当天 anchor -> 当天 deadline
        间隔     = max(min_spacing, 可用分钟 // (条数-1))   （一条时不存在间隔）

    返回 (间隔分钟, 单天容量, 摊满一天需要的天数)。条数多到 min_spacing 也摊不进今天时，
    间隔取 min_spacing、剩下的按单天容量顺延到次日同一时刻起排。
    """
    n = max(int(count), 1)
    hour, minute = deadline_hm
    dl = anchor.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if dl <= anchor:
        dl += timedelta(days=1)
    avail = int((dl - anchor).total_seconds() // 60)
    step = min_spacing if n == 1 else max(int(min_spacing), avail // (n - 1))
    step = check_spacing("video", step)
    cap = day_capacity(anchor, deadline_hm, step)
    return step, cap, (n + cap - 1) // cap


def wizard_time_spread(anchor, mobile_idx, video_idx, mobile_spacing, video_spacing, deadline_hm=None):
    """向导原式 + 「当天不越过 deadline，排满顺延次日」。deadline_hm 为 None 时等价于 wizard_time。"""
    base = anchor + timedelta(minutes=int(mobile_idx) * int(mobile_spacing))
    step = int(video_spacing)
    if deadline_hm is None:
        return base + timedelta(minutes=int(video_idx) * step)
    cap = day_capacity(base, deadline_hm, step)
    day, idx = int(video_idx) // cap, int(video_idx) % cap
    return base + timedelta(days=day, minutes=idx * step)


# ---- 按权重把这批未发布视频分到多个账号环境（2026-11-17 定的档位分配）----

def largest_remainder(total, weights):
    """Hamilton 最大余数法：把 total 按 weights 拆成整数，余数给小数部分最大的，同分时给权重高的。"""
    total, n = int(total), len(weights)
    if n == 0:
        return []
    sw = float(sum(weights)) or 1.0
    raw = [total * w / sw for w in weights]
    base = [int(x) for x in raw]
    rem = total - sum(base)
    order = sorted(range(n), key=lambda i: (-(raw[i] - base[i]), -float(weights[i]), i))
    for i in order[:max(rem, 0)]:
        base[i] += 1
    return base


def cap_adjust(new, used, cap, weights=None):
    """单环境「今日累计 = 新增 + 今日已用」不得越过 cap。

    超出的名额不是随便挪：先按权重从高到低找还有余量的号，且挪完必须仍然
    「今日累计按权重非增、同权重的号相差不超过 1」，否则宁可报出来也不发一个档位倒挂的分配。
    """
    out = list(new)
    if not cap:
        return out, 0
    weights = weights or [1] * len(out)
    room = [max(int(cap) - int(used.get(i, 0)), 0) for i in range(len(out))]
    overflow = 0
    for i in range(len(out)):
        if out[i] > room[i]:
            overflow += out[i] - room[i]
            out[i] = room[i]
    order = sorted(range(len(out)), key=lambda i: (-float(weights[i]), i))
    while overflow:
        day = [out[i] + int(used.get(i, 0)) for i in range(len(out))]
        placed = False
        for i in order:
            if out[i] >= room[i]:
                continue
            hi = [day[j] for j in order if float(weights[j]) > float(weights[i])]
            eq = [day[j] for j in order if float(weights[j]) == float(weights[i]) and j != i]
            if any(d < day[i] + 1 for d in hi) or any(abs(d - (day[i] + 1)) > 1 for d in eq):
                continue          # 低权重的号不能盖过高权重的今日累计
            out[i] += 1
            overflow -= 1
            placed = True
            break
        if not placed:
            break
    return out, overflow


def assign_sequence(quota):
    """把 quota 展开成「第 k 条给第几号环境」的序列：先给额度最多的，同额度给权重靠前的。"""
    remain = [int(q) for q in quota]
    seq = []
    for _ in range(sum(remain)):
        j = max(range(len(remain)), key=lambda k: (remain[k], -k))
        if remain[j] <= 0:
            break
        remain[j] -= 1
        seq.append(j)
    return seq


def used_today_by_mobile(cfg, api, now, phones):
    """今天各账号环境已经占掉的发布条数：扫今天建的 VIDEO_PUBLISH 任务，按运行明细里的 mobileId 归组。

    只读（task/page + task/execution/page），取消的不计。读不到就当 0，并在日志里说清楚。
    """
    idx = {int(p["mobileId"]): k for k, p in enumerate(phones)}
    used = {}
    today = now.strftime("%Y-%m-%d")
    try:
        tasks = api.tasks(60)
    except ApiError as e:
        log("[今日已用] 任务列表读不到（%s）-> 全部按 0 计" % e)
        return used, []
    todays = [t for t in tasks if str(t.get("createdAt") or "")[:10] == today
              and (t.get("capability") or "VIDEO_PUBLISH").upper() == "VIDEO_PUBLISH"]
    for t in todays:
        try:
            es = api.executions(t.get("taskId"), 50)
        except ApiError as e:
            log("[今日已用] %s 明细读取失败：%s" % (t.get("taskId"), e))
            continue
        for e in es:
            if str(e.get("scheduledAt") or e.get("startedAt") or "")[:10] != today:
                continue
            if (e.get("status") or "").upper() == "CANCELED":
                continue
            mid = e.get("mobileId")
            if mid is None or int(mid) not in idx:
                continue
            used[idx[int(mid)]] = used.get(idx[int(mid)], 0) + 1
    return used, todays


# ----------------------------------------------------------------------------
# 今天已经由本脚本提交过的视频（重跑 plan 时自动挡掉，避免重复发布）
# ----------------------------------------------------------------------------
def fired_index(key):
    """schedule 状态里 fired 的键 -> plan detail 的下标；认不出返回 -1。

    轮次打包写的是纯下标（"0"、"14"），另有带 "#" 前缀的老格式（"xxx#3"），
    两种都得认。只认一种的话「今日已提交」就读不到，防重复那两层直接失效——
    2026-11-17 只认带 # 的，结果轮次打包发出去的一条都没被认出来，
    plan 的自动剔除和守护进程的 [避让] 全程空转，只能靠手填 --exclude。
    """
    s = str(key)
    s = s.split("#")[-1] if "#" in s else s
    try:
        return int(s.strip())
    except ValueError:
        return -1


def task_record_names(cfg, today):
    """从两份状态的 tasks 记录里拿今天真发出去的视频名：name -> (来源, 时刻, taskId)。

    fired 那套要靠「plan 文件还在」才能把下标换回名字，plan 一归档/一删就瞎了；
    tasks[*]["videos"] 是发成功当场按名字写的，自成一体。两边都读，取并集，
    这样「今日已发」这份名单不再依赖任何一个附属文件活着。
    不分计划：同一条视频只要今天真发出去过，谁都别再发第二遍。
    """
    out = {}
    for src, st in (("schedule", load_sched_state(cfg)), ("submit", load_state(cfg))):
        for tid, v in (st.get("tasks") or {}).items():
            if not isinstance(v, dict) or str(v.get("createdAt") or "")[:10] != today:
                continue
            for nm in (v.get("videos") or []):
                if nm:
                    out.setdefault(str(nm).lower(),
                                   ("%s任务记录" % src, v.get("createdAt"), tid or v.get("taskId") or ""))
    return out


def names_submitted_today(cfg, today):
    """扫本脚本自己的记录：schedule 的 fired 状态 + submit/plan 记下的 published 台账。

    只认「真的发出去过」的：plan 文件本身不算，否则重跑一次 plan 就把整批挡光。
    客户端向导手工建的任务这里读不到（运行明细里没有视频名），要用 --exclude 显式剔。
    """
    names = set()
    src = []
    p = _sched_state_path(cfg)
    if os.path.exists(p):
        try:
            st = json.load(io.open(p, encoding="utf-8"))
        except Exception as e:
            log("[今日已提交] 读 %s 失败：%s" % (os.path.basename(p), e))
            st = {}
        for plan_key, entries in (st.get("fired") or {}).items():
            pp = resolve_plan_file(cfg, plan_key)
            if not pp:
                continue
            try:
                det = (json.load(io.open(pp, encoding="utf-8")).get("detail")) or []
            except Exception:
                continue
            for k, v in (entries or {}).items():
                if not v or not v.get("taskId"):
                    continue          # 没拿到 taskId = 没真发出去（E114208 那类失败不算）
                if str(v.get("at") or "")[:10] != today:
                    continue
                i = fired_index(k)
                if 0 <= i < len(det) and det[i].get("name"):
                    names.add(str(det[i]["name"]).lower())
                    src.append(plan_key)
        if src:
            log("[今日已提交] schedule 状态里今天已发出 %d 条（来自 %s）"
                % (len(names), "、".join(sorted(set(os.path.basename(s) for s in src)))))
    st2 = load_state(cfg)
    pub = st2.get("published") or {}
    n2 = 0
    for nm, v in pub.items():
        if isinstance(v, dict) and str(v.get("at") or "")[:10] == today:
            names.add(str(nm).lower())
            n2 += 1
    if n2:
        log("[今日已提交] 提交台账里今天记了 %d 条" % n2)
    n3 = len(names)
    names |= set(task_record_names(cfg, today))
    if len(names) > n3:
        log("[今日已提交] 任务记录里再认出 %d 条（plan 文件已归档也认得）" % (len(names) - n3))
    return names


def schedule_lines(plan):
    """排期结果压成几行人话，重点是把「有多少条落到凌晨」直接说出来。"""
    groups = {}
    for s in plan.get("schedule") or []:
        groups.setdefault(s["mobileId"], []).append(s)
    out = []
    for mid, rows in sorted(groups.items()):
        rows.sort(key=lambda x: x["when"])
        first, last = rows[0]["when"], rows[-1]["when"]
        hours = (last - first).total_seconds() / 3600.0
        night = sum(1 for x in rows if x["when"].hour < 7)
        days = len({x["when"].date() for x in rows})
        out.append("        %s(mobileId=%s)：%s -> %s（%d 条 / %d 天 / 跨 %.1f 小时%s）"
                   % (rows[0]["mobile"], mid, first.strftime(TRIGGER_FMT), last.strftime(TRIGGER_FMT),
                      len(rows), days, hours,
                      "，其中 %d 条落在 00:00-07:00 凌晨，建议改 --schedule window_daily" % night
                      if night else ""))
    return out


def dialog_lines(plan):
    """把这份排期换算回向导「批量设置发布时间」对话框要填的三个值。

    这三个值填进客户端，算出来的时刻就和本脚本生成的 发布时间 列完全一致，
    所以粘贴包的时间列不是「只能手抄」——它同时告诉你对话框该敲什么。
    """
    sched = plan.get("schedule") or []
    if not sched:
        return []
    ms = plan.get("mobile_spacing")
    first_mobile = sched[0]["mobileId"]
    base = min(s["when"] for s in sched if s["mobileId"] == first_mobile)
    t0 = min(s["when"] for s in sched)
    out = ["        向导第 4 步「批量设置」对话框填：起始时间 %s / 多手机间隔 %s 分钟 / "
            "单手机多视频间隔 %d 分钟 / 忙时等待 %s"
            % (base.strftime(TRIGGER_FMT), ms if ms is not None else "-",
               plan["spacing"], "勾选" if plan.get("in_queue") else "不勾"),
            "        （最早一条 %s 在 %s；对话框里起始时间留空则按当前时刻起算）"
            % (t0.strftime(TRIGGER_FMT), next(s["mobile"] for s in sched if s["when"] == t0))]
    if (plan.get("days_needed") or 1) > 1:
        out.append("        注意：这批跨 %d 天，「批量设置」对话框只会等差、不会跳天，"
                   "所以别用对话框排——直接粘贴「批量粘贴」sheet，发布时间是逐行给的。"
                   % plan["days_needed"])
    return out


def next_slot(now, anchor, spacing, used, floor=None):
    """在同一账号环境已排的时间里，找出与它们至少相隔 spacing 分钟的最早时刻。

    used 是该 mobileId 已经排定的 datetime 列表。间隔只针对同一台云手机：
    不同账号环境可以同一时刻并行，不必互相让步。
    floor 是硬下限（--after 传进来的上一批最后时刻），比“距离够远”更强：早于它的时刻一律不用。
    """
    t = anchor if anchor > now else now + timedelta(minutes=1)
    if floor is not None and t < floor:
        t = floor
    t = t.replace(second=0, microsecond=0)
    limit = 0
    while limit < 4320:                      # 最多往后挪 3 天，避免死循环
        if (floor is None or t >= floor) and all(abs((t - u).total_seconds()) >= spacing * 60
                                                 for u in used):
            return t
        t += timedelta(minutes=1)
        limit += 1
    return t


def next_slot_window(now, anchor, spacing, used, window, floor=None):
    """只在发布时间窗口内排（如 20:00-23:00、间隔 60 分钟 -> 20:00/21:00/22:00/23:00），
    当天窗口排满就顺延到次日窗口起点。窗口不合法或排不出来就退回普通 next_slot。"""
    sh, sm, eh, em = window or (None, None, None, None)
    if sh is None or eh is None:
        return next_slot(now, anchor, spacing, used, floor)
    if (eh, em) < (sh, sm):                 # 跨零点的窗口，如 20:00-02:00
        eh, em = eh + 24, em
    day = anchor.replace(hour=0, minute=0, second=0, microsecond=0)
    if floor is not None:
        day = min(day, floor.replace(hour=0, minute=0, second=0, microsecond=0))
    for _ in range(60):                     # 最多顺延 60 天
        start = day.replace(hour=sh, minute=sm)
        end = day.replace(hour=0, minute=0) + timedelta(hours=eh, minutes=em)
        t = start if start >= anchor else anchor + timedelta(minutes=1)
        if floor is not None and t < floor:
            t = floor
        while t <= end:
            t = t.replace(second=0, microsecond=0)
            if t > now and (floor is None or t >= floor) \
                    and all(abs((t - u).total_seconds()) >= spacing * 60 for u in used):
                return t
            t += timedelta(minutes=spacing)
        day += timedelta(days=1)
    return next_slot(now, anchor, spacing, used, floor)


def build_plan(cfg, api, args, need_url=True):
    """need_url=False 用于 pack：粘贴包只要内容列，视频 URL 没取到不算 BLOCK（向导会自己上传文件）。"""
    now = datetime.now()
    rows = read_ledger(cfg)
    reasons = {}
    pub_values = [v for v in cfg["ledger"]["unpublished_values"]]
    if args.retry_export:
        fails = read_export(cfg, args.retry_export)
        allow = None
        if args.include_reason:
            allow = {x.strip() for x in args.include_reason.split(",") if x.strip()}
        names, reasons = [], {}
        for f in fails:
            if allow and f["reason"] not in allow:
                continue
            if f["name"]:
                names.append(f["name"])
                reasons[f["name"]] = f["reason"]
        rows = [r for r in rows if r["name"] and r["name"].lower() in {n.lower() for n in names}]
        log("[输入] 来自导出表的失败视频 %d 条" % len(rows))
    else:
        rows = [r for r in rows if r["status"] in pub_values]
        log("[输入] 台账未发布视频 %d 条" % len(rows))

    # 台账状态是人填的，会滞后；导出表是 Hubstudio 的事实。已成功过的默认不再发。
    skipped_published = []
    if not args.allow_republish:
        done, used_exports = published_names_from_exports(cfg)
        if used_exports:
            keep = []
            for r in rows:
                if (r["name"] or "").lower() in done:
                    skipped_published.append(r["name"])
                else:
                    keep.append(r)
            rows = keep
            log("[查重] 参考导出表 %d 份，剔除其中已成功的 %d 条" % (len(used_exports), len(skipped_published)))
        else:
            log("[查重] outputs/exports 下没有导出表，只信台账状态（发过的请手工标已发布）")
        # 台账是手填的会滞后，但本脚本自己今天发出去过的不用你标：状态文件里有名字就自动挡掉
        mine = names_submitted_today(cfg, now.strftime("%Y-%m-%d"))
        if mine:
            keep = [r for r in rows if (r["name"] or "").lower() not in mine]
            log("[查重] 本脚本今天已提交过 %d 条，自动剔掉 %d 条（要硬发用 --allow-republish）"
                % (len(mine), len(rows) - len(keep)))
            rows = keep
    if skipped_published:
        log("[跳过] 导出表显示已成功（台账还没标）：%s" % "、".join(skipped_published[:6])
            + ("…" if len(skipped_published) > 6 else ""))
    if args.only:
        keys = [k.strip() for k in args.only.split(",") if k.strip()]
        rows = [r for r in rows if any(k in (r["name"] or "") or k in (r["video"] or "") for k in keys)]
        log("[筛选] --only %s -> 命中 %d 条" % (keys, len(rows)))
    if getattr(args, "exclude", ""):
        keys = [k.strip() for k in args.exclude.split(",") if k.strip()]
        before = len(rows)
        rows = [r for r in rows if not any(k in (r["name"] or "") or k in (r["video"] or "") for k in keys)]
        log("[筛选] --exclude %s -> 剔掉 %d 条，剩 %d 条" % (keys, before - len(rows), len(rows)))
    fixed_paths = repair_row_paths(cfg, rows, log)
    if fixed_paths:
        log("[路径修正] 台账路径不存在的 %d 处，已按文件名在 %s 下重新定位"
            % (fixed_paths, cfg["video_source"].get("root")))
    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        die("没有需要提交的视频（台账里全是已发布，或筛选条件太窄）")

    seqs = [s.strip() for s in (args.seq or "").split(",") if s.strip()]
    phones = build_phone_index(cfg, api, seqs or None)
    log("[账号环境] " + " / ".join("序列%s %s(mobileId=%s,status=%s)"
                                  % (p.get("number"), p["name"], p["mobileId"], p.get("status"))
                                  for p in phones))

    pb = cfg["publish"]
    min_spacing = check_spacing("video", round_floor_minutes(pb))
    hard_spacing = check_spacing("video", int(pb.get("min_spacing_minutes", MIN_SPACING_FLOOR)))
    deadline_hm = parse_deadline(pb.get("day_last_deadline", DEFAULT_DAY_DEADLINE))
    auto_spacing = bool(pb.get("auto_spacing", True)) and not int(getattr(args, "spacing", 0) or 0)
    spacing = check_spacing("video", max(int(getattr(args, "spacing", 0) or pb.get("spacing_minutes",
                                        min_spacing)), DEFAULT_VIDEO_SPACING_MIN))
    day_cap, days_needed = None, None
    mobile_spacing = check_spacing("mobile", int(getattr(args, "mobile_spacing", 0)
                                        or pb.get("mobile_spacing_minutes", DEFAULT_MOBILE_SPACING_MIN))
                                   or DEFAULT_MOBILE_SPACING_MIN)
    mode = getattr(args, "schedule", "") or pb.get("schedule_mode") or "caption_window"
    start_at = parse_when(getattr(args, "start_time", "") or getattr(args, "single_time", ""))
    if mode == "fixed" and not pb.get("fixed_time"):
        mode = "caption_window"
    uploader = None
    if cfg["video_source"].get("mode") == "s3" and not args.no_upload:
        uploader = Uploader(cfg, api)
        todo = []
        for r in rows:
            lp = (r.get("video") or "").replace("\\", "/")
            if lp and os.path.exists(lp):
                k = os.path.abspath(lp).replace("\\", "/")
                ent = uploader.cache.get(k)
                st = os.stat(lp)
                if not (ent and ent.get("size") == st.st_size
                        and abs(float(ent.get("mtime", 0)) - st.st_mtime) < 2):
                    todo.append(lp)
        if todo:
            mb = sum(os.path.getsize(p) for p in todo) / 1048576.0
            log("[批量上传] 需要新传 %d 个文件 / %.1f MB -> Hubstudio 对象存储（%s/%s），并发 %s 路"
                % (len(todo), mb, cfg["video_source"].get("bucket_code", "public"),
                   cfg["video_source"].get("scene", "taskVideo"),
                   getattr(args, "upload_workers", 4)))
            ok_n, _ = preupload(uploader, todo, workers=getattr(args, "upload_workers", 4))
            log("[批量上传] 成功 %d/%d，缓存里共 %d 个可复用" % (ok_n, len(todo), len(uploader.cache)))
        else:
            log("[批量上传] 全部命中缓存（%d 个），本次不传任何文件" % len(uploader.cache))
    elif args.no_upload:
        log("[批量上传] --no-upload：不传文件，URL 只能用缓存或 url_prefix")
    per_mobile = {}
    used_by_mobile = {}
    after = parse_when(getattr(args, "after", ""))
    floor = after + timedelta(minutes=spacing) if after else None
    if after:
        log("[排期] --after %s：这批第一条不早于 %s（第 2 批接着第 1 批排时用）"
            % (after.strftime(TRIGGER_FMT), floor.strftime(TRIGGER_FMT)))
    sched = []
    missing_title = []
    items, problems = [], []
    fixed = None
    if pb.get("schedule_mode") == "fixed" and pb.get("fixed_time"):
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
            try:
                fixed = datetime.strptime(pb["fixed_time"], fmt)
                break
            except ValueError:
                continue
        if fixed is None:
            die("publish.fixed_time 格式应为 2026-11-15 20:30")

    # wizard 模式：起始时间留空就按「下一整分钟」起算（和向导对话框留空时的行为一致），
    # 分组序号 = 账号环境在表格里首次出现的顺序，与客户端外层 map 的 n 同义。
    mobile_order = [p["mobileId"] for p in phones]
    wanchor = start_at or now.replace(second=0, microsecond=0) + timedelta(minutes=1)
    if floor and wanchor < floor:
        wanchor = floor
    # 按档位权重分配（publish.alloc.enabled 或 --alloc）：不再平均轮转，而是
    # 「这一批新增条数」按 phones[].weight 拆，再拿「今日已用」校验单日上限。
    alloc = pb.get("alloc") or {}
    alloc_on = bool(alloc.get("enabled", False)) or bool(getattr(args, "alloc", False))
    assign, quota, used_today, todays_tasks = None, None, {}, []
    if alloc_on:
        weights = [int(p.get("weight") or 0) for p in phones]
        if sum(weights) <= 0:
            die("publish.alloc 开了但 phones[].weight 没配（每个账号环境要有权重，如 4/2/2/1/1）")
        if alloc.get("used_today", True):
            used_today, todays_tasks = used_today_by_mobile(cfg, api, now, phones)
        else:
            log("[今日已用] alloc.used_today=false：不看今天已发的条数，只按这一批分")
        quota, overflow = cap_adjust(largest_remainder(len(rows), weights), used_today,
                                     int(getattr(args, "max_day", 0)
                                         or alloc.get("max_day_per_mobile", 0)), weights)
        if overflow:
            die("这批 %d 条在「档位不倒挂 + 单号今日累计 ≤ %s 条（含 1号 今天已发的 %d 条）」下装不下，"
                "还差 %d 条。要么放宽 publish.alloc.max_day_per_mobile / --max_day，"
                "要么把多出来的留到明天（今日已用：%s）"
                % (len(rows), alloc.get("max_day_per_mobile"), sum(used_today.values()), overflow,
                   "、".join("%s=%d" % (phones[k]["name"], v) for k, v in sorted(used_today.items()))))
        assign = assign_sequence(quota)
        log("[分配] 未发布 %d 条按权重 %s 拆（今日已用 %d 个任务 / %d 条）："
            % (len(rows), ":".join(str(w) for w in weights), len(todays_tasks), sum(used_today.values())))
        for k, p in enumerate(phones):
            log("        %s(mobileId=%s,权重%s)：新增 %d 条 + 今日已发 %d 条 = 今日累计 %d 条"
                % (p["name"], p["mobileId"], p.get("weight"), quota[k], used_today.get(k, 0),
                   quota[k] + used_today.get(k, 0)))
        per_group = max(quota) if quota else 0
    else:
        per_group = -(-len(rows) // max(len(phones), 1))
    if mode == "wizard":
        if auto_spacing:
            spacing = spread_spacing(wanchor, deadline_hm, per_group, min_spacing)[0]
        caps = [day_capacity(wanchor + timedelta(minutes=mi * mobile_spacing), deadline_hm, spacing)
                for mi in range(len(phones))]
        day_cap = min(caps)
        if alloc_on:
            counts = list(quota)
        else:
            counts = [len(range(i, len(rows), len(phones))) for i in range(len(phones))]
        days_needed = max(-(-counts[i] // max(caps[i], 1)) for i in range(len(phones)))
        over = [phones[i]["name"] for i in range(len(phones)) if counts[i] > caps[i]]
        if auto_spacing:
            log("[排期] 自动间隔：%d 条/环境，从 %s 摊到当天 %02d:%02d ⇒ 每条 %d 分钟（轮次下限 %d，硬规则 %d），"
                "单天每台 %d 条%s"
                % (per_group, wanchor.strftime(TRIGGER_FMT), deadline_hm[0], deadline_hm[1],
                   spacing, min_spacing, hard_spacing, day_cap,
                   "，%s 当天排不下、次日同一时刻接着排（共 %d 天）" % ("、".join(over), days_needed)
                   if days_needed > 1 else "，今天全部排完"))
        log("[排期] 向导同式：起始 %s + 组内序号*%d 分钟 + 分组序号*%s 分钟，当天最晚 %02d:%02d"
            % (wanchor.strftime(TRIGGER_FMT), spacing, mobile_spacing, deadline_hm[0], deadline_hm[1]))

    for i, r in enumerate(rows):
        name = r["name"] or os.path.basename(r["video"] or "")
        local = r["video"].replace("\\", "/") if r["video"] else ""
        cap_path = r["caption_path"].replace("\\", "/") if r["caption_path"] else ""
        mobile = phones[assign[i]] if assign else phones[i % len(phones)]
        problems_row = {"seq": r["seq"], "name": name, "mobile": mobile["name"],
                        "mobileId": mobile["mobileId"], "local": local, "level": "OK",
                        "notes": []}

        if not local or not os.path.exists(local):
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("本地视频不存在：%s" % local)
        else:
            ext = os.path.splitext(local)[1].lower()
            if ext not in VIDEO_EXTS:
                problems_row["level"] = "BLOCK"
                problems_row["notes"].append("视频格式 %s 不支持（仅 mp4/mov）" % ext)
            size_mb = os.path.getsize(local) / 1048576.0
            if size_mb > MAX_VIDEO_MB:
                problems_row["level"] = "BLOCK"
                problems_row["notes"].append("视频 %.0fMB 超过 %dMB 上限" % (size_mb, MAX_VIDEO_MB))
            problems_row["size_mb"] = round(size_mb, 1)

        cap, ok = parse_caption_file(cap_path)
        if not ok:
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("文案文件不存在：%s" % cap_path)
        body = cap["body"].strip()
        # TK 只有一个文案框：【一】标题就是首屏钩子，必须并到正文最前面（Hubstudio 里没有独立标题字段）
        if cfg["caption"].get("include_title", True) and cap.get("title"):
            title = cap["title"].strip()
            if title and title not in body[:len(title) + 40]:
                body = (title + "\n\n" + body).strip()
        if cfg["caption"].get("include_hashtags") and cap.get("hashtags"):
            body = (body + "\n\n" + cap["hashtags"].strip()).strip()
        maxc = int(cfg["caption"].get("max_caption_chars", MAX_CAPTION_DEFAULT))
        if not body:
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("【二】正文解析结果为空")
        elif len(body) > maxc:
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("文案 %d 字符，超过上限 %d" % (len(body), maxc))

        prod = match_product(cfg, name)
        pid = (prod or {}).get("productId", "")
        # 逐条覆盖优先：同一个商品不同视频挂不同小黄车标题（从「小黄车标题-可复制.txt」里挑的）
        ptitle = by_video((prod or {}).get("title_by_video"), name) or (prod or {}).get("productTitle", "")
        if prod is None:
            problems_row["notes"].append("产品映射里没有这个视频名 -> 不挂小黄车")
        elif not PRODUCT_ID_RE.match(pid or ""):
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("商品ID %r 不符合 ^\\d{6,50}$" % pid)
        if pid and not ptitle:
            missing_title.append(name)
        if len(ptitle) > MAX_PRODUCT_TITLE:
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("商品标题 %d 字符 > %d" % (len(ptitle), MAX_PRODUCT_TITLE))

        # 首评：comment_by_video（商家评论话术）优先，其次文案文件里的【六】首评置顶
        comment = by_video(cfg.get("comment_by_video"), name)
        comment_src = "comment_by_video"
        if not comment and cfg["caption"].get("use_pinned_comment"):
            comment = cap["pinned_comment"]
            comment_src = "文案首评" if comment else ""
        if len(comment) > MAX_COMMENT:
            problems_row["notes"].append("首评 %d 字符 > %d，已截断" % (len(comment), MAX_COMMENT))
            comment = comment[:MAX_COMMENT]

        url, err = resolve_video_url(cfg, local, api, uploader)
        if err:
            if need_url:
                problems_row["level"] = "BLOCK"
                problems_row["notes"].append("视频URL：" + err)
            else:
                url = ""
                if problems_row["level"] == "OK":
                    problems_row["level"] = "WARN"
                problems_row["notes"].append("视频URL 未取到（打包不提交，向导里自己上传就行）：%s" % err)
        elif need_url and not args.skip_url_check:
            okh, hsize, herr = head_check(url)
            if not okh:
                problems_row["level"] = "BLOCK"
                problems_row["notes"].append("视频URL 无法访问 %s" % herr)
            else:
                problems_row["url_ok"] = "HTTP 200"
                if hsize and hsize / 1048576.0 > MAX_VIDEO_MB:
                    problems_row["level"] = "BLOCK"
                    problems_row["notes"].append("URL 指向的文件 %.0fMB 超过 %dMB" % (hsize / 1048576.0, MAX_VIDEO_MB))

        # 发布时间
        if problems_row["level"] == "BLOCK":
            problems.append(problems_row)
            continue
        mid = mobile["mobileId"]
        cnt = per_mobile.get(mid, 0)
        if cnt >= min(int(pb.get("chunk_size", 15)), MAX_VIDEOS_PER_MOBILE):
            problems_row["level"] = "BLOCK"
            problems_row["notes"].append("该账号环境本批已排 %d 条，超过单任务上限" % cnt)
            problems.append(problems_row)
            continue
        ex_at = None
        if mode == "immediate":
            ex_at = None
        elif mode == "fixed":
            ex_at = fixed
        elif mode == "wizard":
            # 完全复刻客户端「批量设置发布时间」：起始时间 + 组内序号*单手机间隔 + 分组序号*多手机间隔
            ex_at = wizard_time_spread(wanchor, mobile_order.index(mid), cnt,
                                        mobile_spacing, spacing, deadline_hm)
        else:
            wdt = window_dt(cap, now, now)
            ex_at = start_at or wdt or (now + timedelta(minutes=5 + i * spacing))
        if ex_at is not None:
            booked = used_by_mobile.setdefault(mid, [])
            if mode == "window_daily":
                ex_at = next_slot_window(now, ex_at, spacing, booked, pick_window_range(cap), floor)
            else:
                ex_at = next_slot(now, ex_at, spacing, booked, floor)
            booked.append(ex_at)
            sched.append({"mobileId": mid, "mobile": mobile["name"], "name": name,
                          "seq": r["seq"], "when": ex_at, "window": cap.get("window", "")})
            problems_row["slot_no"] = len(booked)
        per_mobile[mid] = cnt + 1

        # 本地开放接口的 items 结构（实测 2026-11-16 逐字段反查，见文件头「契约」）：
        # 被服务端识别并校验的字段 = mobileId / videoUrl / caption / comment /
        # productId / productTitle / enableAIGC（getLeads、inQueue 同名布尔）。
        # app.asar 里那套 config.videoText 是渲染层内部结构，开放接口不认，videoText 会被忽略。
        item = {"mobileId": int(mid), "videoUrl": url, "videoName": name,
                "caption": body,
                "enableAIGC": bool(pb.get("enable_aigc")),
                "getLeads": bool(pb.get("get_leads")),
                "inQueue": bool(pb.get("queue_when_busy"))}
        item.update(video_meta(local))
        if pid:
            item["productId"] = pid
        if ptitle:
            item["productTitle"] = ptitle
        if comment:
            item["comment"] = comment
        if pb.get("location"):
            item["location"] = pb["location"]
        # triggerTime 保留在计划里供报表/人工核对；开放接口不校验也不解析时间字段。
        item["triggerTime"] = (ex_at or now).strftime(TRIGGER_FMT) if ex_at else ""
        problems_row.update({"triggerTime": item["triggerTime"],
                             "window": cap.get("window", ""),
                             "caption_chars": len(body), "comment": comment,
                             "comment_src": comment_src,
                             "productId": pid, "productTitle": ptitle, "url": url,
                             "phone_serial": mobile.get("number"),
                             "enableAIGC": bool(pb.get("enable_aigc")),
                             "getLeads": bool(pb.get("get_leads")),
                             "reason": reasons.get(name, ""),
                             "local": local})
        items.append({"item": item, "meta": problems_row})
        problems.append(problems_row)

    return {"generated_at": now.strftime(TIME_FMT), "now": now, "phones": phones,
            "items": items, "rows": problems, "spacing": spacing, "mode": mode,
            "auto_spacing": auto_spacing, "min_spacing": min_spacing,
            "day_deadline": "%02d:%02d" % deadline_hm,
            "day_cap": day_cap, "days_needed": days_needed,
            "mobile_spacing": mobile_spacing, "in_queue": bool(pb.get("queue_when_busy")),
            "schedule": sched,
            "start_at": start_at.strftime(TRIGGER_FMT) if start_at else "",
            "used_by_mobile": dict(used_by_mobile),
            "alloc": ({"weights": [p.get("weight") for p in phones], "quota": quota,
                       "used_today": {phones[k]["name"]: v for k, v in sorted(used_today.items())},
                       "max_day": int(getattr(args, "max_day", 0)
                                      or (pb.get("alloc") or {}).get("max_day_per_mobile", 0)),
                       "today_tasks": [t.get("taskName") for t in todays_tasks]} if alloc_on else None),
            "missing_title": missing_title, "skipped_published": skipped_published}


# ----------------------------------------------------------------------------
# 任务名（避免与已有任务重名）
# ----------------------------------------------------------------------------
def next_task_name(api, cfg, prefix):
    """接着现有「新建任务NNN」的最大编号往后排，并且绝不与已存在的任务重名。"""
    seen, max_n = set(), 0
    try:
        for t in api.tasks(size=200):
            nm = t.get("taskName") or ""
            seen.add(nm)
            if nm.startswith(prefix):
                m = re.search(r"(\d+)\s*$", nm)
                if m:
                    max_n = max(max_n, int(m.group(1)))
    except ApiError:
        pass
    n = max_n + 1 if max_n else 1
    while "%s%d" % (prefix, n) in seen:
        n += 1
        if n > max_n + 5000:
            die("任务名推算异常")
    recent = sorted(x for x in seen if x.startswith(prefix))[-5:]
    return "%s%d" % (prefix, n), recent


# ----------------------------------------------------------------------------
# 输出：计划 JSON + 一份 Excel 报告
# ----------------------------------------------------------------------------
def build_payload(cfg, plan, args, task_name):
    """本地开放接口（POST /api/v1/cloud-mobile/tk-automation/video-task/create）的请求体：
    {taskName, queueWhenBusy, items:[扁平字段]}。
    queueWhenBusy 是顶层布尔（实测：给非布尔会报「queueWhenBusy 必须为布尔值 true 或 false」）。"""
    payload = {
        "taskName": task_name,
        "queueWhenBusy": bool(cfg["publish"].get("queue_when_busy")),
        "items": [dict(x["item"]) for x in plan["items"]],
    }
    if args.single_time:                     # --single-time 现在只是排期起点（=--start-time），
        pass                                 # 逐条时刻由 spacing 排开，不再强行压成同一个时间
    return payload


def write_plan_json(cfg, plan, token, payload):
    out_dir = cfg["outputs_dir"]
    os.makedirs(out_dir, exist_ok=True)
    plan_path = os.path.join(out_dir, "tk_publish_plan_%s.json"
                             % datetime.now().strftime("%Y%m%d-%H%M%S"))
    with io.open(plan_path, "w", encoding="utf-8") as f:
        json.dump({"confirm_token": token, "payload": payload,
                   "detail": [{k: v for k, v in x["meta"].items() if k != "local"}
                              for x in plan["items"]],
                   "rows": plan["rows"]}, f, ensure_ascii=False, indent=2)
    return plan_path


def write_report(cfg, plan, payload, token, task_name, plan_path=None):
    xls = os.path.join(cfg["outputs_dir"], "Hubstudio自动上架_计划与校验.xlsx")
    os.makedirs(cfg["outputs_dir"], exist_ok=True)
    write_excel(xls, cfg, plan, payload, token, task_name,
                plan_path or "（本次没有可提交条目，未生成计划文件）")
    return xls


def write_excel(xls, cfg, plan, payload, token, task_name, plan_path):
    import openpyxl
    from openpyxl.styles import Font, Alignment
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    ws = wb.active
    ws.title = "提交计划"
    heads = ["序号", "视频名称", "账号环境", "mobileId", "发布时间(triggerTime)", "文案字符",
             "首评", "商品ID", "商品标题", "AI-Label", "GetLeads", "视频URL"]
    ws.append(heads)
    for x in plan["items"]:
        m = x["meta"]
        ws.append([m.get("seq", ""), m.get("name", ""), m.get("mobile", ""), m.get("mobileId", ""),
                   m.get("triggerTime", ""), m.get("caption_chars", ""), m.get("comment", ""),
                   m.get("productId", ""), m.get("productTitle", ""),
                   "是" if m.get("enableAIGC") else "否",
                   "是" if m.get("getLeads") else "否",
                   m.get("url", "")])
    ws2 = wb.create_sheet("阻断与警告")
    ws2.append(["序号", "视频名称", "账号环境", "级别", "问题", "本地路径"])
    for m in plan["rows"]:
        if m["level"] != "OK":
            ws2.append([m.get("seq", ""), m.get("name", ""), m.get("mobile", ""), m["level"],
                        "；".join(m["notes"]), m.get("local", "")])
    ws3 = wb.create_sheet("参数与判据")
    for k, v in [("确认令牌 token", token), ("任务名 taskName", task_name), ("计划文件", plan_path),
                 ("忙时等待 inQueue", bool(cfg["publish"].get("queue_when_busy"))),
                 ("排期模式 schedule_mode", plan.get("mode", "")),
                 ("单手机多视频间隔(分) 向导字段", plan["spacing"]),
                 ("多手机间隔(分) 向导字段", plan.get("mobile_spacing")),
                 ("排期起点 --start-time", plan.get("start_at") or "（未指定，按文案【四】窗口起点）"),
                 ("单环境单任务上限", MAX_VIDEOS_PER_MOBILE),
                 ("单视频大小上限(MB)", MAX_VIDEO_MB),
                 ("商品ID 判据", r"^\d{6,50}$"),
                 ("商品标题上限", "%d 字符" % MAX_PRODUCT_TITLE),
                 ("首评上限", "%d 字符" % MAX_COMMENT),
                 ("文案上限", "%d 字符" % int(cfg["caption"].get("max_caption_chars", MAX_CAPTION_DEFAULT))),
                 ("可提交条数", len(plan["items"])),
                 ("被拦条数", sum(1 for m in plan["rows"] if m["level"] == "BLOCK")),
                 ("计划生成时间", plan["generated_at"])]:
        ws3.append([k, v])
    wk = wb.create_sheet("排期")
    wk.append(["账号环境", "分组序号(乘多手机间隔)", "组内第几条(乘单手机间隔)", "发布时间",
               "星期", "视频名称", "文案里的时间窗"])
    groups = {}
    for s in plan.get("schedule") or []:
        groups.setdefault(s["mobileId"], []).append(s)
    env_order = []
    for s in plan.get("schedule") or []:
        if s["mobileId"] not in env_order:
            env_order.append(s["mobileId"])
    for mid, rows in sorted(groups.items()):
        rows.sort(key=lambda x: x["when"])
        for n, s in enumerate(rows):
            wk.append([s["mobile"], env_order.index(mid), n, s["when"].strftime(TRIGGER_FMT),
                       s["when"].strftime("%a"), s["name"], s.get("window", "")])
    for s in (ws, ws2, ws3, wk):
        for c in s[1]:
            c.font = bold
        s.freeze_panes = "A2"
        for col in s.columns:
            width = max((len(str(c.value or "")) for c in col), default=8)
            s.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 46)
    wb.save(xls)


def make_token(payload):
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:10]


# ----------------------------------------------------------------------------
# 命令
# ----------------------------------------------------------------------------
def cmd_today(cfg, api, args):
    """一页看完今天：台账未发布多少条、各号今天已经发了多少、按档位这批会怎么分。纯只读。"""
    now = datetime.now()
    rows = [r for r in read_ledger(cfg) if r["status"] in cfg["ledger"]["unpublished_values"]]
    seqs = [s.strip() for s in (args.seq or "").split(",") if s.strip()]
    phones = build_phone_index(cfg, api, seqs or None)
    used, todays = used_today_by_mobile(cfg, api, now, phones)
    mine = names_submitted_today(cfg, now.strftime("%Y-%m-%d"))
    dupes = [r["name"] for r in rows if (r["name"] or "").lower() in mine]
    todo = [r for r in rows if (r["name"] or "").lower() not in mine]
    log("[今天 %s] 台账未发布 %d 条；其中本脚本今天已提交过 %d 条（%s）"
        % (now.strftime("%Y-%m-%d"), len(rows), len(dupes), "、".join(dupes[:8]) or "无"))
    log("        今天已建的视频任务 %d 个：%s"
        % (len(todays), "、".join("%s(%s,共%s条)" % (t.get("taskName"), t.get("status"), t.get("totalCount"))
                                  for t in todays) or "无"))
    for k, p in enumerate(phones):
        log("        %s(mobileId=%s,权重%s,status=%s)：今日已发 %d 条"
            % (p["name"], p["mobileId"], p.get("weight"), p.get("status"), used.get(k, 0)))
    weights = [int(p.get("weight") or 0) for p in phones]
    if sum(weights) and todo:
        pb = cfg["publish"]
        alloc = pb.get("alloc") or {}
        cap = int(getattr(args, "max_day", 0) or alloc.get("max_day_per_mobile", 0))
        quota, overflow = cap_adjust(largest_remainder(len(todo), weights), used, cap, weights)
        log("[分配预览] 待发 %d 条 -> %s%s"
            % (len(todo), "、".join("%s+%d" % (phones[k]["name"], quota[k]) for k in range(len(phones))),
               "（超出单号今日上限 %d 条，差 %d 条发不了）" % (cap, overflow) if overflow else ""))
    return 0


def cmd_phones(cfg, api, args):
    for m in api.mobiles():
        log("%-8s %-26s serial=%-10s status=%-2s 分组=%-16s 备注=%s"
            % (m["mobileId"], m.get("name", ""), m.get("serialNumber", ""), m.get("status", ""),
               m.get("tagName", "") or "", m.get("remark", "") or ""))
    log("status 参考官方文档：任务开始时云手机须处于关机状态（0=关机，2=其他，需以客户端列表为准）")


def cmd_status(cfg, api, args):
    for t in api.tasks(size=args.size):
        log("%-11s %-14s %-22s 总%s 成%s 败%s 取消%s  %s -> %s"
            % (t.get("taskId"), t.get("capability"), t.get("taskName"), t.get("totalCount"),
               t.get("successCount"), t.get("failedCount"), t.get("canceledCount"),
               t.get("createdAt"), t.get("finishedAt") or t.get("status")))


def cmd_watch(cfg, api, args):
    st = load_state(cfg)
    ids = [args.task] if args.task else list(st.get("tasks", {}))
    if not ids:
        die("state 里没有本脚本创建的任务；用 --task mt_xxxxxx 指定")
    import openpyxl
    from openpyxl.styles import Font
    out_dir = cfg["outputs_dir"]
    path = os.path.join(out_dir, "Hubstudio自动上架_运行结果.xlsx")
    while True:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "运行明细"
        ws.append(["taskId", "任务名", "任务状态", "executionId", "mobileId", "状态",
                   "错误码", "失败原因", "计划时间", "开始", "结束", "视频链接"])
        ws2 = wb.create_sheet("任务汇总")
        ws2.append(["taskId", "任务名", "状态", "总数", "成功", "失败", "取消", "创建", "开始", "结束"])
        for tid in ids:
            try:
                d = api.task_detail(tid)
            except ApiError as e:
                ws.append([tid, "", "读取失败", str(e)])
                continue
            try:
                exs = api.executions(tid)
            except ApiError as e:
                exs = []
                ws.append([tid, d.get("taskName"), d.get("status"), "", "", "读取运行明细失败", str(e)])
            ws2.append([tid, d.get("taskName"), d.get("status"), d.get("totalCount"),
                        d.get("successCount"), d.get("failedCount"), d.get("canceledCount"),
                        d.get("createdAt"), d.get("startedAt"), d.get("finishedAt")])
            for e in exs:
                ws.append([tid, d.get("taskName"), d.get("status"), e.get("executionId"),
                           e.get("mobileId"), e.get("status"), e.get("errorCode") or "",
                           e.get("errorMessage") or "", e.get("scheduledAt") or "",
                           e.get("startedAt") or "", e.get("finishedAt") or "",
                           (e.get("result") or {}).get("publishedUrl") or ""])
        for s in (ws, ws2):
            for c in s[1]:
                c.font = Font(bold=True)
            s.freeze_panes = "A2"
        wb.save(path)
        log("[结果] 已写入 %s" % path)
        if args.once:
            break
        time.sleep(max(args.interval, 15))


def state_path(cfg):
    return os.path.join(cfg["outputs_dir"], "tk_publish_state.json")


def load_state(cfg):
    p = state_path(cfg)
    if os.path.exists(p):
        return json.load(io.open(p, encoding="utf-8"))
    return {"tasks": {}, "published": {}}


def save_state(cfg, st):
    os.makedirs(cfg["outputs_dir"], exist_ok=True)
    with io.open(state_path(cfg), "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------------------
# 轮次打包 / 自检的唯一定义（2026-11-17 收拢）
# 之前 cmd_schedule 和 cmd_submit 各写一遍「第 r 轮 = 每台环境的第 r 条」，
# _check_rounds.py / _verify_plan.py 又是外面第三、第四遍——四份逻辑一定会漂移。
# 现在算轮次和核内容只有下面这几个函数，跑批、出包、verify 复核都调它们。
# ----------------------------------------------------------------------------
def group_into_rounds(items, mobile_of):
    """把条目按输入顺序横向切成轮次：同一台环境的第 r 条进第 r 轮。

    一个任务里同一台环境只能出现一次（实测 E010008「items 中的 mobileId 不能重复」），
    所以「批量」只有这一种切法；同一台环境因此天然落到下一轮，间隔由排期步长保证。
    返回 [[第1轮的原始对象...], [第2轮...]]，不改动传入对象。
    """
    seen, rounds = {}, []
    for x in items:
        mid = mobile_of(x)
        r = seen.get(mid, 0)
        seen[mid] = r + 1
        while len(rounds) <= r:
            rounds.append([])
        rounds[r].append(x)
    return rounds


def build_schedule_rows(plan_key, items, detail, fired):
    """把 plan 的 payload/detail 摊成守护进程用的排期行，并用唯一实现切好轮次。

    schedule 与 watchdog 共用这一份：巡检报的「下一轮几点、还剩几条」就是进程真正会发的，
    不再「脚本一套算法、进程另一套算法」。时间解析不了或缺 videoUrl 一律拦——
    开放接口必须有 URL，向导才吃空时间。返回 (rows, by_round)。
    """
    rows = []
    for i, (it, d) in enumerate(zip(items, detail)):
        txt = d.get("triggerTime") or ""
        try:
            when = datetime.strptime(txt, TRIGGER_FMT)
        except ValueError:
            die("计划第 %d 条没有可用 triggerTime（%r），重跑 plan" % (i + 1, txt))
        if not it.get("videoUrl"):
            die("计划第 %d 条缺 videoUrl，schedule 走的是开放接口、必须有 URL；重跑 plan 且不要加 --no-upload" % (i + 1))
        rows.append({"i": i, "when": when, "item": it, "detail": d,
                     "key": "%s#%d" % (plan_key, i), "done": (fired or {}).get(str(i))})
    by_round = group_into_rounds(rows, lambda s: s["item"].get("mobileId"))
    for r, grp in enumerate(by_round):
        for s in grp:
            s["round"] = r
    return rows, by_round


def audit_content(cfg, item, name):
    """逐条核文案/首评/商品标题/商品ID/体积，返回问题列表（空 = 过）。

    上限都从配置取：文案 = caption.max_caption_chars（客户端 tkTextRules 默认 1000），
    首评 = MAX_COMMENT，商品ID 用 products 里 match 命中的那条比对（不再硬编码品目表）。
    """
    out = []
    cap_max = int((cfg.get("caption") or {}).get("max_caption_chars", MAX_CAPTION_DEFAULT))
    cap = str(item.get("caption") or "")
    if not cap.strip():
        out.append("%s 文案为空" % name)
    elif len(cap) > cap_max:
        out.append("%s 文案 %d 字 > %d" % (name, len(cap), cap_max))
    cm = str(item.get("comment") or "")
    if len(cm) > MAX_COMMENT:
        out.append("%s 首评 %d 字 > %d" % (name, len(cm), MAX_COMMENT))
    t = str(item.get("productTitle") or "")
    if t and (len(t) > MAX_PRODUCT_TITLE or set(t) & TITLE_BAD_CHARS):
        out.append("%s 商品标题不合规：%r" % (name, t))
    pid = str(item.get("productId") or "")
    if pid and not PRODUCT_ID_RE.match(pid):
        out.append("%s 商品ID 不合规：%s" % (name, pid))
    prod = match_product(cfg, str(name))
    if prod and str(prod.get("productId") or "") and pid != str(prod["productId"]):
        out.append("%s 商品ID %s 与配置里 %s 的 %s 不符"
                   % (name, pid, "/".join(prod.get("match") or []), prod["productId"]))
    if float(item.get("videoSize") or 0) / 1048576.0 > MAX_VIDEO_MB:
        out.append("%s 体积 %.1fMB > %dMB"
                   % (name, float(item.get("videoSize") or 0) / 1048576.0, MAX_VIDEO_MB))
    return out


def audit_spacing(times_by_mobile, min_gap, last_deadline, when=None):
    """核每台环境的相邻间隔与当天截止：times_by_mobile = {环境名: [datetime 已升序]}。"""
    out = []
    for mob, ts in (times_by_mobile or {}).items():
        gaps = [int((b - a).total_seconds() // 60) for a, b in zip(ts, ts[1:])]
        if [g for g in gaps if g < min_gap]:
            out.append("%s 间隔 %s < %d 分钟" % (mob, gaps, min_gap))
        if ts and last_deadline:
            limit = ts[-1].replace(hour=last_deadline[0], minute=last_deadline[1],
                                   second=0, microsecond=0)
            if ts[-1] > limit:
                out.append("%s 最晚一条 %s 越过当天 %02d:%02d"
                           % (mob, ts[-1].strftime("%H:%M"), last_deadline[0], last_deadline[1]))
    return out


def spacing_advice(times_by_mobile, floor, hard):
    """够硬规则但低于轮次下限的，只提醒不 BLOCK（今天已排好的 30 分钟轮次不该被推翻）。"""
    out = []
    if floor <= hard:
        return out
    for mob, ts in (times_by_mobile or {}).items():
        gaps = [int((b - a).total_seconds() // 60) for a, b in zip(ts, ts[1:])]
        low = [g for g in gaps if hard <= g < floor]
        if low:
            out.append("%s 间隔 %s 够硬规则 %d 分钟，但低于轮次下限 %d 分钟："
                       "自检要是查出明确失败，重发得占掉一个空档，下一轮就得往后推"
                       % (mob, low, hard, floor))
    return out


def cmd_submit(cfg, api, args):
    if not args.token:
        die("submit 必须带 --token（先运行 plan 拿令牌）")
    latest = find_latest_plan(cfg)
    if not latest:
        die("没有找到计划文件，请先运行 plan")
    data = json.load(io.open(latest, encoding="utf-8"))
    payload = data["payload"]
    if data["confirm_token"] != args.token:
        die("token 与最近一次 plan 不一致（%s != %s）。重新跑 plan 再用新 token。"
            % (args.token, data["confirm_token"]))
    if not payload.get("items"):
        die("计划文件里没有 items（%s）；老格式的 plan 不能提交，重跑 plan 生成新 token" % latest)
    items = payload["items"]
    detail = data.get("detail") or []
    # 一个任务只能有一个云手机（实测 E010008「items 中的 mobileId 不能重复」，2026-11-17）。
    # 所以批量 = 按环境切片，一台一个任务；这也和客户端向导「一台环境一个任务」的行为一致。
    groups = {}
    for i, it in enumerate(items):
        mid = it.get("mobileId")
        nm = detail[i].get("mobile") if i < len(detail) else ""
        g = groups.setdefault(mid, {"name": nm or str(mid), "items": []})
        g["items"].append(dict(it))
    log("[提交] %s 共 %d 条 / %d 个账号环境 -> %s"
        % (payload.get("taskName"), len(items), len(groups), latest))
    for g in groups.values():
        if len(g["items"]) > MAX_VIDEOS_PER_MOBILE:
            die("%s 这一台 %d 条，超过单任务 %d 条上限；重跑 plan 拆批"
                % (g["name"], len(g["items"]), MAX_VIDEOS_PER_MOBILE))

    made, failed_groups = [], []
    # 实测（2026-11-17 两次 E010008）：开放接口的 create 要求 items 里 mobileId 互不相同，
    # 也就是「一个任务 = 每台环境一条」，不是「一台环境一次甩 15 条」。客户端向导内部没这个限制。
    # 所以批量按轮次切：第 r 轮 = 每台环境的第 r 条，最多 5 条一个任务，同一台天然排在下一轮。
    base = payload.get("taskName")
    lanes = sorted(groups.values(), key=lambda g: -len(g["items"]))
    name_of = {str(mid): g["name"] for mid, g in groups.items()}
    rounds = group_into_rounds([it for g in lanes for it in g["items"]],
                               lambda it: it.get("mobileId"))
    orders = []
    for r, batch in enumerate(rounds):
        who = [str(name_of.get(str(it.get("mobileId")), "")).split("-")[0] for it in batch]
        orders.append(("%s-R%d" % (base, r + 1),
                       {"name": "+".join(who), "items": batch, "mobileId": None,
                        "videos": [str(it.get("videoName") or "") for it in batch]}))
    log("[批量] %d 轮，每轮最多 %d 台：%s"
        % (len(rounds), len(lanes), "，".join("%s(%d 条)" % (o[0].rsplit("-R", 1)[-1], len(o[1]["items"]))
                                             for o in orders)))
    for tname, g in orders:
        body = {"taskName": tname, "queueWhenBusy": payload.get("queueWhenBusy", True),
                "items": g["items"]}
        try:
            res = api.create_video_task(body)
        except ApiError as e:
            log("[被拒] %s（%s，%d 条）code=%s msg=%s requestId=%s"
                % (tname, g["name"], len(g["items"]), e.res.get("code"), e.res.get("msg"),
                   e.res.get("requestId")))
            failed_groups.append((tname, g["name"], len(g["items"])))
            continue
        tid = res.get("taskId") or (res.get("task") or {}).get("taskId") or ""
        log("[成功] %s %-22s %2d 条 -> taskId=%s" % (tname, g["name"], len(g["items"]), tid))
        made.append((tid, tname, g["name"], len(g["items"]), g["videos"]))

    if made:
        st = load_state(cfg)
        st.setdefault("tasks", {})
        st.setdefault("published", {})
        for tid, tname, gname, n, vids in made:
            st["tasks"][tid] = {"taskName": tname, "createdAt": datetime.now().strftime(TIME_FMT),
                                "items": n, "plan": latest, "mobile": gname,
                                "videos": [v for v in vids if v]}
            # 逐条记名字：plan 的自动剔除（names_submitted_today）认的就是这份，
            # 以前 submit 只记任务不记名字，导致 submit 发过的一律要手填 --exclude。
            for v in vids:
                if v:
                    st["published"][v.lower()] = {"at": datetime.now().strftime(TIME_FMT),
                                                  "taskId": tid, "taskName": tname}
        save_state(cfg, st)
    if failed_groups:
        log("[未完] 这些环境没建成，可直接重跑本命令（已建好的环境不会重复提交，看上面的 taskId）：%s"
            % "；".join("%s(%s,%d 条)" % f for f in failed_groups))
    log("[核对] python tk_publish.py status --size %d 看任务；跑完 watch --once 出运行结果，再 ledger 标台账"
        % max(len(made) + 2, 5))


def cmd_ledger(cfg, api, args):
    """把成功发布的视频在台账里标成已发布。默认只输出到 outputs，--write-back 才改原表。"""
    st = load_state(cfg)
    done = {}
    for tid in st.get("tasks", {}):
        try:
            for e in api.executions(tid):
                if e.get("status") == "SUCCEEDED":
                    url = (e.get("result") or {}).get("publishedUrl") or ""
                    done[url] = e
        except ApiError as ex:
            log("[警告] %s 读取失败：%s" % (tid, ex))
    rows = read_ledger(cfg)
    hits = [r["name"] for r in rows if r["status"] in cfg["ledger"]["unpublished_values"]]
    log("成功发布链接 %d 条；台账未发布 %d 条" % (len(done), len(hits)))
    log("说明：本地 API 的运行明细不含视频名，无法自动逐条对应。"
        "请核对后用 plan 的清单手工回写，或用 --mark <视频名,视频名> 指定要标记为已发布的行。")
    names = set()
    if args.mark:
        names = {n.strip().lower() for n in args.mark.split(",") if n.strip()}
    if not names:
        return
    if not args.write_back:
        import openpyxl
        src = cfg["ledger"]["path"]
        dst = os.path.join(cfg["outputs_dir"], "发布跟踪表_待回写副本.xlsx")
        wb = openpyxl.load_workbook(src, data_only=False)
        ws = wb[cfg["ledger"]["sheet"]]
        col = cfg["ledger"]["columns"]["status"]
        header = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
        si = header.index(col) + 1
        ni = header.index(cfg["ledger"]["columns"]["name"]) + 1
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, ni).value
            if v and str(v).strip().lower() in names:
                ws.cell(r, si).value = cfg["ledger"]["published_value"]
        wb.save(dst)
        log("[副本] 已生成 %s（原表未改动；确认后再用 --write-back）" % dst)
        return
    import openpyxl
    wb = openpyxl.load_workbook(cfg["ledger"]["path"], data_only=False)
    ws = wb[cfg["ledger"]["sheet"]]
    header = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    si = header.index(cfg["ledger"]["columns"]["status"]) + 1
    ni = header.index(cfg["ledger"]["columns"]["name"]) + 1
    n = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, ni).value
        if v and str(v).strip().lower() in names:
            ws.cell(r, si).value = cfg["ledger"]["published_value"]
            n += 1
    wb.save(cfg["ledger"]["path"])
    log("[回写] 已把 %d 行标为%s（原表）" % (n, cfg["ledger"]["published_value"]))


def cmd_cancel(cfg, api, args):
    st = load_state(cfg)
    if args.task not in st.get("tasks", {}):
        die("P2：只允许取消本脚本创建的任务。state 里没有 %s。" % args.task)
    log("[取消] %s（这会把未执行条目置为取消，等同手工「取消发布」）" % args.task)
    log(json.dumps(api.cancel_task(args.task), ensure_ascii=False)[:400])


def plans_dir(cfg):
    """归档目录：outputs/plans。计划文件一天能攒二十几份，全平铺在 outputs 里
    根本看不出哪份在用；旧的挪进来，只留最新几份在手边。"""
    return os.path.join(cfg["outputs_dir"], "plans")


def resolve_plan_file(cfg, name):
    """按「原样路径 -> outputs -> outputs/plans」找一个计划文件，返回绝对路径或 None。

    归档之后 fired 状态里记的还是裸文件名，所有回读都得走这里，
    不然「文件明明在，只是挪了个目录」就会退化成「今日已发认不出来」。
    """
    if not name:
        return None
    p = str(name)
    if os.path.exists(p):
        return p
    for d in (cfg["outputs_dir"], plans_dir(cfg)):
        q = os.path.join(d, os.path.basename(p))
        if os.path.exists(q):
            return q
    return None


def lock_plan_name(cfg):
    """守护进程锁里写死的那份计划名（归档时必须留下它，它还要按名字回读 detail）。"""
    p = _sched_lock_path(cfg)
    try:
        lk = json.load(io.open(p, encoding="utf-8"))
        if pid_alive(int(lk.get("pid") or 0)):
            return os.path.basename(str(lk.get("plan") or ""))
    except Exception:
        pass
    return ""


def archive_old_plans(cfg, keep=3):
    """把 outputs 里多余的旧计划移进 outputs/plans：只移动、不删除，可随时移回来。

    留下最新 keep 份 + 守护进程锁里那份（可能比最新还旧，但它还在按点发）。
    """
    d = cfg["outputs_dir"]
    if not os.path.isdir(d):
        return []
    fs = sorted(f for f in os.listdir(d)
                if f.startswith("tk_publish_plan_") and f.endswith(".json"))
    keep_name = lock_plan_name(cfg)
    moved = []
    for f in fs[:-keep] if keep > 0 else fs:
        if f == keep_name:
            continue
        try:
            os.makedirs(plans_dir(cfg), exist_ok=True)
            dst = os.path.join(plans_dir(cfg), f)
            if os.path.exists(dst):
                continue
            os.rename(os.path.join(d, f), dst)
            moved.append(f)
        except OSError as e:
            log("[归档] %s 没挪动：%s" % (f, e))
    return moved


def find_latest_plan(cfg):
    d = cfg["outputs_dir"]
    if not os.path.isdir(d):
        return None
    fs = sorted(f for f in os.listdir(d) if f.startswith("tk_publish_plan_") and f.endswith(".json"))
    if fs:
        return os.path.join(d, fs[-1])
    pd = plans_dir(cfg)
    if os.path.isdir(pd):
        ps = sorted(f for f in os.listdir(pd)
                    if f.startswith("tk_publish_plan_") and f.endswith(".json"))
        if ps:
            return os.path.join(pd, ps[-1])
    return None


def cmd_plan(cfg, api, args):
    if args.retry_export and not os.path.exists(args.retry_export):
        die("导出表不存在：%s" % args.retry_export)
    plan = build_plan(cfg, api, args)
    prefix = cfg["publish"].get("task_name_prefix", "新建任务")
    task_name, existing = next_task_name(api, cfg, prefix)
    token = None
    plan_path = None
    xls_path = None
    ro = bool(getattr(args, "no_out", False))
    payload = build_payload(cfg, plan, args, task_name)
    if plan["items"] and not ro:
        token = make_token(payload)
        plan_path = write_plan_json(cfg, plan, token, payload)
    if not ro:
        xls_path = write_report(cfg, plan, payload, token or "", task_name, plan_path)
    elif plan["items"]:
        log("[只读] --no-out：不落计划 json、不覆盖「计划与校验」表、不发令牌。"
            "真要发还得去掉它重跑一次——令牌只能来自落盘的那一份，这样 submit 对不上号就会拒绝。")
    blocks = sum(1 for m in plan["rows"] if m["level"] == "BLOCK")
    log("--------------------------------------------------------------")
    log("可提交 %d 条 / 被拦 %d 条 / 台账候选 %d 条" % (len(plan["items"]), blocks, len(plan["rows"])))
    if plan.get("missing_title"):
        log("[提示] %d 条命中的商品没填 productTitle（选填，≤30 字符）：%s"
            % (len(plan["missing_title"]), "、".join(plan["missing_title"][:3]) + ("…" if len(plan["missing_title"]) > 3 else "")))
    if existing:
        log("已存在的任务名（不会重名）：%s" % "，".join(existing))
    log("本次任务名：%s   忙时等待：%s" % (task_name, bool(cfg["publish"].get("queue_when_busy"))))
    log("[排期] 模式 %s / 间隔 %d 分钟（%s）/ 当天最晚 %s / 起点 %s"
        % (plan.get("mode"), plan["spacing"],
           "自动摊，下限 %d" % plan["min_spacing"] if plan.get("auto_spacing") else "固定",
           plan.get("day_deadline") or "-",
           plan.get("start_at") or "此刻（下一整分钟）"))
    for line in schedule_lines(plan):
        log(line)
    for line in dialog_lines(plan):
        log(line)
    for m in plan["rows"]:
        if m["level"] != "OK":
            log("[%s] %s / %s：%s" % (m["level"], m.get("seq"), m.get("name"), "；".join(m["notes"])))
    if token:
        log("计划文件：%s" % plan_path)
        log("确认令牌：%s" % token)
        log("确认后提交：python tk_publish.py submit --token %s" % token)
    elif ro and plan["items"]:
        log("[只读] 可提交 %d 条，本次刻意不生成令牌（--no-out）。" % len(plan["items"]))
    else:
        log("没有任何可提交条目，未生成计划文件（也就没有 token，submit 会拒绝）。")
    if xls_path:
        log("校验报告：%s" % xls_path)
    if plan_path:
        keep = max(1, int(getattr(args, "keep_plans", 3)))
        moved = archive_old_plans(cfg, keep)
        if moved:
            log("[归档] %d 份旧计划移进 outputs/plans/（原地只留最新 %d 份 + 守护进程锁里那份，"
                "读状态的地方两个目录都认）" % (len(moved), keep))


# ----------------------------------------------------------------------------
# pack：向导「批量粘贴」用的粘贴包（对应录屏里黄色标记的「批量上传」+「批量设置发布时间」）
# ----------------------------------------------------------------------------
WIZARD_HEAD = ["账号环境", "素材名称", "封面图", "视频文案", "小黄车", "小黄车商品标题",
               "评论", "添加定位", "AI-Label", "Get Leads", "发布时间", "忙时等待"]
WIZARD_WIDTH = [18, 16, 8, 70, 20, 22, 46, 10, 9, 10, 17, 9]
BOOL_YN = ("N", "Y")


def alt_schedule(plan, mode):
    """同一批视频在另一种排期模式下的时刻，用来在粘贴包里做对照列。

    continuous   = 从窗口起点一路 +spacing（会排进凌晨）
    window_daily = 只在窗口内排，排满顺延次日
    """
    now = datetime.now()
    spacing = plan["spacing"]
    used, out = {}, {}
    for x in plan["items"]:
        m = x["meta"]
        mid = str(m.get("mobileId"))
        info = {"window": m.get("window", ""), "window_start": ""}
        anchor = window_dt(info, now, now) or (now + timedelta(minutes=5))
        booked = used.setdefault(mid, [])
        rng = pick_window_range(info)
        if mode == "window_daily" and rng[2] is not None:
            t = next_slot_window(now, anchor, spacing, booked, rng)
        else:
            t = next_slot(now, anchor, spacing, booked)
        booked.append(t)
        out[(mid, m.get("seq"))] = t
    return out


def write_paste_pack(cfg, items, xls, blocked=None, warns=None, alt=None, alt_name=""):
    """把 items（[{meta, item}]）写成向导「批量编辑」能吃的 12 列粘贴包 + 上传清单。

    cmd_pack 和 cmd_paste 共用这一份，别各写一遍——两遍就会有一遍忘改
    （空列必须真空字符串、别填占位符这类坑就是改了一处漏了另一处）。
    """
    blocked = blocked or []
    warns = warns or []
    out_dir = os.path.dirname(xls) or "."
    os.makedirs(out_dir, exist_ok=True)
    listing = os.path.splitext(xls)[0] + "_批量上传清单.txt"
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "批量粘贴"
    ws.append(WIZARD_HEAD)
    pb = cfg["publish"]
    for x in items:
        m, it = x["meta"], x["item"]
        ws.append([m.get("mobile", ""), m.get("name", ""), "", it.get("caption", ""),
                   it.get("productId", ""), it.get("productTitle", ""), it.get("comment", ""),
                   pb.get("location", "") or "",
                   BOOL_YN[1] if it.get("enableAIGC") else BOOL_YN[0],
                   BOOL_YN[1] if it.get("getLeads") else BOOL_YN[0],
                   it.get("triggerTime", ""),
                   BOOL_YN[1] if it.get("inQueue") else BOOL_YN[0]])
    head_fill = PatternFill("solid", fgColor="FFF3D6")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(vertical="center", horizontal="center")
    for i, w in enumerate(WIZARD_WIDTH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=(c.column in (4, 7)))
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("核对")
    ws2.append(["序号", "素材名称", "账号环境", "发布时间", "文案字符", "是否超1000",
                "商品标题", "标题字符", "首评", "首评字符", "文案里的时间窗",
                alt_name or "对照", "本地视频路径", "已上传的 CDN 链接"])
    red = PatternFill("solid", fgColor="FFD6D6")
    for x in items:
        m, it = x["meta"], x["item"]
        t = (alt or {}).get((str(m.get("mobileId")), m.get("seq")))
        ws2.append([m.get("seq", ""), m.get("name", ""), m.get("mobile", ""),
                    it.get("triggerTime", ""), len(it.get("caption", "")),
                    "超限" if len(it.get("caption", "")) > 1000 else "",
                    it.get("productTitle", ""), len(it.get("productTitle", "")),
                    it.get("comment", ""), len(it.get("comment", "")), m.get("window", ""),
                    t.strftime(TRIGGER_FMT) if t else "", m.get("local", ""), it.get("videoUrl", "")])
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
    for i, w in enumerate([7, 16, 18, 17, 9, 10, 22, 9, 46, 9, 16, 17, 52, 60], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for row in ws2.iter_rows(min_row=2):
        if row[5].value == "超限":
            for c in row:
                c.fill = red
    ws2.freeze_panes = "C2"
    if blocked or warns:
        ws3 = wb.create_sheet("被拦与提醒")
        ws3.append(["序号", "视频名称", "账号环境", "级别", "问题", "本地路径"])
        for m in blocked + warns:
            ws3.append([m.get("seq", ""), m.get("name", ""), m.get("mobile", ""), m["level"],
                        "；".join(m["notes"]), m.get("local", "")])
        for c in ws3[1]:
            c.font = Font(bold=True)
            c.fill = head_fill
        for i, w in enumerate([7, 18, 18, 8, 46, 52], 1):
            ws3.column_dimensions[get_column_letter(i)].width = w
    wb.save(xls)

    with io.open(listing, "w", encoding="utf-8-sig") as f:
        dirs = []
        for x in items:
            p = (x["meta"].get("local") or "").replace("\\", "/")
            d = os.path.dirname(p)
            if d and d not in dirs:
                dirs.append(d)
        f.write("# 「批量上传」按这个顺序选文件（第 1 行 = 粘贴包第 1 行）\n")
        f.write("# 目录：%s\n" % " | ".join(dirs))
        f.write("# 共 %d 个 / %.1f MB\n" % (len(items),
                                          sum(os.path.getsize(x["meta"]["local"]) for x in items
                                              if x["meta"].get("local") and os.path.exists(x["meta"]["local"]))
                                          / 1048576.0))
        for x in items:
            f.write("%s\n" % x["meta"].get("local", ""))
    return xls, listing


def cmd_pack(cfg, api, args):
    plan = build_plan(cfg, api, args, need_url=False)
    if not plan["items"]:
        die("没有可打包的条目（台账全是已发布，或筛选太窄）")
    blocked = [m for m in plan["rows"] if m["level"] == "BLOCK"]
    warns = [m for m in plan["rows"] if m["level"] == "WARN"]
    out_dir = cfg["outputs_dir"]
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    xls = os.path.join(out_dir, args.pack_out or ("TK批量粘贴包_%s.xlsx" % stamp))
    other = alt_schedule(plan, "window_daily" if plan.get("mode") != "window_daily" else "continuous")
    other_name = "窗口内排期(对照)" if plan.get("mode") != "window_daily" else "连续排期(对照)"
    xls, listing = write_paste_pack(cfg, plan["items"], xls, blocked, warns, other, other_name)
    log("[粘贴包] %s（%d 行，列顺序与客户端表格一致，整列复制即可）" % (xls, len(plan["items"])))
    log("[批量上传清单] %s" % listing)
    log("[排期] 模式 %s / 间隔 %d 分钟 / 共 %d 条" % (plan.get("mode"), plan["spacing"], len(plan["items"])))
    for line in schedule_lines(plan):
        log(line)
    for line in dialog_lines(plan):
        log(line)
    log("[改时] 之后想整列换时间，不用重新打包：python tk_publish.py retime --file \"%s\" "
        "--start-time \"11-18 20:00\" --spacing 60 --mobile-spacing 5" % os.path.basename(xls))
    if blocked:
        over = [m for m in blocked if any("单任务上限" in n for n in m["notes"])]
        other = [m for m in blocked if m not in over]
        log("[被拦] %d 条没进包（明细见「被拦与提醒」表）" % len(blocked))
        if over:
            log("        %d 条是「一台账号环境一个任务最多 %d 条」拦的，接着打包用：" % (len(over), MAX_VIDEOS_PER_MOBILE))
            log("        python tk_publish.py pack --only \"%s\" --seq 22 --spacing 60"
                % ",".join(str(m.get("name", "")).replace(".mp4", "") for m in over))
        for m in other[:8]:
            log("        [%s] %s：%s" % (m["level"], m.get("name"), "；".join(m["notes"])))
    if warns:
        log("[提醒] %d 条带 WARN，已进包：%s" % (len(warns), "、".join(str(m.get("name")) for m in warns[:5])
                                            + ("…" if len(warns) > 5 else "")))


def _find_paste_sheet(wb):
    """返回 (工作表, {列名: 列号})；优先「批量粘贴」，否则找表头里有「发布时间」的那张。"""
    for name in (["批量粘贴"] if "批量粘贴" in wb.sheetnames else []) + wb.sheetnames:
        ws = wb[name]
        hdr = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v is not None:
                hdr[str(v).strip()] = c
        if "发布时间" in hdr and "素材名称" in hdr:
            return ws, hdr
    die("这份表里没有「批量粘贴」结构的表（要同时有 素材名称 和 发布时间 两列）")


def _resolve_pack_file(cfg, args):
    if args.file:
        p = args.file if os.path.isabs(args.file) else os.path.join(cfg["outputs_dir"], args.file)
        if not os.path.exists(p):
            die("找不到粘贴包：%s" % p)
        return p.replace("\\", "/")
    out = cfg["outputs_dir"]
    cands = [f for f in glob.glob(os.path.join(out, "*粘贴包*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
    if not cands:
        die("outputs 下没有 *粘贴包*.xlsx，用 --file 指定一个")
    return max(cands, key=os.path.getmtime).replace("\\", "/")


def cmd_paste(cfg, api, args):
    """从已落盘的 plan 直接出「批量粘贴包」：只装还没发出去的，时间沿用 plan，过期就整体后移。

    为什么需要它：能落发布时间的三条路由（task-main/save/v2、
    task-main/batch-update-trigger-time、task-detail/update-trigger-time）都不在
    本地开放接口的代理白名单里（app.asar 那张表共 52 条，逐条核过），直连一律
    E010006；而白名单里的 video-task/create 又不解析任何时间字段。所以「一次把
    整天排完」只剩客户端向导的「批量编辑」这一个入口。守护进程只能按点一条条发，
    人想一次上完就得有张能直接 Ctrl+V 的表——就是这个命令。
    不发任何请求、不读台账、不重传视频，纯本地换算。
    """
    plan_path = (resolve_plan_file(cfg, args.file) if args.file else find_latest_plan(cfg))
    if not plan_path or not os.path.exists(str(plan_path)):
        die("找不到计划文件，先跑一次 plan（或 --file 指定 tk_publish_plan_*.json）")
    data = json.load(io.open(plan_path, encoding="utf-8"))
    det = data.get("detail") or []
    items_pl = (data.get("payload") or {}).get("items") or []
    if len(det) != len(items_pl):
        die("计划文件里 detail(%d) 与 payload.items(%d) 对不上，这份不能用"
            % (len(det), len(items_pl)))
    sent = sent_names_today(cfg)
    keep = []
    for m, it in zip(det, items_pl):
        if str(m.get("name") or "").lower() in sent:
            continue
        keep.append({"meta": m, "item": dict(it)})
    if not keep:
        die("这个计划里没有未发布的条目了（今天发过 %d 条）" % len(sent))

    pb = cfg["publish"]
    min_gap = int(pb.get("min_spacing_minutes", MIN_SPACING_FLOOR))
    round_gap = round_floor_minutes(pb)
    dl = parse_deadline(pb.get("day_last_deadline", "22:00"))
    fmt = TRIGGER_FMT if "%H" in TRIGGER_FMT else "%Y-%m-%d %H:%M"
    for x in keep:
        try:
            x["_t"] = datetime.strptime(str(x["item"].get("triggerTime")), fmt)
        except Exception:
            die("条目 %s 的发布时间 %r 解析不了" % (x["meta"].get("name"), x["item"].get("triggerTime")))
    lead = max(1, int(args.lead))
    first = min(x["_t"] for x in keep)
    shift = datetime.now() + timedelta(minutes=lead) - first
    anchored = ""
    if shift.total_seconds() > 0:
        for x in keep:
            x["_t"] = x["_t"] + shift
        anchored = "（最早一条 %s 距此刻不足 %d 分钟，整批后移 %d 分钟）" % (
            first.strftime("%H:%M"), lead, int(shift.total_seconds() // 60))
    for x in keep:
        x["item"]["triggerTime"] = x["_t"].strftime(fmt)

    by_mobile = {}
    for x in keep:
        by_mobile.setdefault(x["meta"].get("mobile"), []).append(x)
    for xs in by_mobile.values():
        xs.sort(key=lambda x: x["_t"])
    last = max(x["_t"] for x in keep)
    # 判据只有 audit_* 这一份实现：出包和 verify 复核算的是同一套（2026-11-17 收拢）
    problems = ["%s %d 条 > 单任务上限 %d 条" % (mob, len(xs), MAX_VIDEOS_PER_MOBILE)
                for mob, xs in by_mobile.items() if len(xs) > MAX_VIDEOS_PER_MOBILE]
    problems += audit_spacing({m: [x["_t"] for x in xs] for m, xs in by_mobile.items()}, min_gap, dl)
    for a in spacing_advice({m: [x["_t"] for x in xs] for m, xs in by_mobile.items()},
                            round_gap, min_gap):
        log("[提醒] %s" % a)
    for x in keep:
        problems += audit_content(cfg, x["item"], x["meta"].get("name"))
    if problems:
        for p in problems:
            log("[自检不通过] %s" % p)
        die("共 %d 项问题，没出包（P6：未通过不得发布）" % len(problems))

    out_dir = cfg["outputs_dir"]
    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    xls = os.path.join(out_dir, args.pack_out or ("TK批量粘贴包_剩余_%s.xlsx" % stamp))
    keep.sort(key=lambda x: (x["_t"], str(x["meta"].get("mobile"))))
    xls, listing = write_paste_pack(cfg, keep, xls)
    log("[粘贴包] %s（%d 行 / 台数 %d / 无 BLOCK）%s" % (xls, len(keep), len(by_mobile), anchored))
    log("[排期] 最早 %s -> 最晚 %s，同一台最小间隔 %d 分钟"
        % (min(x["_t"] for x in keep).strftime("%H:%M"), last.strftime("%H:%M"),
           min([int((b["_t"] - a["_t"]).total_seconds() // 60)
                for xs in by_mobile.values() for a, b in zip(xs, xs[1:])] or [0])))
    for mob, xs in sorted(by_mobile.items(), key=lambda kv: -len(kv[1])):
        log("        %-26s %d 条  %s" % (mob, len(xs),
                                         " ".join(x["_t"].strftime("%H:%M") for x in xs)))
    log("[上传清单] %s" % listing)
    dpid = ""
    lock_p = _sched_lock_path(cfg)
    if os.path.exists(lock_p):
        try:
            lk = json.load(io.open(lock_p, encoding="utf-8"))
            if pid_alive(int(lk.get("pid") or 0)):
                dpid = str(lk.get("pid"))
        except Exception:
            pass
    log("[用法] 1) python _set_clip.py \"%s\"  把「批量粘贴」sheet 送进剪贴板" % os.path.basename(xls))
    log("       2) 客户端 发视频 -> 批量编辑 -> 点第一格 Ctrl+V -> 确定（别按 Esc，会整窗丢弃）")
    if dpid:
        log("       3) 粘之前先停守护进程，否则到点会再发一遍：PowerShell 执行 Stop-Process -Id %s" % dpid)
    else:
        log("       3) 当前没有守护进程在跑（锁里无可活进程），可以直接粘")
    log("[提醒] 这批 %d 条是本次计划没发完的；今天已发 %d 条已自动剔除。"
        "粘之前若又过了一轮，重跑本命令即可（纯本地、不发请求、不重传视频）" % (len(keep), len(sent)))


def cmd_verify(cfg, api, args):
    """只读复核一份计划：内容一致性 + 轮次打包 + 排期下限，一次核完；不发请求、不写文件、不碰台账。

    它是 2026-11-16 留的两个一次性脚本（_verify_plan.py / _check_rounds.py）的正式替身：
    商品ID 口径改走 config.products（和 plan 构建时同一个 match_product，不再硬编码品目表），
    轮次切法改走 group_into_rounds（和守护进程真正建任务时打的包是同一份实现），
    间隔与当天截止取 config.publish 的 min_spacing_minutes / day_last_deadline。
    P6 说的「先核对信息准确、完整、一致，再完成一次完整自检」就是这条命令；不通过非 0 退出。
    """
    plan_path = resolve_plan_file(cfg, args.file) if args.file else find_latest_plan(cfg)
    if not plan_path or not os.path.exists(str(plan_path)):
        die("找不到计划文件，先跑一次 plan（或用 --file 指定 tk_publish_plan_*.json）")
    data = json.load(io.open(str(plan_path), encoding="utf-8"))
    det = data.get("detail") or []
    items = (data.get("payload") or {}).get("items") or []
    if len(det) != len(items):
        die("计划里 detail(%d) 与 payload.items(%d) 对不上，这份不能用" % (len(det), len(items)))
    log("[复核] %s / %d 条 / 任务名 %s（只读：不发请求、不写文件）"
        % (os.path.basename(str(plan_path)), len(det), (data.get("payload") or {}).get("taskName")))
    pb = cfg["publish"]
    min_gap = int(pb.get("min_spacing_minutes", MIN_SPACING_FLOOR))
    round_gap = round_floor_minutes(pb)
    dl = parse_deadline(pb.get("day_last_deadline", DEFAULT_DAY_DEADLINE))
    problems, warns, fam, lane_t, seen = [], [], {}, {}, {}
    caps = []
    for i, (d, it) in enumerate(zip(det, items)):
        nm = str(d.get("name") or "")
        key = nm.lower()
        if key in seen:
            problems.append("%s 在计划里出现两次（第 %d 与第 %d 条）" % (nm, seen[key] + 1, i + 1))
        seen[key] = i
        fam[nm.split("-")[0]] = fam.get(nm.split("-")[0], 0) + 1
        caps.append(len(str(it.get("caption") or "")))
        problems += audit_content(cfg, it, nm)
        if str(it.get("videoName") or "") != nm:
            problems.append("%s payload.videoName=%r 与台账视频名不符" % (nm, it.get("videoName")))
        url = str(it.get("videoUrl") or "")
        if not url.startswith("https://"):
            problems.append("%s 视频 URL 不是 https：%s" % (nm, url[:60]))
        if url != str(d.get("url") or ""):
            problems.append("%s 校验时用的 URL 和 payload 里的不是同一条" % nm)
        if d.get("level") == "BLOCK":
            problems.append("%s 计划里判 BLOCK：%s" % (nm, "；".join(d.get("notes") or [])))
        elif d.get("notes"):
            warns.append("%s 计划里有提醒：%s" % (nm, "；".join(d.get("notes") or [])))
        if by_video(cfg.get("comment_by_video"), nm) and str(d.get("comment_src") or "") != "comment_by_video":
            warns.append("%s 首评取自 %s（配置里另有专门话术）" % (nm, d.get("comment_src") or "文案首评"))
        txt = str(d.get("triggerTime") or "")
        try:
            when = datetime.strptime(txt, TRIGGER_FMT)
        except ValueError:
            problems.append("%s 发布时间 %r 解析不了（应为 %s）" % (nm, txt, TRIGGER_FMT))
            continue
        if str(it.get("triggerTime") or "") != txt:
            warns.append("%s detail 与 payload 的时间不一致：%s vs %s" % (nm, txt, it.get("triggerTime")))
        lane_t.setdefault(str(d.get("mobile") or it.get("mobileId")), []).append(when)
    for mob in lane_t:
        lane_t[mob].sort()
    problems += audit_spacing(lane_t, min_gap, dl)
    warns += spacing_advice(lane_t, round_gap, min_gap)
    by_round = group_into_rounds(list(zip(det, items)), lambda pair: pair[1].get("mobileId"))
    log("[轮次] %d 轮 / %d 台环境（一个任务 = 每台各一条，E010008 不许同台重复）"
        % (len(by_round), len({it.get("mobileId") for it in items})))
    for r, grp in enumerate(by_round):
        mids = [it.get("mobileId") for _, it in grp]
        if len(set(mids)) != len(mids):
            problems.append("第 %d 轮同一台环境出现两次，建任务会被 E010008 拒" % (r + 1))
        if len(grp) > MAX_VIDEOS_PER_MOBILE:
            problems.append("第 %d 轮 %d 条 > 单任务上限 %d 条" % (r + 1, len(grp), MAX_VIDEOS_PER_MOBILE))
        log("        R%-2d %s  %d 条：%s" % (r + 1, min(str(d.get("triggerTime") or "?")[-5:] for d, _ in grp),
                                            len(grp), "、".join(str(d.get("name")) for d, _ in grp)))
    log("[时间线] 同一台相邻两条硬规则不低于 %d 分钟（新排期建议 %d 分钟，给自检重发留空档），"
        "当天最晚不越过 %02d:%02d" % (min_gap, round_gap, dl[0], dl[1]))
    for mob, ts in sorted(lane_t.items(), key=lambda kv: -len(kv[1])):
        gaps = [int((b - a).total_seconds() // 60) for a, b in zip(ts, ts[1:])]
        log("        %-26s %2d 条  %s->%s  间隔%s" % (mob, len(ts), ts[0].strftime("%H:%M"),
                                                    ts[-1].strftime("%H:%M"), gaps or "-"))
    stale = plan_is_stale(cfg, det)
    if stale:
        problems.append("计划比配置旧（守护进程只认落盘时的 payload）：%s" % "，".join(stale[:6]))
    sent = sent_names_today(cfg)
    again = sorted({str(d.get("name")) for d in det if str(d.get("name") or "").lower() in sent})
    if again:
        log("[重叠] 计划里 %d 条今天已发过：%s（守护进程会自动避让；paste 出包会直接剔掉）"
            % (len(again), "、".join(again)))
    log("[分布] %s | 文案长度 %d-%d 字 | 视频名 %d 个不重复"
        % (fam, min(caps) if caps else 0, max(caps) if caps else 0, len(seen)))
    for w in warns[:10]:
        log("[提醒] %s" % w)
    if warns[10:]:
        log("[提醒] …另有 %d 条同类提醒" % len(warns[10:]))
    if problems:
        for x in problems[:30]:
            log("[自检不通过] %s" % x)
        if len(problems) > 30:
            log("[自检不通过] …另有 %d 项" % (len(problems) - 30))
        die("复核 %d 条，%d 项不通过（P6：未通过不得发布）" % (len(det), len(problems)))
    log("[通过] %d 条全部核对与自检通过，%d 轮可以按计划发。" % (len(det), len(by_round)))



def cmd_watchdog(cfg, api, args):
    """只读巡检：守护进程还活着吗、下一轮几点、今天真发出去几条、台账还剩几条没标。

    零写入、零请求：只读锁文件、两份状态 json、计划 json 和台账（openpyxl 只读打开），
    连本地接口都不碰，所以 Hubstudio 没开也能跑。有问题一律非 0 退出，
    计划任务/cron 拿退出码就能报警——2026-11-14 那条 14:30 的 cron 位子是给它留的，
    目前仍是 disabled：要不要真让机器无人值守地自己报警，等你点头再说。
    排期口径全部走 build_schedule_rows，报的「还剩几条」就是守护进程真正会发的。
    """
    problems, warns = [], []
    lock_p = _sched_lock_path(cfg)
    lk = {}
    if os.path.exists(lock_p):
        try:
            lk = json.load(io.open(lock_p, encoding="utf-8"))
        except Exception as e:
            problems.append("锁文件读不了（%r），挡不住第二个进程" % (e,))
    pid = int(lk.get("pid") or 0)
    alive = pid_alive(pid)
    plan_name = str(lk.get("plan") or "")
    if alive:
        log("[锁] 守护进程活着：PID %s / 计划 %s / %s 起 / 令牌 %s"
            % (pid, plan_name, lk.get("startedAt"), lk.get("token")))
    else:
        log("[锁] %s" % ("没有守护进程在跑（锁文件不存在）" if not lk
                        else "PID %s 已经死了（%s 起，计划 %s）" % (pid or "?", lk.get("startedAt"), plan_name or "?")))
        if pid:
            warns.append("锁里那个 PID 已经不在了，锁文件还留着：不挡新进程，但只看文件容易误判有人在跑")

    path = ""
    if getattr(args, "file", ""):
        # 显式指定优先：人工想巡检某份归档掉的旧计划时，不该被锁里那份顶掉
        path = resolve_plan_file(cfg, args.file) or ""
        if not path:
            problems.append("--file 指定的计划找不到：%s（outputs 和归档目录 plans/ 都翻过了）" % args.file)
    elif plan_name:
        for cand in (os.path.join(cfg["outputs_dir"], plan_name),
                     os.path.join(plans_dir(cfg), plan_name)):
            if os.path.exists(cand):
                path = cand
                break
        if not path:
            problems.append("锁里记的计划文件找不到（%s），归档目录 plans/ 里也没有，那份排期没法巡检" % plan_name)
    else:
        path = find_latest_plan(cfg) or ""
        if path:
            log("[锁] 没有在册进程，改按最新计划巡检：%s" % os.path.basename(str(path)))
    pend, missed, dodged, done_cnt = [], [], [], 0
    if path and os.path.exists(str(path)):
        data = json.load(io.open(str(path), encoding="utf-8"))
        det = data.get("detail") or []
        items = (data.get("payload") or {}).get("items") or []
        if len(det) != len(items):
            problems.append("计划 %s 的 detail(%d) 与 payload.items(%d) 对不上"
                            % (os.path.basename(str(path)), len(det), len(items)))
        elif alive and str(lk.get("token") or "") != str(data.get("confirm_token") or ""):
            problems.append("锁里的令牌与计划文件不一致（%s != %s），进程可能在跑另一份计划"
                            % (lk.get("token"), data.get("confirm_token")))
        elif [d for d in det if not str(d.get("triggerTime") or "").strip()]:
            # 别让 build_schedule_rows 的 die() 把只读巡检搞成「半截报错」：先自己说清楚
            problems.append("计划里有 %d 条没有可用发布时间（这份本来就交给不了 schedule，只能用 submit 或粘贴包发）"
                            % len([d for d in det if not str(d.get("triggerTime") or "").strip()]))
        elif [it for it in items if not it.get("videoUrl")]:
            problems.append("计划里有 %d 条没有 videoUrl（--no-upload 出的计划不能挂给守护进程，重跑 plan）"
                            % len([it for it in items if not it.get("videoUrl")]))
        else:
            key = os.path.basename(str(path))
            fired = (load_sched_state(cfg).get("fired") or {}).get(key, {})
            sched, by_round = build_schedule_rows(key, items, det, fired)
            pend = [s for s in sched if not s["done"]]
            for s in sched:
                d = s["done"] or {}
                if d.get("missed"):
                    missed.append(s)
                elif d.get("skip") == "dup":
                    dodged.append(s)
                elif d.get("taskId"):
                    done_cnt += 1
            log("[排期] %s：%d 条 / %d 轮 —— 已发 %d、避让 %d、判 MISSED %d、还剩 %d"
                % (key, len(sched), len(by_round), done_cnt, len(dodged), len(missed), len(pend)))
            if pend:
                nxt = min(pend, key=lambda s: s["when"])
                grp = [s for s in pend if s["round"] == nxt["round"]]
                log("[下一轮] 第 %d 轮 %d 条 @ %s：%s"
                    % (nxt["round"] + 1, len(grp), nxt["when"].strftime(TRIGGER_FMT),
                       "、".join(str(s["detail"].get("name")) for s in grp)))
                late = (datetime.now() - nxt["when"]).total_seconds() / 60.0
                if not alive:
                    problems.append("没有守护进程在跑，可计划里还剩 %d 条没发（下一轮第 %d 轮 %s）——不重启就永远不会发出去"
                                    % (len(pend), nxt["round"] + 1, nxt["when"].strftime(TRIGGER_FMT)))
                elif late > CATCHUP_DEFAULT_MIN:
                    problems.append("进程活着，但第 %d 轮 %s 已过点 %.0f 分钟（超过 %d 分钟宽限）还没动——多半卡在接口重试，或被人 Ctrl+C 了"
                                    % (nxt["round"] + 1, nxt["when"].strftime("%H:%M"), late, CATCHUP_DEFAULT_MIN))
            else:
                log("[待发] 0 条：这份计划已全部处理完，不需要进程守着")
            if missed:
                problems.append("有 %d 条判了 MISSED 没发出去：%s"
                                % (len(missed), "、".join(str(s["detail"].get("name")) for s in missed[:8])))
            if pend and not alive:
                log("[重启] 接着发就跑：powershell -ExecutionPolicy Bypass -File launch_schedule.ps1 -Token <计划文件 %s 里的 confirm_token>"
                    % os.path.basename(str(path)))

    today = datetime.now().strftime("%Y-%m-%d")
    sent = task_record_names(cfg, today)
    per_src = {}
    for nm, v in sent.items():
        per_src[v[0]] = per_src.get(v[0], 0) + 1
    log("[今日] 状态记录里真发出去 %d 条（%s）"
        % (len(sent), "，".join("%s %d 条" % (k, n) for k, n in sorted(per_src.items())) or "无记录来源"))
    try:
        rows = read_ledger(cfg)
    except SystemExit:
        raise
    except Exception as e:
        rows = []
        warns.append("台账读不了（%r），今天剩多少未发布这条没核到" % (e,))
    if rows:
        unpub = set(cfg["ledger"].get("unpublished_values") or ["未发布", ""])
        un = [r for r in rows if r["status"] in unpub]
        log("[台账] 共 %d 行，标未发布 %d 行（%s）" % (len(rows), len(un), cfg["ledger"]["path"]))
        lag = sorted({str(r["name"]) for r in un if str(r["name"]).lower() in sent})
        if lag:
            warns.append("台账里 %d 条今天已发出却还标未发布：%s（回写走 ledger --mark 名,名，默认只出副本）"
                         % (len(lag), "、".join(lag)))
    for w in warns:
        log("[提醒] %s" % w)
    if problems:
        for x in problems:
            log("[巡检不通过] %s" % x)
        die("巡检发现 %d 项问题（只读，本次没动任何东西）" % len(problems))
    log("[巡检通过] 进程%s，还剩 %d 条待发，台账与状态没打架。"
        % ("活着，守着下一轮" if alive else "不在（没有待发条目，不用守）", len(pend)))



# ----------------------------------------------------------------------------
# 自动自检 + 重试（2026-11-17 加）
#
# 口径来自用户 2026-11-17 的要求：任务建完 15 分钟查一次发布结果；下一轮发布前 15 分钟
# 再查上一轮；没发出去的自动重发，并记下检查结果和重试状态，别重复触发。
#
# 为什么不是「点客户端那个重试按钮」：
# 录屏（2026-11-18 任务详情页）识别到两种重跑：
#   ① 行内「重试」（失败行操作列）→ restartTask → 内部路由
#      /cloud-mobile//v1/cloud-mobile/task-detail/retry，参数体 {id: <执行行内部行id>}；
#      点完该行状态变「等待运行」、计划时间刷新到当下，顶部弹「重试成功」toast，不重选素材/文案。
#   ② 顶部「批量重试」→ batchRetryTask → 内部路由
#      /cloud-mobile/v1/cloud-mobile/task-main/retry，参数体 {taskId}；先弹确认框
#      「确定重试所有失败任务吗？不含取消的失败任务数：N，继续将重新开机并产生费用，请确认。」
# 这两条都不在本地开放接口白名单里（白名单只有 11 条 tk-automation，
# 2026-11-17 从 app.asar 逐条数过；2026-11-18 又拿空 body 直连 6873 复验 task-detail/retry、
# task-main/retry 及多种前缀写法，一律 E010006 请检查请求路径是否正确）。
# 另外开放接口 task/execution/page 只回 executionId/mobileId/status/result/errorCode…，
# 不回行内重试要的内部行 id，即便路由放行也缺锚点。P5 也不许去点界面。所以这里用唯一允许的
# 写路径 video-task/create 给这条视频重建一个发布任务，payload 直接从原计划取（不重传视频、
# 不重新解析文案），效果就是「重新发布对应视频」；这是后台对客户端「行内重试」的等价实现。
#
# 安全边界：只查本脚本自己建的任务（状态里 tasks/retryTasks 记的那些），绝不碰用户手工建的
# mt_* 任务；「结果不明」（已有 publishedUrl，或错误码 4041/E114204）一律不自动重发——
# 那种很可能已经发出去了，再发就是同账号两条一样内容，掉权重（2026-11-17 定的口径）；
# 重发同样守「当天不越过 22:00、一个任务里同台只出现一次、给这台还没发的轮次让位」。
#
# 2026-11-18 用户口径：「完善一下自检遇到失败时，直接点击重试就ok」。落到代码上是三件事：
#   1) 明确失败（FAILED 且没链接、不是 4041/E114204）不再傻等「上次+30 分钟」，当场重发——
#      那次尝试根本没发出去，不占发布间隔，客户端点重试也是这个节奏（--retry-wait-gap 退回旧行为）。
#   2) 别人（用户自己在客户端）点了重试，接口里是同一 executionId 原地重跑：这时撤掉自己排好的
#      pending、不再建重发任务，两边撞车就是同账号双发。今天 11:05:50 就是这个场景。
#   3) 给「一键重发某任务全部失败条目」补个正主子命令 retry（旧日志里让人「手动 retry --force」
#      是句空话，那时根本没这个命令）。
# ----------------------------------------------------------------------------
AMBIGUOUS_CODES = {"4041", "E114204"}      # 发布动作可能已做完、只是没取到状态：绝不自动重发
CHECK_AFTER_MIN = 15                       # 任务建完多少分钟后查第一次
CHECK_BEFORE_NEXT_MIN = 15                 # 下一轮发布前多少分钟再查上一轮
CLEAR_KINDS = ("ok", "cancelled")          # 这两种算「这条查清楚了」，不再重复查、也不会重复重试


def _ts(txt):
    """状态文件里的时刻 -> datetime；解析不了给 None（parse_when 会 die，自检里不能 die）。"""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(str(txt or ""), fmt)
        except ValueError:
            continue
    return None


def _check_state_path(cfg):
    return os.path.join(cfg["outputs_dir"], "tk_publish_check_state.json")


def _check_lock_path(cfg):
    return os.path.join(cfg["outputs_dir"], "tk_publish_check.lock")


def load_check_state(cfg):
    """自检/重试自己那份状态。守护进程写 tk_publish_schedule_state.json，这边只读不写它，
    两个进程各写各的文件才不会互相覆盖（今天 fired 那个坑就是这么踩出来的）。"""
    p = _check_state_path(cfg)
    d = None
    if os.path.exists(p):
        try:
            with io.open(p, encoding="utf-8") as f:
                d = json.load(f)
        except ValueError:
            log("[警告] 自检状态 %s 读不出，这次按空状态走（不动原文件）" % p)
    if not isinstance(d, dict):
        d = {}
    for k, v in (("checks", {}), ("retries", {}), ("retryTasks", {}), ("pending", []),
                 ("externalRetries", [])):
        d.setdefault(k, v)
    return d


def save_check_state(cfg, cs):
    if cs.get("_readOnly"):
        return _check_state_path(cfg)               # 守护进程在跑，手查的结果不落盘
    p = _check_state_path(cfg)
    cs["updatedAt"] = datetime.now().strftime(TIME_FMT)
    tmp = p + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        json.dump(cs, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)
    return p


def load_plan_data(cfg, path, cache=None):
    """读计划文件（带缓存，一次自检里同一份计划要看好几遍）。"""
    if cache is None:
        cache = {}
    key = str(path or "")
    if not key:
        return None
    key = resolve_plan_file(cfg, key) or key        # 计划被归档到 outputs/plans 后也能找回来
    if key not in cache:
        try:
            with io.open(key, encoding="utf-8") as f:
                cache[key] = json.load(f)
        except Exception:
            cache[key] = None
    return cache[key]


def task_item_map(cfg, st, tid, rec, cache=None):
    """mobileId -> {index, name, mobile, plan, at}：把一条执行对回「哪台环境发了哪个视频」。

    execution/page 里没有视频名（只有 mobileId/status/errorCode/errorMessage/publishedUrl/时间），
    名字只能从计划文件反查：fired[计划][下标].taskId 等于这个任务的那些下标。
    一个任务里同一台环境只会出现一次（E010008 逼的），所以 mobileId -> 视频名 是唯一映射，不会串。
    重发建的任务不走 fired，直接用它自己记的 map。
    """
    if rec.get("map"):
        return {str(k): dict(v) for k, v in rec["map"].items()}
    out = {}
    for plan_key, entries in (st.get("fired") or {}).items():
        # 状态里 rec["plan"] 可能是归档前的老路径，必须走 resolve_plan_file 换回现址，
        # 不然「文件明明在，只是挪了个目录」就退化成「这条对不上名字、失败也不敢重发」。
        path = resolve_plan_file(cfg, rec.get("plan") or plan_key)
        data = load_plan_data(cfg, path, cache)
        if not data:
            continue
        det = data.get("detail") or []
        items = (data.get("payload") or {}).get("items") or []
        for k, v in (entries or {}).items():
            if not v or str(v.get("taskId") or "") != str(tid):
                continue
            i = fired_index(k)
            if i < 0 or i >= len(items):
                continue
            d = det[i] if i < len(det) else {}
            out[str((items[i] or {}).get("mobileId"))] = {
                "index": i, "name": d.get("name") or "", "mobile": d.get("mobile") or "",
                "plan": str(path), "at": v.get("at") or ""}
    return out


def classify_execution(row, now, created=None):
    """一条执行 -> (判定, 说明)。判定：ok / pending / retryable / ambiguous / cancelled。"""
    status = str(row.get("status") or "").upper()
    url = str(((row.get("result") or {}) or {}).get("publishedUrl") or "").strip()
    code = str(row.get("errorCode") or "").strip()
    msg = str(row.get("errorMessage") or "").strip()
    if status == "SUCCEEDED":
        return ("ok", "已发布 %s" % url) if url else ("ambiguous", "SUCCEEDED 但没给链接，得人工看一眼")
    if status in ("CANCELLED", "CANCELED"):
        return ("cancelled", "已取消%s" % ("：%s %s" % (code, msg) if (code or msg) else ""))
    if url:
        return ("ambiguous", "状态还是 %s 但已有链接 %s，重发就是同账号两条一样内容" % (status or "空", url))
    if code in AMBIGUOUS_CODES:
        return ("ambiguous", "%s %s（发布动作可能已做完，状态没取到）" % (code, msg or "无说明"))
    if status in ("FAILED", "FAILURE", "ERROR"):
        return ("retryable", "%s %s" % (code or "无错误码", msg or "无说明"))
    late = ""
    if created is not None:
        late = "，距建任务已 %d 分钟" % int((now - created).total_seconds() // 60)
    return ("pending", "%s 还没终态%s" % (status or "空状态", late))


def externally_rerun(prev_item, row):
    """这条执行是不是被人在客户端点了「重试/批量重试」原地重跑了。

    按钮的指纹（2026-11-18 实测同一任务的三条失败行）：executionId 不变，startedAt 被刷到
    更晚的一刻，状态从 FAILED 回到 RUNNING，链接/错误码是重跑后的新值。
    开放接口没有重试路由，所以这种重跑一定是界面侧或别人干的；我们得认出来，别再发一遍。
    返回 (是否重跑, 说明)。
    """
    if not prev_item:
        return False, ""
    a, b = _ts(prev_item.get("startedAt")), _ts(row.get("startedAt"))
    if a and b and b > a:
        return True, "startedAt %s -> %s" % (a.strftime("%H:%M:%S"), b.strftime("%H:%M:%S"))
    pa = str(prev_item.get("status") or "").upper()
    pb = str(row.get("status") or "").upper()
    if pa in ("FAILED", "FAILURE", "ERROR") and pb in ("RUNNING", "PENDING", "QUEUED",
                                                       "WAITING", "SCHEDULED", "CREATED"):
        return True, "%s -> %s" % (pa, pb or "空")
    return False, ""


def active_plan(cfg):
    """锁里那份计划（没锁就取最新一份）-> (path, sched)。sched 里每条带 when/item/detail/round/done。
    读不出、条数对不上、时刻/URL 有空的，一律给空 sched —— 自检不能因为计划文件坏就跟着死。"""
    try:
        with io.open(_sched_lock_path(cfg), encoding="utf-8") as f:
            lk = json.load(f)
    except Exception:
        lk = {}
    path = (resolve_plan_file(cfg, lk.get("plan")) if lk.get("plan") else find_latest_plan(cfg))
    if not path:
        return None, []
    data = load_plan_data(cfg, path)
    if not data:
        return path, []
    items = (data.get("payload") or {}).get("items") or []
    det = data.get("detail") or []
    if len(det) != len(items):
        return path, []
    if [d for d in det if not str(d.get("triggerTime") or "").strip()] \
            or [it for it in items if not it.get("videoUrl")]:
        return path, []          # 这份本来就交给不了守护进程，自检也别硬啃
    fired = (load_sched_state(cfg).get("fired") or {}).get(os.path.basename(path), {})
    try:
        sched, _rounds = build_schedule_rows(os.path.basename(path), items, det, fired)
    except SystemExit:
        return path, []
    return path, sched


def last_publish_by_mobile(cfg, cs, today):
    """每台环境今天最后一次「真建出任务」的时刻：从自检记录里的 scheduledAt/finishedAt 取，
    没查过的再用守护进程 fired 的 at 兜底。重发要拿它算间隔，不能只信计划里的理论时刻。"""
    out = {}
    for rec in (cs.get("checks") or {}).values():
        for it in (rec.get("items") or []):
            if str(rec.get("createdAt") or "")[:10] != today:
                continue
            mid = str(it.get("mobileId"))
            when = _ts(it.get("finishedAt")) or _ts(it.get("scheduledAt")) or _ts(it.get("startedAt"))
            if when and (mid not in out or when > out[mid]):
                out[mid] = when
    st = load_sched_state(cfg)
    cache = {}
    for plan_key, entries in (st.get("fired") or {}).items():
        path = resolve_plan_file(cfg, plan_key)
        data = load_plan_data(cfg, path, cache)
        if not data:
            continue
        items = (data.get("payload") or {}).get("items") or []
        for k, v in (entries or {}).items():
            if not v or not v.get("taskId"):
                continue
            i = fired_index(k)
            when = _ts(v.get("at"))
            if i < 0 or i >= len(items) or when is None or str(v.get("at"))[:10] != today:
                continue
            mid = str((items[i] or {}).get("mobileId"))
            if mid not in out or when > out[mid]:
                out[mid] = when
    return out


def check_pass(cfg, api, args, now=None, reason="", read_only=False, only=None):
    """查一遍今天本脚本建的任务的执行结果，写进自检状态。

    返回 (可重发清单, 统计, cs, st)。默认只读：只有 task/execution/page，不建任务、不改守护进程状态。
    read_only=True 时结果不落盘（自检守护进程正在跑，两边各写各的会把同一条重发两遍）。
    only=[taskId...] 时只查这些任务（retry 子命令用），不在今天名单里的也从 state 里捞出来查；
    配 --force-check 连「建完未满 --check-after 分钟」和「上次已终态」两道跳过也一并绕过。
    """
    now = now or datetime.now()
    st = load_sched_state(cfg)
    cs = load_check_state(cfg)
    if read_only:
        cs["_readOnly"] = True
    after = max(int(getattr(args, "check_after", CHECK_AFTER_MIN)), 0)
    checks = cs.setdefault("checks", {})
    cache = {}
    fcache = {}
    retryable = []
    rerun_keys = set()
    stats = {"tasks": 0, "items": 0, "ok": 0, "pending": 0, "ambiguous": 0,
             "retryable": 0, "cancelled": 0, "unreadable": 0, "skipped": 0, "rerun": 0,
             "next_due": None}
    today = now.strftime("%Y-%m-%d")
    force = bool(getattr(args, "force_check", False))
    watch = [(tid, rec, "schedule") for tid, rec in (st.get("tasks") or {}).items()
             if isinstance(rec, dict) and str(rec.get("createdAt") or "")[:10] == today]
    watch += [(tid, rec, "retry") for tid, rec in (cs.get("retryTasks") or {}).items()
              if isinstance(rec, dict) and str(rec.get("createdAt") or "")[:10] == today]
    stats["not_found"] = []
    if only:
        have = {str(t[0]) for t in watch}
        for tid2, rec2 in (st.get("tasks") or {}).items():        # 昨天建的也能点名重查
            if str(tid2) in {str(o) for o in only} and str(tid2) not in have \
                    and isinstance(rec2, dict):
                watch.append((tid2, rec2, "manual"))
                have.add(str(tid2))
        watch = [w for w in watch if str(w[0]) in {str(o) for o in only}]
        stats["not_found"] = [str(o) for o in only if str(o) not in {str(w[0]) for w in watch}]
    watch.sort(key=lambda x: str(x[1].get("createdAt") or ""))
    for tid, rec, src in watch:
        created = _ts(rec.get("createdAt")) or _ts(rec.get("scheduledFor"))
        if created and not force and (now - created).total_seconds() < after * 60:
            stats["skipped"] += 1
            due = created + timedelta(minutes=after)
            if stats["next_due"] is None or due < stats["next_due"]:
                stats["next_due"] = due                 # 记下最早该查的时刻，别让守护进程空转
            continue                                    # 还没到「建完 15 分钟」，下一圈再来
        prev = checks.get(tid) or {}
        prev_items = {str(i.get("executionId")): i for i in (prev.get("items") or [])
                      if i.get("executionId")}
        if prev.get("terminal") and not getattr(args, "force_check", False):
            stats["skipped"] += 1
            continue                                    # 上次已查成终态，不再重复查（也就不会重复重试）
        stats["tasks"] += 1
        try:
            rows = api.executions(tid, size=50)
        except ApiError as e:
            stats["unreadable"] += 1
            checks[tid] = {"taskName": rec.get("taskName"), "createdAt": rec.get("createdAt"),
                           "checkedAt": now.strftime(TIME_FMT), "reason": reason, "terminal": True,
                           "error": "%s %s" % (e.res.get("code"), e.res.get("msg")),
                           "items": prev.get("items") or [], "source": src}
            log("[检查] %s taskId=%s 检查=%s 读取失败：%s（任务不在了就到此为止，不重试）"
                % (rec.get("taskName"), tid, now.strftime("%H:%M"), e))
            continue
        names = task_item_map(cfg, st, tid, rec, cache)
        if not names and len(rows) == 1 and len(rec.get("videos") or []) == 1:
            # 计划文件彻底没了，但「一条一任务」这种还能靠任务记录里的唯一名字对上，
            # 至少日志里不是「对不上名字」；重发仍然要求能取回 payload，所以这里不带 index。
            names = {str(rows[0].get("mobileId")): {"index": None, "name": rec["videos"][0],
                                                    "mobile": (rec.get("mobiles") or [""])[0],
                                                    "plan": "", "at": rec.get("createdAt") or ""}}
        items, cnt = [], {}
        for x in rows:
            m = names.get(str(x.get("mobileId"))) or {}
            kind, why = classify_execution(x, now, created)
            relaxed = kind == "ambiguous" and bool(getattr(args, "retry_ambiguous", False))
            if relaxed:
                kind, why = "retryable", "按 --retry-ambiguous 放宽：%s" % why
            # 客户端「重试/批量重试」是同一 executionId 原地重跑：认出来就别再自己发一遍，
            # 之前排好的重发也撤掉（今天 11:05:50 那三条就是这么撞上的）。
            rerun, rerun_why = externally_rerun(prev_items.get(str(x.get("executionId"))), x)
            if rerun and kind == "retryable" and not relaxed:
                kind, why = "pending", "已被客户端重试（%s），先看它自己的结果｜原判定：%s" % (rerun_why, why)
            elif rerun:
                why = "%s｜已被客户端重试（%s）" % (why, rerun_why)
            # 脚本自己已经补发过的：原任务里那行会永远停在 FAILED，不能继续当「还要重试」报出去，
            # 更不能喊交人工——补发可能正在跑、也可能已经发出去了，照旧记录手工再发就是同账号双发。
            if kind == "retryable" and not relaxed:
                fkey = str(m.get("name") or "").lower()
                frec = (cs.get("retries") or {}).get(fkey) if fkey else None
                if frec and frec.get("retryTaskId") and str(frec.get("retryTaskId")) != str(tid):
                    fkind, fwhy = follow_up_state(api, frec, fcache)
                    if fkind in ("ok", "pending"):
                        ftag = str(frec.get("retryTaskName") or frec.get("retryTaskId"))
                        kind = "ok" if fkind == "ok" else "pending"
                        why = "本脚本已补发（%s）：%s｜原任务那行仍记 %s" % (ftag, fwhy, why)
            items.append({"mobileId": x.get("mobileId"), "name": m.get("name") or "",
                          "mobile": m.get("mobile") or "", "executionId": x.get("executionId"),
                          "status": x.get("status"), "errorCode": x.get("errorCode"),
                          "errorMessage": x.get("errorMessage"),
                          "url": ((x.get("result") or {}) or {}).get("publishedUrl") or "",
                          "scheduledAt": x.get("scheduledAt"), "startedAt": x.get("startedAt"),
                          "finishedAt": x.get("finishedAt"), "kind": kind, "why": why,
                          "rerun": rerun_why if rerun else "", "relaxed": relaxed,
                          "index": m.get("index"), "plan": m.get("plan")})
            cnt[kind] = cnt.get(kind, 0) + 1
            stats["items"] += 1
            stats[kind] = stats.get(kind, 0) + 1
            log("[检查] 任务=%s taskId=%s 检查=%s 视频=%s 环境=%s 状态=%s 判定=%s 原因=%s"
                % (rec.get("taskName"), tid, now.strftime("%H:%M"),
                   m.get("name") or "（对不上计划里的名字）",
                   (m.get("mobile") or x.get("mobileId")), x.get("status"), kind, why))
            if rerun:
                stats["rerun"] += 1
                rerun_keys.add(str(m.get("name") or "").lower())
                cs.setdefault("externalRetries", []).append({
                    "at": now.strftime(TIME_FMT), "taskId": tid, "taskName": rec.get("taskName"),
                    "executionId": x.get("executionId"), "mobileId": x.get("mobileId"),
                    "name": m.get("name") or "", "mobile": m.get("mobile") or "",
                    "was": "%s %s" % ((prev_items.get(str(x.get("executionId"))) or {}).get("status"),
                                      (prev_items.get(str(x.get("executionId"))) or {}).get("errorCode")),
                    "now": "%s %s" % (x.get("status"), x.get("errorCode")), "how": rerun_why})
                cs["externalRetries"] = cs["externalRetries"][-200:]
                log("[外部重试] 任务=%s taskId=%s 检查=%s 视频=%s 环境=%s 执行=%s：%s —— "
                    "接口里这条已经被重跑（客户端点了重试），本轮不建重发任务，排好的重发也撤掉"
                    % (rec.get("taskName"), tid, now.strftime("%H:%M"), m.get("name") or "?",
                       (m.get("mobile") or x.get("mobileId")), x.get("executionId"), rerun_why))
                continue
            if kind != "retryable":
                continue
            if m.get("index") is None:
                log("[不重发] 这条对不上计划里的下标，payload 取不出来，只能人工看：%s" % rec.get("taskName"))
                continue
            retryable.append({"taskId": tid, "taskName": rec.get("taskName"), "plan": m.get("plan"),
                              "index": m["index"], "name": m.get("name") or "", "mobileId": x.get("mobileId"),
                              "mobile": m.get("mobile") or "", "sentAt": m.get("at") or "",
                              "why": why, "executionId": x.get("executionId"), "source": src,
                              # 「结果不明被放宽」的条目不享受当场重发，仍旧要等满间隔：
                              # 它可能已经发出去了，抢那 30 分钟没意义，撞车才要命。
                              "relaxed": relaxed,
                              # 去重单位是「这条视频」，不是「某个任务的某一行」：重发建的任务再失败，
                              # 键还是同一个名字，max_retry 才拦得住「重试套重试」。
                              "key": (str(m.get("name") or "").lower()
                                      or "%s#%s" % (tid, m.get("index")))})
        checks[tid] = {"taskName": rec.get("taskName"), "createdAt": rec.get("createdAt"),
                       "checkedAt": now.strftime(TIME_FMT), "reason": reason, "source": src,
                       "terminal": bool(items) and all(i["kind"] in CLEAR_KINDS for i in items),
                       "counts": cnt, "items": items}
    rerun_keys.discard("")
    if rerun_keys and cs.get("pending"):
        keep = [p for p in cs["pending"] if str(p.get("key") or "").lower() not in rerun_keys]
        if len(keep) != len(cs["pending"]):
            dropped = [str(p.get("name") or "?") for p in cs["pending"]
                       if str(p.get("key") or "").lower() in rerun_keys]
            cs["pending"] = keep
            log("[撤重发] %d 条不再自己发（客户端那边已经重跑）：%s"
                % (len(dropped), "、".join(dropped)))
    save_check_state(cfg, cs)
    return retryable, stats, cs, st


def retry_slot(cfg, args, cs, mobile_id, sent_at, now, sched, immediate=False):
    """这条重发落在什么时候。

    immediate=True（明确失败：FAILED 且没有链接，也不是 4041/E114204）= 当场重发，跟客户端点
      「重试」一个节奏，不再等满 min_spacing_minutes——上一次根本没发出去任何东西，不该占发布间隔
      （用户 2026-11-18 口径「自检遇到失败时，直接点击重试就ok」）。
    immediate=False（结果不明被 --retry-ambiguous 放宽，或 --retry-wait-gap 要退回旧行为）
      = 仍旧 max(now, 这台今天最后一次发布 + 间隔下限)。
    两种都要躲开这台接下来还没发的轮次（前后各 >= 间隔），且不越过当天 day_last_deadline；
    --retry-now 连让位都不管（自己确认要把同一台间隔压短才用）。返回 (时刻 or None, 说明)。
    """
    gap = max(int(cfg["publish"].get("min_spacing_minutes", MIN_SPACING_FLOOR)), 1)
    hour, minute = parse_deadline(cfg["publish"].get("day_last_deadline", DEFAULT_DAY_DEADLINE))
    dl = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if dl <= now:
        return None, "今天 %s 已过，这条不排（明天重跑 plan 再说）" % dl.strftime("%H:%M")
    when, why_slot = now, "明确失败没发出去，按客户端点重试的节奏当场重发"
    last = None
    if not immediate:
        last = last_publish_by_mobile(cfg, cs, now.strftime("%Y-%m-%d")).get(str(mobile_id))
        if last:
            earliest = last + timedelta(minutes=gap)
            if earliest > when:
                when = earliest.replace(second=0, microsecond=0)
        why_slot = "上次 %s + %d 分钟" % ((last.strftime("%H:%M") if last else "无记录"), gap)
    if getattr(args, "retry_now", False):
        return now, "立刻重发（--retry-now，连让位也不等）"
    nxt = [s["when"] for s in sched if not s["done"] and str(s["item"].get("mobileId")) == str(mobile_id)]
    if nxt and when < nxt[0] and nxt[0] - when < timedelta(minutes=gap):
        when = nxt[0] + timedelta(minutes=gap)
        log("[让位] %s 这台 %s 的下一轮在 %s，重发往后挪到 %s，免得把同一台的间隔压短"
            % (sent_at or "?", mobile_id, nxt[0].strftime("%H:%M"), when.strftime("%H:%M")))
        why_slot = "%s；又给 %s 的下一轮让位到 %s" % (why_slot, nxt[0].strftime("%H:%M"),
                                                     when.strftime("%H:%M"))
    if when > dl:
        return None, "挪完就越过今天 %s 了（间隔规则优先），这条交人工/明天" % dl.strftime("%H:%M")
    return when, why_slot


def create_retry_tasks(cfg, api, args, cs, batch, now, base_name):
    """把一批（每台环境最多一条）重发条目建成任务。"""
    pick = task_name_guard(api)
    prefix = "%s-RT" % base_name
    want_n = 1 + len([1 for v in (cs.get("retryTasks") or {}).values()
                      if str(v.get("createdAt") or "")[:10] == now.strftime("%Y-%m-%d")])
    tname = pick(prefix, want_n)[0]
    one = {"taskName": tname,
           "queueWhenBusy": bool(cfg["publish"].get("queue_when_busy")),
           "items": [b["item"] for b in batch]}
    who = "、".join("%s:%s" % (str(b["mobile"]).split("-")[0], b["name"]) for b in batch)
    log("[重发] %s %d 条（%s）原因=%s" % (tname, len(batch), who,
                                        "；".join(b["why"] for b in batch)[:180]))
    if args.dry_run:
        log("[dry-run] 没真建任务：payload 就绪（%s）" % who)
        return None
    try:
        res = api.create_video_task(one)
    except ApiError as e:
        code = str(e.res.get("code"))
        for b in batch:
            rec = cs["retries"].setdefault(b["key"], {"attempts": 0, "history": []})
            rec["attempts"] = int(rec.get("attempts") or 0) + 1
            rec.setdefault("history", []).append({"at": now.strftime(TIME_FMT), "error": code,
                                   "msg": str(e.res.get("msg"))[:120]})
        log("[重发被拒] %s -> code=%s msg=%s（记一次尝试，不硬撞）" % (tname, code, e.res.get("msg")))
        save_check_state(cfg, cs)
        return None
    tid = res.get("taskId") or (res.get("task") or {}).get("taskId") or ""
    cs["retryTasks"][tid] = {
        "taskName": tname, "createdAt": now.strftime(TIME_FMT), "items": len(batch),
        "plan": batch[0]["plan"], "round": "重发",
        "mobiles": sorted({str(b["mobile"]) for b in batch}),
        "videos": [b["name"] for b in batch],
        "map": {str(b["mobileId"]): {"index": b["index"], "name": b["name"],
                                     "mobile": b["mobile"], "plan": b["plan"],
                                     "at": now.strftime(TIME_FMT)} for b in batch}}
    for b in batch:
        rec = cs["retries"].setdefault(b["key"], {"attempts": 0, "history": []})
        rec["attempts"] = int(rec.get("attempts") or 0) + 1
        rec.update({"name": b["name"], "mobileId": b["mobileId"], "mobile": b["mobile"],
                    "origTaskId": b["taskId"], "origExecution": b["executionId"],
                    "firedAt": now.strftime(TIME_FMT), "retryTaskId": tid,
                    "retryTaskName": tname, "why": b["why"]})
        rec.setdefault("history", []).append({"at": now.strftime(TIME_FMT), "taskId": tid})
    log("[重发已建] %s -> taskId=%s（这条今天第 %d 次尝试）"
        % (tname, tid, cs["retries"][batch[0]["key"]]["attempts"]))
    save_check_state(cfg, cs)
    return tid


def token_matches_plan(cfg, args, data, name=""):
    """带 --token 时必须和计划文件里的确认串一致，防止拿错计划乱发。"""
    tok = str(getattr(args, "token", "") or "")
    if not tok:
        return True
    want = str((data or {}).get("confirm_token") or "")
    if not want or tok == want:
        return True
    log("[不重发] %s：token 与计划里的确认串不一致" % name)
    return False


def follow_up_state(api, rec, cache):
    """脚本自己建的那条补发任务，现在到底是什么结果。

    原任务里那行失败记录不会自己变成功（客户端点重试才是原地重跑，脚本补发是新建任务），
    所以「已试过 N 次，交人工」绝不能只看原行——补发可能还在跑、也可能已经发出去了。
    返回 (判定, 说明)，判定同 classify_execution；没建过补发任务返回 (None, "")。
    """
    tid = str(rec.get("retryTaskId") or "")
    if not tid:
        return None, ""
    if tid not in cache:
        try:
            cache[tid] = api.executions(tid, size=50)
        except ApiError as e:
            cache[tid] = None
            cache["%s#err" % tid] = "%s" % e.res.get("code")
    rows = cache.get(tid)
    if rows is None:
        return None, "补发任务读不到（%s）" % (cache.get("%s#err" % tid) or "空")
    if not rows:
        return None, "补发任务里还没有执行明细"
    want = str(rec.get("mobileId") or "")
    best, best_at = None, None
    for r in rows:
        if str(r.get("mobileId")) != want:
            continue
        at = _ts(r.get("startedAt")) or _ts(r.get("createdAt")) or datetime.min
        if best is None or at >= best_at:
            best, best_at = r, at
    if best is None:
        return None, "补发任务里没有这台环境"
    return classify_execution(best, datetime.now())


def retry_pass(cfg, api, args, cs, items, now=None):
    """明确失败的条目 -> 算时刻 -> 建重发任务（后台等价客户端「重试」按钮）。

    2026-11-18 起的默认：明确失败当场重发，不再等「上次发布+30 分钟」；只有跟这台还没发的轮次
    撞间隔的才先进 cs["pending"]，由 checkd 到点再发（--retry-wait-gap 退回旧行为，
    --retry-now 连让位都不等）。结果不明被 --retry-ambiguous 放宽的那些仍旧等满间隔。
    每条最多 --max-retry 次，超过就只报不试，避免重试套重试。
    同一圈里同一条只排一次：一条视频在一个 pass 里最多烧掉一次重试额度。
    """
    now = now or datetime.now()
    max_retry = max(int(getattr(args, "max_retry", 1)), 0)
    immediate_ok = (not getattr(args, "retry_wait_gap", False)) \
        and bool(getattr(args, "retry_immediate", True))
    _path, sched = active_plan(cfg)
    cache, batches, deferred, blocked = {}, [], [], []
    fcache = {}
    queued, seen = {}, set()
    for it in items:
        key = it.get("key") or "%s#%s" % (it["taskId"], it["index"])
        seen.add(key)
        rec = (cs.get("retries") or {}).get(key) or {}
        tried = int(rec.get("attempts") or 0) + queued.get(key, 0)
        if tried >= max_retry:
            # 额度用完了，但补发任务自己可能还在跑、或者已经发出去了——先问它，别急着喊交人工。
            tag = str(rec.get("retryTaskName") or rec.get("retryTaskId") or "")
            kind, fwhy = (follow_up_state(api, rec, fcache) if rec.get("retryTaskId")
                          else (None, "没建过补发任务（上次建任务被接口打回了？）"))
            if kind in ("ok", "pending"):
                if rec.get("followReported") != kind:
                    rec["followReported"] = kind
                    log("[补发%s] %s（%s taskId=%s %s）—— 原任务里那行还是 FAILED，"
                        "别照着旧记录手工再发，等补发自己的结果"
                        % ("已发布" if kind == "ok" else "进行中", it["name"], tag,
                           rec.get("retryTaskId"), fwhy))
                continue
            blocked.append("%s（已试过 %d 次，上限 %d；补发 %s %s，交人工）"
                           % (it["name"], tried, max_retry, tag or "（无）",
                              {"retryable": "仍明确失败：%s" % fwhy,
                               "ambiguous": "结果不明：%s" % fwhy,
                               "cancelled": "已取消：%s" % fwhy}.get(
                                  kind, "状态未知：%s" % (fwhy or "查不到"))))
            continue
        data = load_plan_data(cfg, it["plan"], cache)
        if not data:
            blocked.append("%s（计划文件读不出：%s）" % (it["name"], it["plan"]))
            continue
        if not token_matches_plan(cfg, args, data, it["name"]):
            blocked.append("%s（token 对不上计划）" % it["name"])
            continue
        payload = data.get("payload") or {}
        src_items = payload.get("items") or []
        if not (0 <= int(it["index"]) < len(src_items)):
            blocked.append("%s（计划下标越界）" % it["name"])
            continue
        item = dict(src_items[it["index"]])
        problems = audit_content(cfg, item, it["name"])
        if problems:
            blocked.append("%s（重发前自检没过：%s）" % (it["name"], "；".join(problems[:2])))
            continue
        when, why_slot = retry_slot(cfg, args, cs, it["mobileId"], it.get("sentAt"), now, sched,
                                    immediate=(immediate_ok and not it.get("relaxed")))
        if when is None:
            blocked.append("%s（%s）" % (it["name"], why_slot))
            continue
        base = str(payload.get("taskName") or "新建任务")
        one = dict(it); one.update({"item": item, "key": key, "when": when,
                                    "base": base, "slotWhy": why_slot})
        if when > now + timedelta(seconds=60):
            queued[key] = queued.get(key, 0) + 1
            deferred.append(one)
            log("[顺延重发] %s（%s）排在 %s：%s"
                % (it["name"], it["mobile"], when.strftime(TRIGGER_FMT), why_slot))
            continue
        queued[key] = queued.get(key, 0) + 1
        batches.append(one)
    keep = []
    for p in (cs.get("pending") or []):
        if p.get("key") in seen:
            continue
        when = _ts(p.get("when"))
        if when and when > now:
            keep.append(p)
    cs["pending"] = [{"key": d["key"], "name": d["name"], "mobileId": d["mobileId"],
                      "mobile": d["mobile"], "index": d["index"], "plan": d["plan"],
                      "taskId": d["taskId"], "executionId": d["executionId"],
                      "sentAt": d.get("sentAt") or "", "why": d["why"], "base": d["base"],
                      "when": d["when"].strftime(TIME_FMT)} for d in deferred] + keep
    uniq = []
    for b in blocked:
        if b not in uniq:
            uniq.append(b)
    if uniq:
        log("[不重发] %d 条：%s" % (len(uniq), "；".join(uniq)))
    if deferred:
        log("[重发] %d 条已排进 pending，到点由 checkd 建任务（--retry-now 可立刻发）" % len(deferred))
    groups = {}
    for b in batches:
        groups.setdefault(b["base"], []).append(b)
    for base, grp in groups.items():
        rounds = group_into_rounds(grp, lambda x: x["mobileId"])
        for r, batch in enumerate(rounds):
            create_retry_tasks(cfg, api, args, cs, batch, now, base)
    save_check_state(cfg, cs)
    return len(batches), len(deferred), len(blocked)


def still_failed(cfg, api, tid, exid, since=None):
    """到点真要重发前，再看一眼原执行现在的状态。

    今天见过 RUNNING 带链接、几分钟后自己变 SUCCEEDED 的（示例产品A-01），
    顺延期间它可能已经发出去了，这时候再发就是同账号两条一样内容。
    since = 当初排这条 pending 时的发单时刻：接口里 startedAt 比它新，说明有人在客户端点了
    「重试/批量重试」原地重跑（2026-11-18 实测同 executionId、startedAt 被刷新），
    这时重跑自己要么在跑、要么又失败了——在跑就别抢，又失败照旧该发。
    """
    if not tid or not exid:
        return False, "pending 里没带原任务或原执行，不猜"
    try:
        rows = api.executions(tid, size=50)
    except ApiError as e:
        return False, "原任务读不到了（%s），当不成功也不当失败" % e.res.get("code")
    a = _ts(since)
    for r in rows:
        if str(r.get("executionId")) == str(exid):
            kind, why = classify_execution(r, datetime.now())
            b = _ts(r.get("startedAt"))
            if a and b and b > a:
                why = "客户端已重跑（%s -> %s）：%s" % (a.strftime("%H:%M:%S"),
                                                       b.strftime("%H:%M:%S"), why)
            if kind == "retryable":
                return True, why
            return False, "现在判定是 %s：%s" % (kind, why)
    return False, "原执行 %s 已经不在任务里了" % exid


def flush_pending(cfg, api, args, cs, now):
    """到点的 pending 拿出来重发；没到点的留着。"""
    due, keep = [], []
    for p in (cs.get("pending") or []):
        when = _ts(p.get("when"))
        if when is None:
            continue
        if when <= now + timedelta(seconds=60):
            due.append(p)
        else:
            keep.append(p)
    if not due:
        return 0
    cs["pending"] = keep
    items = []
    cache = {}
    for p in due:
        data = load_plan_data(cfg, p.get("plan"), cache)
        src = ((data or {}).get("payload") or {}).get("items") or []
        try:
            i = int(p.get("index"))                  # 0 是真下标，别被 `or -1` 当成缺省
        except (TypeError, ValueError):
            i = -1
        if not (0 <= i < len(src)):
            log("[丢弃] pending 里的 %s 计划下标越界，不发了" % p.get("name"))
            continue
        if not token_matches_plan(cfg, args, data, p.get("name")):
            log("[丢弃] pending 里的 %s token 和计划对不上，不发了" % p.get("name"))
            continue
        alive, why_now = still_failed(cfg, api, p.get("taskId"), p.get("executionId"),
                                      p.get("sentAt"))
        if not alive:
            log("[丢弃] pending 里的 %s 不用重发了：%s" % (p.get("name"), why_now))
            continue
        problems = audit_content(cfg, src[i], p.get("name"))
        if problems:
            log("[丢弃] pending 里的 %s 重发前自检没过（%s），不发了" % (p.get("name"), "；".join(problems[:2])))
            continue
        q = dict(p)
        q.update({"item": dict(src[i]), "key": p["key"], "when": now,
                  "base": p.get("base") or "新建任务",
                  "taskName": p.get("taskId"), "taskId": p.get("taskId"),
                  "executionId": p.get("executionId"), "why": p.get("why") or "到点重发"})
        items.append(q)
    if not items:
        save_check_state(cfg, cs)
        return 0
    rounds = group_into_rounds(items, lambda x: x["mobileId"])
    n = 0
    for batch in rounds:
        if create_retry_tasks(cfg, api, args, cs, batch, now, batch[0]["base"]):
            n += len(batch)
    save_check_state(cfg, cs)
    return n


def cmd_check(cfg, api, args):
    """一次性自检：查今天本脚本建的任务，明确失败的当场重发（只看不动加 --no-retry）。"""
    now = datetime.now()
    live = {}
    try:
        with io.open(_check_lock_path(cfg), encoding="utf-8") as f:
            live = json.load(f)
    except Exception:
        live = {}
    busy = pid_alive(live.get("pid")) and int(live.get("pid") or -1) != os.getpid()
    if busy:
        log("[只查不写] 自检守护进程（PID %s，%s 起）还在跑：这次结果不落盘，"
            "要看实时进度直接 tail 它的 check_*.log；两份进程各写各的会把同一条重发两遍，"
            "所以这次不重发——失败的重发交给那个守护进程，它本来就会当场重发。"
            % (live.get("pid"), live.get("startedAt")))
        args.no_retry = True
    items, stats, cs, _st = check_pass(cfg, api, args, now, reason="手动查", read_only=busy)
    log("[自检] 查了 %d 个任务 / %d 条执行：成功 %d、未定 %d、结果不明 %d、已取消 %d、明确失败 %d、"
        "读不到 %d、客户端已重跑 %d（本轮跳过 %d 个没到点或已终态的）"
        % (stats["tasks"], stats["items"], stats["ok"], stats["pending"], stats["ambiguous"],
           stats["cancelled"], stats["retryable"], stats["unreadable"], stats["rerun"],
           stats["skipped"]))
    if stats["ambiguous"]:
        log("[提醒] 结果不明 %d 条不自动重发（可能已经发出去了，重复发掉权重）。"
            "确认账号页没这条后，可以像客户端那样一次性补上：点任务详情页顶部「批量重试」，"
            "或后台跑 python tk_publish.py retry --task <taskId> --retry-ambiguous。"
            % stats["ambiguous"])
    n_pending = flush_pending(cfg, api, args, cs, now) if (args.retry and not args.no_retry) else 0
    if args.no_retry or not args.retry:
        if items:
            log("[待重发] 明确失败 %d 条：%s（现在只查不重发）"
                % (len(items), "、".join(i["name"] for i in items)))
        return
    if not items:
        if n_pending:
            log("[重发] 本次到点的 pending 已发 %d 条" % n_pending)
        return
    sent, deferred, blocked = retry_pass(cfg, api, args, cs, items, now)
    log("[重发] 立即建任务 %d 条、排进 pending %d 条、拦下 %d 条" % (sent, deferred, blocked))


def cmd_retry(cfg, api, args):
    """后台版「重试」按钮：把任务里现在失败的几条一次性重新发出去。

    客户端那两个按钮（2026-11-18 从 app.asar + 用户录屏核实）：
      顶部「批量重试」  -> /cloud-mobile/v1/cloud-mobile/task-main/retry   参数 {taskId}
                            弹「确定重试所有失败任务吗？不含取消的失败任务数：N」，原地重跑所有失败行
      行内「重试」      -> /cloud-mobile/v1/cloud-mobile/task-detail/retry 参数 {id: 内部行 id}
    两条都不在本地开放接口白名单（11-18 拿 10 种拼法直连 6873 全是 E010006；同一批探测里白名单内的
    task/cancel 回 E010008 参数校验失败，可见 E010006 就是没放行），行内那条连内部行 id
    开放接口也不回。P5 又不许点界面。所以这里用唯一放行的写路径 video-task/create 按原计划 payload
    重建一个发布任务，效果等同「把这条视频重新发一遍」，代价是多一个新 taskId、且要守本脚本的
    让位/截止规则。守护进程 checkd 在跑时默认让位（两个进程会把同一条发两遍）。
    """
    now = datetime.now()
    want = [x.strip() for x in str(args.task or "").split(",") if x.strip()]
    st = load_sched_state(cfg)
    cs = load_check_state(cfg)
    known = set([str(k) for k in (st.get("tasks") or {})] + [str(k) for k in (cs.get("retryTasks") or {})])
    for tid in want:
        if tid not in known:
            log("[不碰] %s：不在本脚本 state 记过的任务里。按 P2 这种只认 taskId 的活儿我不动，"
                "手工建的任务请在客户端任务详情页点「重试/批量重试」。" % tid)
    if not any(t in known for t in want) and want:
        return
    live = {}
    try:
        with io.open(_check_lock_path(cfg), encoding="utf-8") as f:
            live = json.load(f)
    except Exception:
        live = {}
    busy = pid_alive(live.get("pid")) and int(live.get("pid") or -1) != os.getpid()
    if busy and not args.force:
        log("[让位] 自检守护进程（PID %s，%s 起）还在跑，它本来就会当场重发明确失败；"
            "两个进程一起重发会把同一条发两遍。要替它动手就加 --force（同时建议先确认它卡住了）。"
            % (live.get("pid"), live.get("startedAt")))
        return
    args.force_check = True                     # 手动点重试就该立刻查，别等 15 分钟、也别认旧终态
    items, stats, cs, _st = check_pass(cfg, api, args, now, reason="手动重发",
                                       read_only=busy, only=want or None)
    for tid in (stats.get("not_found") or []):
        log("[查不到] %s：state 里没有这份任务记录，或已被删除；只重发本脚本建过的任务。" % tid)
    if not items:
        log("[重试] 没有「明确失败」的条目可重发（查了 %d 个任务 / %d 条执行：成功 %d、未定 %d、"
            "结果不明 %d、已取消 %d、客户端已重跑 %d）。"
            % (stats["tasks"], stats["items"], stats["ok"], stats["pending"],
               stats["ambiguous"], stats["cancelled"], stats["rerun"]))
        if stats["ambiguous"]:
            log("[重试] 结果不明的那几条要一起补发，加 --retry-ambiguous（先确认账号页没发出去过）。")
        return
    log("[重试] 待重发 %d 条：%s" % (len(items), "、".join(
        "%s(%s %s)" % (i["name"] or "?", str(i["mobile"]).split("-")[0], i["why"]) for i in items)))
    sent, deferred, blocked = retry_pass(cfg, api, args, cs, items, now)
    log("[重试] 当场建任务 %d 条、排进 pending %d 条、拦下 %d 条%s"
        % (sent, deferred, blocked, "（dry-run：没真建任务）" if args.dry_run else ""))


def cmd_checkd(cfg, api, args):
    """常驻自检：建完 --check-after 分钟查一次；下一轮发布前 --check-before 分钟再查上一轮；
    明确失败的当场重发（客户端点重试的节奏），要给后面轮次让位的排进 pending 到点再发；
    接口里发现同一条已被人在客户端重跑，就撤掉自己那条 pending，绝不双发。
    只写 tk_publish_check_state.json，不碰守护进程那份。"""
    p = acquire_check_lock(cfg, args)
    log("[自检守护] 起于 %s（PID %d）：建完 %d 分钟查一次，下一轮前 %d 分钟再查，"
        "明确失败自动重发=%s（当场重发=%s），结果不明自动重发=%s，单条重试上限 %d 次，"
        "客户端已重跑=认得（撤自己的重发，不双发）"
        % (datetime.now().strftime(TIME_FMT), os.getpid(), int(args.check_after),
           int(args.check_before), not args.no_retry,
           bool(not getattr(args, "retry_wait_gap", False)), bool(args.retry_ambiguous),
           int(args.max_retry)))
    stats = {"pending": 0, "retryable": 0, "unreadable": 0, "next_due": None}
    try:
        while True:
            now = datetime.now()
            try:
                items, stats, cs, _st = check_pass(cfg, api, args, now, reason="到点自检")
                flush_pending(cfg, api, args, cs, now)
                if not args.no_retry and items:
                    retry_pass(cfg, api, args, cs, items, now)
            except SystemExit:
                raise
            except Exception as e:
                log("[自检出错] %r（本地接口没通？客户端开着再等下一圈）" % e)
            csc = load_check_state(cfg)
            interval = timedelta(minutes=max(int(args.check_interval), 1))
            wake = now + interval
            _path, sched = active_plan(cfg)
            nxt = [s["when"] for s in sched if not s["done"] and s["when"] > now]
            if nxt:
                pre = min(nxt) - timedelta(minutes=int(args.check_before))
                if pre > now:
                    wake = min(wake, pre)               # 下一轮前 15 分钟：先校验上一轮
                else:
                    wake = min(wake, max(min(nxt) - timedelta(seconds=60),
                                         now + timedelta(seconds=30)))
            for q in (csc.get("pending") or []):
                w = _ts(q.get("when"))
                if w and w > now:
                    wake = min(wake, w)
            due = stats.get("next_due")
            if due:
                wake = min(wake, due)
            if not nxt and not (csc.get("pending") or []) and not due \
                    and not stats.get("pending") and not stats.get("retryable") \
                    and not stats.get("unreadable"):
                amb = int(stats.get("ambiguous") or 0)
                log("[自检退出] %s：今天建的任务后面没有待发轮次，也没有待重发，收工。%s"
                    % (datetime.now().strftime(TIME_FMT),
                       ("还有 %d 条「结果不明」没落定，得人工看一眼（确认账号页没这条后，"
                        "客户端点「批量重试」或跑 python tk_publish.py retry --task <taskId> "
                        "--retry-ambiguous）；要我继续盯就重起 checkd" % amb) if amb else
                       "本轮查到的都已落定"))
                break
            wake = max(wake, now + timedelta(seconds=30))
            secs = max(int((wake - datetime.now()).total_seconds()), 30)
            log("[自检等待] 下次 %s（%d 秒后）" % (wake.strftime("%H:%M"), secs))
            while secs > 0:
                time.sleep(min(secs, 20))
                secs -= 20
    finally:
        try:
            if int((json.load(io.open(p, encoding="utf-8")).get("pid") or -1)) == os.getpid():
                os.remove(p)
        except Exception:
            pass


def acquire_check_lock(cfg, args):
    """自检进程单实例锁：两个进程各查各的会把同一条重发两遍。"""
    p = _check_lock_path(cfg)
    try:
        with io.open(p, encoding="utf-8") as f:
            holder = json.load(f)
    except Exception:
        holder = {}
    if pid_alive(holder.get("pid")) and int(holder.get("pid") or -1) != os.getpid():
        die("已经有一个自检进程在跑（PID %s，%s 起）。两个进程会把同一条失败视频重发两遍，"
            "先人工停掉旧的。" % (holder.get("pid"), holder.get("startedAt")))
    me = {"pid": os.getpid(), "startedAt": datetime.now().strftime(TIME_FMT),
          "retry": not getattr(args, "no_retry", False)}
    for _ in range(5):
        try:
            fd = os.open(p, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            with io.open(fd, "w", encoding="utf-8") as fh:
                json.dump(me, fh, ensure_ascii=False, indent=2)
            return p
        except OSError:
            try:
                with io.open(p, encoding="utf-8") as f:
                    cur = json.load(f)
            except Exception:
                cur = {}
            if pid_alive(cur.get("pid")) and int(cur.get("pid") or -1) != os.getpid():
                die("锁里已有自检进程（PID %s），不再起第二个" % cur.get("pid"))
            try:
                os.remove(p)
            except OSError:
                break
            time.sleep(0.3)
    log("[警告] 自检锁没抢到也没写下去，启动前自己确认只有一个进程")
    return p


def cmd_retime(cfg, api, args):
    """批量修改发布时间：把已有粘贴包的「发布时间」列按向导对话框的原式整列重排。

    客户端源码（app.asar module 13504）里那个对话框的算法就一行：
        triggerTime = 起始时间 + 组内视频序号 * 单手机多视频间隔 + 手机分组序号 * 多手机间隔
    本函数逐字复刻它，所以「脚本算好的时间」和「你在向导里点一次批量设置」得到的是同一张表。
    只改时间列，内容列（文案/小黄车/评论）一律不动；默认另存新文件。
    """
    import openpyxl
    path = _resolve_pack_file(cfg, args)
    wb = openpyxl.load_workbook(path)
    ws, hdr = _find_paste_sheet(wb)
    c_env, c_name, c_time = hdr["账号环境"], hdr["素材名称"], hdr["发布时间"]

    min_spacing = check_spacing("video", round_floor_minutes(cfg["publish"]))
    deadline_hm = parse_deadline(cfg["publish"].get("day_last_deadline", DEFAULT_DAY_DEADLINE))
    auto_spacing = bool(cfg["publish"].get("auto_spacing", True)) and not int(args.spacing or 0)
    spacing = check_spacing("video", args.spacing or cfg["publish"].get("spacing_minutes", min_spacing))
    mobile_spacing = check_spacing("mobile", args.mobile_spacing
                                   or cfg["publish"].get("mobile_spacing_minutes",
                                                          DEFAULT_MOBILE_SPACING_MIN))

    rows = []
    for r in range(2, ws.max_row + 1):
        nm = ws.cell(r, c_name).value
        if not nm or not str(nm).strip():
            continue
        env = str(ws.cell(r, c_env).value or "").strip() or "(空账号环境)"
        old = str(ws.cell(r, c_time).value or "").strip()
        rows.append({"r": r, "name": str(nm).strip(), "env": env, "old": old})
    if not rows:
        die("%s 的「%s」里没有数据行" % (os.path.basename(path), ws.title))

    order = []
    for x in rows:
        if x["env"] not in order:
            order.append(x["env"])
    idx = {e: i for i, e in enumerate(order)}
    cnt = {}

    start = parse_when(args.start_time)
    if start is None:
        olds = []
        for x in rows:
            try:
                olds.append(datetime.strptime(x["old"], TRIGGER_FMT))
            except ValueError:
                pass
        first_env = rows[0]["env"]
        same = [datetime.strptime(x["old"], TRIGGER_FMT) for x in rows
                if x["env"] == first_env and x["old"]]
        start = (min(same) if same else (min(olds) if olds else
                 datetime.now().replace(second=0, microsecond=0) + timedelta(minutes=1)))
        log("[起始时间] 没给 --start-time，沿用原表里 %s 的第一条 %s"
            % (first_env, start.strftime(TRIGGER_FMT)))
    start = start.replace(second=0, microsecond=0)

    per_env_total = {}
    for x in rows:
        per_env_total[x["env"]] = per_env_total.get(x["env"], 0) + 1
    busiest = max(per_env_total.values() or [1])
    if auto_spacing:
        spacing = spread_spacing(start, deadline_hm, busiest, min_spacing)[0]
        log("[自动间隔] 单环境最多 %d 条，%s 起、当天最晚 %02d:%02d ⇒ 每条 %d 分钟，单天每台 %d 条"
            % (busiest, start.strftime(TRIGGER_FMT), deadline_hm[0], deadline_hm[1],
               spacing, day_capacity(start, deadline_hm, spacing)))

    floor = parse_when(args.after)
    changed = blank = 0
    for x in rows:
        k = cnt.get(x["env"], 0)
        cnt[x["env"]] = k + 1
        t = wizard_time_spread(start, idx[x["env"]], k, mobile_spacing, spacing, deadline_hm)
        if floor and t < floor:
            t = wizard_time_spread(floor, idx[x["env"]], k, mobile_spacing, spacing, deadline_hm)
        txt = t.strftime(TRIGGER_FMT)
        if not x["old"]:
            blank += 1
        elif x["old"] != txt:
            changed += 1
        ws.cell(x["r"], c_time).value = txt
        x["new"] = txt

    # 「核对」表里同名的时间列一起改，免得两张表对不上
    fixed_other = 0
    for name in wb.sheetnames:
        if name == ws.title:
            continue
        s2 = wb[name]
        h2 = {str(s2.cell(1, c).value or "").strip(): c for c in range(1, s2.max_column + 1)}
        if "发布时间" not in h2 or "素材名称" not in h2:
            continue
        m = {x["name"]: x["new"] for x in rows}
        for r in range(2, s2.max_row + 1):
            nm = str(s2.cell(r, h2["素材名称"]).value or "").strip()
            if nm in m:
                s2.cell(r, h2["发布时间"]).value = m[nm]
                fixed_other += 1

    out_path = path
    if not args.in_place:
        stem, ext = os.path.splitext(path)
        out_path = os.path.join(cfg["outputs_dir"], args.out or (os.path.basename(stem) + "_改时" + ext)) \
            if not args.out else (args.out if os.path.isabs(args.out)
                                  else os.path.join(cfg["outputs_dir"], args.out))
    if args.in_place and os.path.dirname(os.path.abspath(path)) != os.path.abspath(cfg["outputs_dir"]):
        die("--in-place 只允许覆盖 outputs_dir 里的文件（P3：不往素材目录写），去掉这个参数改另存")
    wb.save(out_path)

    log("[批量修改发布时间] %s -> %s" % (path, out_path))
    log("[算法] 起始 %s + 组内序号*%d 分钟 + 分组序号(%s) *%d 分钟，共 %d 行 / %d 个账号环境"
        % (start.strftime(TRIGGER_FMT), spacing, " / ".join("%s=%d" % (e, idx[e]) for e in order),
           mobile_spacing, len(rows), len(order)))
    log("[结果] 改了 %d 行时间，原本空着 %d 行；「核对」同步 %d 行" % (changed, blank, fixed_other))
    for e in order:
        s = [x for x in rows if x["env"] == e]
        ts = [datetime.strptime(x["new"], TRIGGER_FMT) for x in s]
        gaps = sorted(int((ts[i + 1] - ts[i]).total_seconds() // 60) for i in range(len(ts) - 1))
        night = sum(1 for t in ts if t.hour < 7)
        log("        %s：%s -> %s（%d 条 / 同环境最小间隔 %s 分钟%s）"
            % (e, s[0]["new"], s[-1]["new"], len(s), gaps[0] if gaps else "-",
               "，其中 %d 条落在 00:00-07:00 凌晨" % night if night else ""))
    allt = sorted(datetime.strptime(x["new"], TRIGGER_FMT) for x in rows)
    if len(set(x["new"] for x in rows)) != len(rows):
        log("[提醒] 有重复时刻：多手机间隔 %d 分钟对同一账号环境内的 %d 分钟间隔没能错开，"
            "请把 --mobile-spacing 调到能整除的值或更大" % (mobile_spacing, spacing))
    log("[跨度] %s -> %s（%.1f 小时 / %d 天）" % (allt[0].strftime(TRIGGER_FMT),
            allt[-1].strftime(TRIGGER_FMT),
            (allt[-1] - allt[0]).total_seconds() / 3600.0, len({t.date() for t in allt})))
    for x in rows[:12]:
        log("        %-14s %-19s -> %s" % (x["name"], x["old"] or "(空)", x["new"]))
    if len(rows) > 12:
        log("        …其余 %d 行见文件" % (len(rows) - 12))
    log("[向导对照] 想在客户端一次点出来：第 4 步「批量设置」填 起始时间 %s / 多手机间隔 %d / "
        "单手机多视频间隔 %d / 忙时等待按原列" % (start.strftime(TRIGGER_FMT), mobile_spacing, spacing))
    log("[提醒] 这些时间只在客户端向导里生效；video-task/create 不解析时间字段，脚本提交是建完即发。")


def _sched_state_path(cfg):
    return os.path.join(cfg["outputs_dir"], "tk_publish_schedule_state.json")


def load_sched_state(cfg):
    p = _sched_state_path(cfg)
    if os.path.exists(p):
        try:
            return json.load(io.open(p, encoding="utf-8"))
        except ValueError:
            pass
    return {"fired": {}}


def sent_names_today(cfg, plan_key=""):
    """读盘上状态：今天「已经真发出去」的视频名 -> (计划文件, 时刻, taskId)。

    plan_key 传自己的计划名就排除自己，用来给守护进程做跨计划避让：
    两个人（比如 14:30 的定时任务和人工接管）几乎同时起了守护进程时，
    同一条视频只让先醒的那个发，后醒的直接跳过，不会重复发布掉权重。
    """
    today = datetime.now().strftime("%Y-%m-%d")
    out = {}
    st = load_sched_state(cfg)
    for pk, entries in (st.get("fired") or {}).items():
        if plan_key and os.path.basename(str(pk)) == plan_key:
            continue
        pp = resolve_plan_file(cfg, pk)
        if not pp:
            continue
        try:
            det = (json.load(io.open(pp, encoding="utf-8")).get("detail")) or []
        except Exception:
            continue
        for k, v in (entries or {}).items():
            if not v or not v.get("taskId") or str(v.get("at") or "")[:10] != today:
                continue
            i = fired_index(k)
            if 0 <= i < len(det) and det[i].get("name"):
                out[str(det[i]["name"]).lower()] = (os.path.basename(pp), v.get("at"), v.get("taskId"))
    for nm, v in (load_state(cfg).get("published") or {}).items():
        if isinstance(v, dict) and str(v.get("at") or "")[:10] == today:
            out.setdefault(str(nm).lower(), ("提交台账", v.get("at"), v.get("taskId") or ""))
    # 任务记录这份不参与 plan_key 排除：它记的是「真发出去过」，
    # 自己发过的本来就不该再发，跨计划就更该让路。
    for nm, v in task_record_names(cfg, today).items():
        out.setdefault(nm, v)
    return out


def plan_is_stale(cfg, detail):
    """计划文件是否比现在的配置旧：配置里后来补了商品标题/首评，而计划里还是空的。

    守护进程只认 plan 落盘时算好的 payload，不会回头读配置。所以配置一更新，
    旧计划必须作废重跑，否则发出去的车里没标题、首评退回文案里那句。
    """
    hits = []
    for d in detail or []:
        nm = str(d.get("name") or "")
        stem = re.sub(r"\.(mp4|mov|m4v|avi|mkv|webm)$", "", nm, flags=re.I)
        prod = match_product(cfg, nm)
        want_t = by_video((prod or {}).get("title_by_video"), nm) or (prod or {}).get("productTitle", "")
        if want_t and not str(d.get("productTitle") or ""):
            hits.append("%s 缺商品标题" % stem)
        want_c = by_video(cfg.get("comment_by_video"), nm)
        if want_c and not str(d.get("comment") or ""):
            hits.append("%s 缺首评" % stem)
        elif want_c and str(d.get("comment_src") or "") != "comment_by_video" \
                and str(d.get("comment") or "") != want_c:
            hits.append("%s 首评是旧来源(%s)" % (stem, d.get("comment_src")))
    return hits


def save_sched_state(cfg, st):
    """写盘前把盘上那份并进来：别的进程记的「真发出去」不能被本地覆盖。

    只往 st 自己的 dict 里就地合并，绝不把 st 换成 disk 那套新对象。
    守护进程把 fired 这种局部引用绑在 st["fired"][plan_key] 上，一换对象，
    它后面的写入就落到没人再看的孤儿字典里：2026-11-17 就是因此丢了 16:42
    那一轮的 fired（tasks 写进去了、fired 没写），重启会照着 fired 把那 5 条重发一遍。
    """
    p = _sched_state_path(cfg)
    disk = load_sched_state(cfg)
    lf = st.setdefault("fired", {})
    for plan_key, entries in (disk.get("fired") or {}).items():
        mine = lf.setdefault(plan_key, {})
        for k, v in (entries or {}).items():
            old = mine.get(k)
            if old and old.get("taskId") and not (v or {}).get("taskId"):
                continue          # 本地记的是真发出去的，别让盘上的空值盖掉
            if k not in mine:
                mine[k] = v       # 只补本地没有的（别的计划/别的进程记的）
    lt = st.setdefault("tasks", {})
    for tid, v in (disk.get("tasks") or {}).items():
        lt.setdefault(tid, v)
    tmp = p + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def pid_alive(pid):
    """Windows 下只读地判断进程还在不在。

    绝不能用 os.kill(pid, 0)「探活」：CPython 在 Windows 上会把普通信号值转成
    TerminateProcess(handle, code)，等于把那个进程杀掉。
    """
    try:
        pid = int(pid or 0)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    try:
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.windll.kernel32
        STILL_ACTIVE = 259
        h = k32.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = wintypes.DWORD()
        ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
        k32.CloseHandle(h)
        return bool(ok) and code.value == STILL_ACTIVE
    except Exception:
        return False


def _sched_lock_path(cfg):
    return os.path.join(cfg["outputs_dir"], "tk_publish_schedule.lock")


def acquire_sched_lock(cfg, plan_key, args):
    """守护进程单实例锁（2026-11-17 加）：同一份计划被两个进程各跑一遍 = 间隔规则作废。

    那天手工 launch_schedule.ps1 起了两个 PID（14020、6364）跑同一个令牌，
    才发现这条路径完全没有防护。锁文件只记 PID/计划/令牌，进程死了锁自然失效，不会永久卡住。
    """
    p = _sched_lock_path(cfg)
    try:
        holder = json.load(io.open(p, encoding="utf-8"))
    except Exception:
        holder = {}
    pid = holder.get("pid")
    same_plan = str(holder.get("plan") or "") == str(plan_key)
    if pid_alive(pid) and same_plan and int(pid) != os.getpid():
        die("已经有守护进程在跑同一份计划：%s（PID %s，令牌 %s，%s 起）。"
            "同一份计划开两个进程会把每条视频发两遍，先人工停掉旧的再启动。"
            % (plan_key, pid, holder.get("token"), holder.get("startedAt")))
    me = {"pid": os.getpid(), "plan": plan_key, "token": getattr(args, "token", ""),
          "startedAt": datetime.now().strftime(TIME_FMT)}
    # 必须用 O_EXCL 抢锁：先读再覆盖写的写法，两个进程同一秒起会双双放行（15:29 那次就是这样）。
    acquired = False
    for _ in range(5):
        try:
            fd = os.open(p, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            with io.open(fd, "w", encoding="utf-8") as fh:
                json.dump(me, fh, ensure_ascii=False, indent=2)
            acquired = True
            break
        except OSError:
            try:
                cur = json.load(io.open(p, encoding="utf-8"))
            except Exception:
                cur = {}
            other = cur.get("pid")
            if pid_alive(other) and int(other or -1) != os.getpid():
                if str(cur.get("plan") or "") == str(plan_key):
                    die("锁里已有跑同一份计划的守护进程（PID %s，%s 起，令牌 %s），不再起第二个"
                        % (other, cur.get("startedAt"), cur.get("token")))
                log("[锁] 上一个持有者 PID %s 跑的是另一份计划（%s），两份计划的视频靠状态文件互相避让"
                    % (other, cur.get("plan")))
                acquired = True
                break
            try:
                os.remove(p)      # 持有者早死了，这把锁是垃圾
            except OSError:
                break
            time.sleep(0.3)
    if not acquired:
        log("[警告] 守护进程锁没抢到也没写下去，这次挡不住第二个进程，启动前自己确认只有一个")
    return p


def task_name_guard(api):
    """任务名避让：开放接口对重名任务直接 E114208 整条打回，视频等于没发。

    2026-11-17 15:30 那次守护进程按 plan 落盘时的固定序号命名（新建任务2002-1/-2），
    撞上 14:41 那批已建的同名任务，第一条 示例产品A-01 就被拒了。所以名字必须现取现让：
    任务列表每 60 秒刷一次，本进程内取过的名字先占住。
    """
    cache = {"ts": 0.0, "taken": set()}

    def next_name(prefix, prefer):
        if time.time() - cache["ts"] > 60:
            try:
                rows = api.tasks(200) or []
                cache["taken"] = {str(r.get("taskName") or "") for r in rows}
            except Exception as e:
                log("[警告] 取任务列表失败（%s），任务名只能避开本次运行内已取的" % e)
            cache["ts"] = time.time()
        n = max(1, int(prefer))
        while True:
            cand = "%s%d" % (prefix, n)
            if cand not in cache["taken"]:
                cache["taken"].add(cand)
                return cand, n
            n += 1

    return next_name


def cmd_schedule(cfg, api, args):
    """后台定时发布：脚本自己等到点，再逐条调 video-task/create。全程不碰客户端界面。

    为什么必须自己等：本地开放接口白名单（app.asar 里那张 proxy map，2026-11-17 重新逐条确认）
    只有 video-task/create 能建发视频任务，而它不解析任何时间字段；客户端真正用来排期的
    task-main/save/v2、task-main/batch-update-trigger-time、task-detail/update-trigger-time
    都不在白名单里（直连一律 E010006）。所以「定时」这件事只能由进程等到点再发，
    发出去的时刻 = 到点那一刻，和 plan 算出来的 triggerTime 对齐。

    安全：必须带 plan 的 --token；已发过的条目记进 tk_publish_schedule_state.json，
    重启不会重发；超过 --catchup 分钟才醒来的条目判 MISSED 不补发（避免一醒来就连环发）。
    """
    if not args.token:
        die("schedule 必须带 --token（先运行 plan 拿令牌）")
    path = (resolve_plan_file(cfg, args.file) if args.file else find_latest_plan(cfg))
    if not path:
        die("没有找到计划文件，请先运行 plan")
    data = json.load(io.open(path, encoding="utf-8"))
    if data.get("confirm_token") != args.token:
        die("token 与计划文件不一致（%s != %s，%s）" % (args.token, data.get("confirm_token"), path))
    payload = data["payload"]
    items = payload.get("items") or []
    detail = data.get("detail") or []
    if not items:
        die("计划文件里没有 items（%s），重跑 plan" % path)
    if len(detail) != len(items):
        die("计划文件的 detail(%d) 与 items(%d) 条数不一致，不敢猜时刻，重跑 plan"
            % (len(detail), len(items)))

    plan_key = os.path.basename(path)
    stale = plan_is_stale(cfg, detail)
    if stale:
        die("计划文件 %s 比当前配置旧（%s，共 %d 处）：配置里补过商品标题/首评就必须重跑 plan，"
            "旧令牌不要拿来启动守护进程，否则发出去的小黄车没标题、首评还是旧来源。"
            % (plan_key, "；".join(stale[:3]), len(stale)))
    st = load_sched_state(cfg)
    fired = st["fired"].setdefault(plan_key, {})
    base_name = payload.get("taskName") or "新建任务"
    pick_name = task_name_guard(api)
    catchup = max(int(args.catchup), 0)
    # 排期行与轮次切法都抽给 build_schedule_rows 了（2026-11-17 收拢）：
    # 守护进程和 watchdog 读的是同一份「还剩哪些没发、下一轮几点」，不会各算各的。
    # 轮次横向打包的依据不变：一个任务里不许同一台环境出现两次（E010008），
    # 但「每台环境各一条」装进同一个任务是允许的——14:54 的 新建任务2002-R1 就是 5 条一任务建成、5 条全成功。
    sched, by_round = build_schedule_rows(plan_key, items, detail, fired)
    rounds = len(by_round)
    lanes = len({s["item"].get("mobileId") for s in sched})
    batch = bool(getattr(args, "round_batch", True)) and lanes > 1 and rounds > 1

    log("[定时发布] 计划 %s / %d 条 / 迟到宽限 %d 分钟 / 令牌 %s"
        % (plan_key, len(sched), catchup, args.token))
    log("[打包] %s" % ("按轮次横向打包：%d 轮 / %d 个任务（每台环境每轮最多 1 条，同轮各环境并行执行）"
                       % (rounds, rounds) if batch else
                       "一条一任务：%d 个任务（--one-per-task 或只有一台环境）" % len(sched)))
    def label(s):
        d = s["done"]
        if not d:
            return "待发"
        if d.get("taskId"):
            return "已发 taskId=%s" % d["taskId"]
        if d.get("skip") == "dup":
            return "已避让(%s)" % d.get("other", "别处发过")
        return "已跳过(MISSED)" if d.get("missed") else "已试过但失败 %s" % (d.get("error") or "?")

    for s in sched:
        log("        %s  R%-2d %-14s %-19s %s" % (s["when"].strftime(TRIGGER_FMT), s["round"] + 1,
                                                  s["detail"].get("mobile", ""),
                                                  s["detail"].get("name", ""), label(s)))
    if args.dry_run:
        log("[dry-run] 只打印时间线，不发任何任务。")
        return
    todo = [s for s in sched if not s["done"]]
    if not todo:
        log("[定时发布] 这个计划的条目状态里都已处理过（发过/跳过/试过失败），什么都不做。要重来请重跑 plan 换新文件。")
        return
    acquire_sched_lock(cfg, plan_key, args)

    while True:
        now = datetime.now()
        pending = [s for s in todo if not s["done"]]
        if not pending:
            break
        nxt = min(pending, key=lambda s: s["when"])
        group = sorted([s for s in pending if s["round"] == nxt["round"]], key=lambda s: s["when"]) \
            if batch else [nxt]
        head = group[0]
        if now > head["when"] + timedelta(minutes=catchup):
            log("[跳过] %s %s 计划时刻 %s 已过宽限 %d 分钟，判 MISSED 不补发"
                % (head["when"].strftime(TRIGGER_FMT), head["detail"].get("name"),
                   head["when"].strftime(TRIGGER_FMT), catchup))
            head["done"] = {"taskId": "", "missed": True, "at": now.strftime(TIME_FMT)}
            fired[str(head["i"])] = head["done"]
            save_sched_state(cfg, st)
            continue
        if now < head["when"]:
            wait = int((head["when"] - now).total_seconds()) + 1
            log("[等待] 下一批 %s %s（第 %d 轮共 %d 条），还要等 %d 秒（进程保持前台终端可用，Ctrl+C 可停）"
                % (head["when"].strftime(TRIGGER_FMT), head["detail"].get("name"),
                   head["round"] + 1, len(group), wait))
            while wait > 0:
                time.sleep(min(wait, 20))
                wait -= 20
            continue
        clash_map = sent_names_today(cfg)
        keep = []
        for s in group:
            nm = str(s["detail"].get("name") or "").lower()
            clash = clash_map.get(nm)
            if clash:
                log("[避让] %s 今天已由 %s 在 %s 发出（taskId=%s），这条不再发第二次"
                    % (s["detail"].get("name"), clash[0], clash[1], clash[2]))
                s["done"] = {"taskId": "", "skip": "dup", "other": "%s@%s" % clash[:2],
                             "at": datetime.now().strftime(TIME_FMT)}
                fired[str(s["i"])] = s["done"]
                continue
            keep.append(s)
        if len(keep) != len(group):
            save_sched_state(cfg, st)
        if not keep:
            continue
        # 同一个任务里 mobileId 不许重复：真撞上了每台只留第一条，剩下的下一圈自己当 head
        seen, fire = set(), []
        for s in keep:
            mid = s["item"].get("mobileId")
            if mid in seen:
                continue
            seen.add(mid)
            fire.append(s)
        fire = fire[:MAX_VIDEOS_PER_MOBILE]
        prefix = "%s-R" % base_name if batch else "%s-" % base_name
        want = "%s%d" % (prefix, fire[0]["round"] + 1 if batch else fire[0]["i"] + 1)
        tname, _n = pick_name(prefix, int(want[len(prefix):]))
        if tname != want:
            log("[改名] 默认任务名 %s 已被占用（E114208 会把整批打回），改用 %s" % (want, tname))
        one = {"taskName": tname,
               "queueWhenBusy": payload.get("queueWhenBusy", True),
               "items": [s["item"] for s in fire]}
        who = "、".join("%s:%s" % (str(s["detail"].get("mobile") or "").split("-")[0],
                                   s["detail"].get("name")) for s in fire)
        try:
            res = api.create_video_task(one)
        except ApiError as e:
            code = str(e.res.get("code"))
            log("[被拒] %s（%d 条：%s）-> code=%s msg=%s requestId=%s"
                % (tname, len(fire), who, code, e.res.get("msg"), e.res.get("requestId")))
            if code == "LOCAL":
                # 根本没连上本地接口 = 服务端没建任务，重试不会重复发布；客户端可能只是暂时没开。
                fire[0]["tries"] = int(fire[0].get("tries") or 0) + 1
                if fire[0]["tries"] < int(args.max_retries):
                    log("[重试] 本地接口没通（客户端没开？），%d/%d 次，等 60 秒再来 %s"
                        % (fire[0]["tries"], int(args.max_retries), fire[0]["detail"].get("name")))
                    time.sleep(60)
                    continue
                log("[放弃] 连续 %d 次连不上本地接口，这一批判 FAILED；确认客户端开着再重跑本命令"
                    % fire[0]["tries"])
            for s in fire:
                s["done"] = {"taskId": "", "error": code, "at": datetime.now().strftime(TIME_FMT)}
                fired[str(s["i"])] = s["done"]
            save_sched_state(cfg, st)
            continue
        tid = res.get("taskId") or (res.get("task") or {}).get("taskId") or ""
        log("[已发] %s 第%d轮 %d 条（%s）-> taskId=%s"
            % (head["when"].strftime(TRIGGER_FMT), fire[0]["round"] + 1, len(fire), who, tid))
        for s in fire:
            s["done"] = {"taskId": tid, "at": datetime.now().strftime(TIME_FMT)}
            fired[str(s["i"])] = s["done"]
        st.setdefault("tasks", {})[tid] = {"taskName": tname,
                                           "createdAt": datetime.now().strftime(TIME_FMT),
                                           "items": len(fire), "plan": path,
                                           "scheduledFor": head["when"].strftime(TRIGGER_FMT),
                                           "round": fire[0]["round"] + 1,
                                           "mobiles": sorted({str(s["detail"].get("mobile") or "") for s in fire}),
                                           "videos": [s["detail"].get("name") for s in fire]}
        save_sched_state(cfg, st)

    ok = sum(1 for s in sched if (s["done"] or {}).get("taskId"))
    dup = [s for s in sched if (s["done"] or {}).get("skip") == "dup"]
    bad = [s for s in sched if s["done"] and not s["done"].get("taskId")
           and s["done"].get("skip") != "dup"]
    log("[定时发布] 完成：发出 %d 条，避让别处已发 %d 条，未发/失败 %d 条%s"
        % (ok, len(dup), len(bad),
           "（" + "，".join(s["detail"].get("name", "?") for s in bad) + "）" if bad else ""))
    log("[核对] python tk_publish.py status --size 8 看任务；跑完 watch --once 出运行结果，再 ledger 标台账")


def main():
    ap = argparse.ArgumentParser(description="Hubstudio 批量发视频 后台自动上架脚本")
    ap.add_argument("command", choices=["plan", "submit", "status", "watch", "phones", "ledger",
                                        "cancel", "pack", "paste", "verify", "watchdog", "retime", "schedule",
                                        "today", "check", "checkd", "retry"])
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                     "tk_publish.config.json"))
    ap.add_argument("--token", default="")
    ap.add_argument("--retry-export", default="", help="Hubstudio 导出的运行结果 xlsx，只重发其中失败的行")
    ap.add_argument("--include-reason", default="", help="只重发这些失败原因，逗号分隔，如 商品添加失败,取消执行")
    ap.add_argument("--limit", type=int, default=0, help="只取前 N 条，便于小批量试跑")
    ap.add_argument("--only", default="", help="只发视频名包含这些关键字的行，逗号分隔，如 示例产品A-01,示例产品A-02")
    ap.add_argument("--exclude", default="",
                    help="反着筛：视频名/路径含这些关键字的行剔掉不发，逗号分隔。"
                         "台账状态滞后、导出表里又没有今天的结果时，用它挡掉已经发成功的那几条（重复发布很伤）")
    ap.add_argument("--seq", default="", help="只发到这些「序列」号的账号环境（UI 的序列=接口 number），逗号分隔；留空用 config.phones")
    ap.add_argument("--alloc", action="store_true",
                    help="plan/pack：按 phones[].weight 的档位权重分配这批视频，而不是平均轮转"
                         "（等价于 config.publish.alloc.enabled=true）")
    ap.add_argument("--max-day", dest="max_day", type=int, default=0,
                    help="配合 --alloc：单账号环境「今日累计（新增+今日已发）」上限，超出就往权重次高的号挪；0=不限制")
    ap.add_argument("--task", default="", help="watch/cancel/retry 指定 taskId（retry 可逗号给多个；留空=今天本脚本建的全部任务）")
    ap.add_argument("--force", action="store_true",
                    help="retry：自检守护进程还在跑也要自己动手（先确认它真的卡住了，否则两个进程会把同一条发两遍）")
    ap.add_argument("--interval", type=int, default=120, help="watch 轮询秒数")
    ap.add_argument("--once", action="store_true", help="watch 只跑一轮")
    ap.add_argument("--size", type=int, default=50, help="status 显示条数")
    ap.add_argument("--skip-url-check", action="store_true", help="plan 时不做 HTTP 可达性检查")
    ap.add_argument("--allow-republish", action="store_true",
                    help="允许重发导出表里已成功的视频（默认剔除，防重复发布）")
    ap.add_argument("--no-upload", action="store_true",
                    help="plan 时不上传视频，只复用已有缓存（严格只读）")
    ap.add_argument("--no-out", dest="no_out", action="store_true",
                    help="plan：只出计划与自检，不写 tk_publish_plan_*.json、不覆盖「计划与校验」表、不发令牌"
                         "（排查用；配 --no-upload 就是全程零写入。要真发必须去掉它重跑取新令牌）")
    ap.add_argument("--keep-plans", dest="keep_plans", type=int, default=3,
                    help="plan：落盘后把更早的计划 json 移进 outputs/plans/，原地保留几份（默认 3；守护进程锁里那份永远不动）")
    ap.add_argument("--url-prefix", default="",
                    help="临时覆盖 config.video_source：按该 CDN/OSS 前缀把本地路径换成 URL")
    ap.add_argument("--single-time", default="", help="排期起点（等价 --start-time，旧参数名保留）")
    ap.add_argument("--start-time", default="", help="排期起点/向导「起始时间」：2026-11-17 20:00 / 11-17 20:00 / 20:00；留空按文案【四】窗口起点（wizard 模式留空则取下一整分钟）")
    ap.add_argument("--after", default="", help="这批第一条至少排在该时刻之后一个间隔（第 2 批接第 1 批时用）")
    ap.add_argument("--spacing", type=int, default=0,
                    help="单手机多视频间隔（分钟，向导允许 10~1440），默认取 config.spacing_minutes=60")
    ap.add_argument("--mobile-spacing", type=int, default=0,
                    help="多手机间隔（分钟，向导允许 5~1440），默认取 config=5；--schedule wizard 时生效")
    ap.add_argument("--schedule", default="", choices=["", "caption_window", "window_daily", "fixed",
                                                       "immediate", "wizard"],
                    help="排期模式：wizard=逐字复刻向导「批量设置发布时间」的算法（推荐，和客户端算出来一模一样）；"
                         "caption_window=从窗口起点连续排（会排进凌晨）；window_daily=只在窗口内排，排满顺延次日")
    ap.add_argument("--upload-workers", type=int, default=4, help="批量上传并发路数（1~8，默认 4）")
    ap.add_argument("--pack-out", default="", help="pack/paste：粘贴包文件名（落在 outputs_dir）")
    ap.add_argument("--lead", type=int, default=3,
                    help="paste：整批后移时，最早一条距此刻至少留几分钟（默认 3，给向导点确定的时间）")
    ap.add_argument("--file", default="", help="retime：要改时的粘贴包 xlsx（留空取 outputs 里最新的 *粘贴包*.xlsx）；"
                                              "schedule/verify/watchdog：计划 json（留空取最新，watchdog 留空则优先用锁里那份）")
    ap.add_argument("--catchup", type=int, default=CATCHUP_DEFAULT_MIN,
                    help="schedule：进程迟醒多久之内仍补发这条（分钟，默认 10；超过判 MISSED 不连环补发）")
    ap.add_argument("--one-per-task", dest="round_batch", action="store_false", default=True,
                    help="schedule：一条一个任务（旧行为）。默认按轮次横向打包：一个任务 = 每台环境各一条，"
                         "同一环境仍隔 --spacing 分钟。开放接口不许同一任务里 mobileId 重复，跨手机打包是允许的"
                         "（2026-11-17 实测 新建任务2002-R1 = 5 台各 1 条，5 条全成功）")
    ap.add_argument("--dry-run", action="store_true",
                    help="schedule/checkd/check：只打印将要做的，不建任何任务")
    ap.add_argument("--max-retries", type=int, default=10,
                    help="schedule：连不上本地接口时同一条最多重试几次（每次隔 60 秒，默认 10）")
    ap.add_argument("--check-after", dest="check_after", type=int, default=CHECK_AFTER_MIN,
                    help="check/checkd：任务建完多少分钟后查发布结果（默认 15）")
    ap.add_argument("--check-before", dest="check_before", type=int, default=CHECK_BEFORE_NEXT_MIN,
                    help="checkd：下一轮发布前多少分钟再查上一轮（默认 15）")
    ap.add_argument("--check-interval", dest="check_interval", type=int, default=5,
                    help="checkd：最长多久查一圈（分钟，默认 5；到点事件会提前唤醒）")
    ap.add_argument("--retry", dest="retry", action="store_true", default=True,
                    help="check：明确失败当场重发（2026-11-18 起这是默认行为，checkd 也一样；只看不动用 --no-retry）")
    ap.add_argument("--no-retry", dest="no_retry", action="store_true",
                    help="checkd/check：只查只记，不建重发任务（先观察时用这个）")
    ap.add_argument("--retry-now", dest="retry_now", action="store_true",
                    help="重发连「给这台接下来还没发的轮次让位」都不等，立刻建任务"
                         "（会压短同一台环境的发布间隔，确认要这样才用）")
    ap.add_argument("--retry-wait-gap", dest="retry_wait_gap", action="store_true",
                    help="退回旧行为：明确失败也要等满「这台今天最后一次发布 + min_spacing_minutes」才重发"
                         "（默认不等——那次根本没发出去，客户端点重试也是当场就发）")
    ap.add_argument("--retry-ambiguous", dest="retry_ambiguous", action="store_true",
                    help="把「结果不明」（已有链接 / 4041 / E114204）也当失败重发。"
                         "默认绝不这么做：这类很可能已经发出去了，重复发布掉权重。"
                         "放宽出来的条目仍旧等满间隔，不享受当场重发")
    ap.add_argument("--max-retry", dest="max_retry", type=int, default=1,
                    help="单条视频今天最多自动重发几次（默认 1，超过只报不试，避免重试套重试）")
    ap.add_argument("--force-check", dest="force_check", action="store_true",
                    help="check：已查成终态的任务也重新查一遍（默认跳过，防重复触发重试）")
    ap.add_argument("--out", default="", help="retime：输出文件名（默认 <原名>_改时.xlsx；--in-place 时忽略）")
    ap.add_argument("--in-place", action="store_true",
                    help="retime：直接覆盖原粘贴包（只允许覆盖 outputs_dir 里的文件，E 盘原件不碰）")
    ap.add_argument("--mark", default="", help="ledger：要标记为已发布的视频名，逗号分隔")
    ap.add_argument("--write-back", action="store_true", help="ledger：直接改原始台账（默认只出副本）")
    args = ap.parse_args()

    if not os.path.exists(args.config):
        # load_config 找不到文件会退回内置 DEFAULT_CONFIG：那套默认值也是真路径，
        # --config 打错一个字就神不知鬼不觉地换了一套台账/输出目录（2026-11-17 巡检时撞上这坑）。
        die("配置文件不存在：%s（退回内置默认值太危险，直接停）" % args.config)
    cfg = load_config(args.config)
    if getattr(args, "url_prefix", ""):
        cfg["video_source"] = {"mode": "url_prefix", "url_prefix": args.url_prefix,
                               "root": cfg["video_source"].get("root", "")}
    api = Hubstudio(cfg["hubstudio"]["host"], int(cfg["hubstudio"]["port"]),
                    float(cfg["hubstudio"].get("min_interval_sec", 0.8)))
    if args.command == "phones":
        cmd_phones(cfg, api, args)
    elif args.command == "today":
        cmd_today(cfg, api, args)
    elif args.command == "status":
        cmd_status(cfg, api, args)
    elif args.command == "plan":
        cmd_plan(cfg, api, args)
    elif args.command == "submit":
        cmd_submit(cfg, api, args)
    elif args.command == "watch":
        cmd_watch(cfg, api, args)
    elif args.command == "ledger":
        cmd_ledger(cfg, api, args)
    elif args.command == "cancel":
        cmd_cancel(cfg, api, args)
    elif args.command == "pack":
        cmd_pack(cfg, api, args)
    elif args.command == "paste":
        cmd_paste(cfg, api, args)
    elif args.command == "verify":
        cmd_verify(cfg, api, args)
    elif args.command == "watchdog":
        cmd_watchdog(cfg, api, args)
    elif args.command == "retime":
        cmd_retime(cfg, api, args)
    elif args.command == "schedule":
        cmd_schedule(cfg, api, args)
    elif args.command == "check":
        cmd_check(cfg, api, args)
    elif args.command == "checkd":
        cmd_checkd(cfg, api, args)
    elif args.command == "retry":
        cmd_retry(cfg, api, args)
    else:
        die("未实现的命令：%s" % args.command)


if __name__ == "__main__":
    main()
