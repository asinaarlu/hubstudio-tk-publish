# -*- coding: utf-8 -*-
"""批量发布后的只读盯梢：把今天本脚本建的所有任务的运行明细汇总成一行。

用法：python -X utf8 batch_watch.py            # 打一行摘要（状态有变化才打，配合外部循环）
      python -X utf8 batch_watch.py --full     # 连失败的执行逐条列出来

只做读接口（task/detail、task/execution/page），不建任务、不取消、不碰云手机开关机。
"""
import io
import json
import os
import sys
import datetime

import tk_publish as T

HERE = os.path.dirname(os.path.abspath(__file__))


def all_task_ids(cfg):
    ids = []
    for fn in ("tk_publish_state.json", "tk_publish_schedule_state.json"):
        p = os.path.join(HERE, fn)
        if not os.path.exists(p):
            continue
        try:
            st = json.load(io.open(p, encoding="utf-8"))
        except Exception:
            continue
        for tid, v in (st.get("tasks") or {}).items():
            if str(v.get("createdAt") or "")[:10] == datetime.date.today().strftime("%Y-%m-%d"):
                ids.append((tid, v.get("taskName") or "", v.get("mobile") or ""))
    seen, out = set(), []
    for t in ids:
        if t[0] not in seen:
            seen.add(t[0])
            out.append(t)
    return out


def main():
    cfg = T.load_config(os.path.join(HERE, "tk_publish.config.json"))
    api = T.Hubstudio(cfg["hubstudio"]["host"], int(cfg["hubstudio"]["port"]),
                      float(cfg["hubstudio"].get("min_interval_sec", 0.8)))
    full = "--full" in sys.argv
    cnt = {"SUCCEEDED": 0, "FAILED": 0, "RUNNING": 0, "PENDING": 0, "CANCELED": 0}
    errs, links, tasks = [], [], 0
    for tid, name, mobile in all_task_ids(cfg):
        try:
            exs = api.executions(tid)
        except T.ApiError as e:
            errs.append("%s 读取失败 %s" % (name, e))
            continue
        tasks += 1
        for e in exs:
            s = str(e.get("status") or "")
            cnt[s] = cnt.get(s, 0) + 1
            if s == "FAILED":
                errs.append("%s mobile=%s %s %s" % (name, e.get("mobileId"),
                                                    e.get("errorCode"), (e.get("errorMessage") or "")[:40]))
            u = (e.get("result") or {}).get("publishedUrl") or ""
            if u:
                links.append(u)
    line = ("任务 %d 个 | 成功 %d 失败 %d 运行 %d 排队 %d 取消 %d | 已出链接 %d"
            % (tasks, cnt["SUCCEEDED"], cnt["FAILED"], cnt["RUNNING"], cnt["PENDING"],
               cnt["CANCELED"], len(links)))
    if errs:
        line += " | 异常 %d 条：%s" % (len(errs), "；".join(errs[:3]))
    print(line)
    if full:
        for e in errs:
            print("   ", e)


if __name__ == "__main__":
    main()
