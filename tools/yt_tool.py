#!/usr/bin/env python3
"""YouTube 검색·자막 수집 도구 (사전점검 체크리스트용).

사용법
  python3 tools/yt_tool.py search "아파트 사전점검 셀프" [--n 25]
      → 검색 결과를 조회수 순으로 출력 (id, 조회수, 길이(초), 채널, 제목)
  python3 tools/yt_tool.py fetch <영상 id 또는 URL> [...]
      → 한국어 자막을 받아 tools/transcripts/<id>.txt 로 저장하고
        data/youtube_videos.tsv 에 조회수·좋아요 등을 추가/갱신

필요: yt-dlp (pip install yt-dlp). 클라우드 서버는 YouTube 봇 차단을 받는 경우가 많아
player_client=tv_embedded → android_vr 순서로 시도한다.
"""
import argparse
import csv
import os
import re
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TRANSCRIPTS = os.path.join(ROOT, "tools", "transcripts")
VIDEOS_TSV = os.path.join(ROOT, "data", "youtube_videos.tsv")
FIELDS = ["id", "views", "likes", "upload_date", "channel", "title"]
CLIENTS = ["tv_embedded", "android_vr"] * 3  # 봇 차단이 들쭉날쭉해서 번갈아 3번까지 재시도


def video_id(s):
    m = re.search(r"(?:v=|youtu\.be/|shorts/)([\w-]{11})", s)
    return m.group(1) if m else s.strip()


def search(query, n):
    out = subprocess.run(
        ["yt-dlp", "--flat-playlist", "--print",
         "%(id)s\t%(view_count)s\t%(duration)s\t%(channel)s\t%(title)s",
         f"ytsearch{n}:{query}"],
        capture_output=True, text=True, timeout=300).stdout
    rows = [l.split("\t") for l in out.splitlines() if l.count("\t") >= 4]
    rows.sort(key=lambda r: -int(r[1]) if r[1].isdigit() else 0)
    for r in rows:
        print("\t".join(r))


def vtt_to_text(path):
    lines = []
    for l in open(path, encoding="utf-8"):
        l = l.strip()
        if not l or "-->" in l or l.startswith(("WEBVTT", "Kind:", "Language:")) or l.isdigit():
            continue
        l = re.sub(r"<[^>]+>", "", l).strip()
        if l and l not in lines[-3:]:
            lines.append(l)
    return " ".join(lines)


def load_videos():
    if not os.path.exists(VIDEOS_TSV):
        return {}
    with open(VIDEOS_TSV, encoding="utf-8") as fh:
        return {r["id"]: r for r in csv.DictReader(fh, delimiter="\t")}


def save_videos(videos):
    rows = sorted(videos.values(), key=lambda r: -int(r["views"]) if str(r["views"]).isdigit() else 0)
    with open(VIDEOS_TSV, "w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, delimiter="\t", lineterminator="\n")
        w.writeheader()
        w.writerows({k: r.get(k, "") for k in FIELDS} for r in rows)


def fetch(ids):
    os.makedirs(TRANSCRIPTS, exist_ok=True)
    videos = load_videos()
    for raw in ids:
        vid = video_id(raw)
        meta = None
        for client in CLIENTS:
            res = subprocess.run(
                ["yt-dlp", "--skip-download", "--write-subs", "--write-auto-subs",
                 "--sub-langs", "ko,ko-orig", "--sub-format", "vtt",
                 "--extractor-args", f"youtube:player_client={client}",
                 "-o", os.path.join(TRANSCRIPTS, "%(id)s.%(ext)s"),
                 "--print", "META\t%(id)s\t%(view_count)s\t%(like_count)s\t%(upload_date)s\t%(channel)s\t%(title)s",
                 "--no-simulate", f"https://www.youtube.com/watch?v={vid}"],
                capture_output=True, text=True, timeout=300)
            meta = next((l for l in res.stdout.splitlines() if l.startswith("META\t")), None)
            if meta:
                break
            time.sleep(5)
        if not meta:
            print(f"[실패] {vid}: YouTube가 접속을 막았거나 영상이 없습니다", file=sys.stderr)
            continue
        _, *vals = meta.split("\t")
        videos[vid] = dict(zip(FIELDS, vals))
        vtts = [f for f in os.listdir(TRANSCRIPTS) if f.startswith(vid) and f.endswith(".vtt")]
        vtts.sort(key=lambda f: "orig" in f)  # 'ko'(수동/번역) 우선, 없으면 'ko-orig'(자동)
        if vtts:
            text = vtt_to_text(os.path.join(TRANSCRIPTS, vtts[0]))
            with open(os.path.join(TRANSCRIPTS, vid + ".txt"), "w", encoding="utf-8") as fh:
                fh.write(text)
            print(f"[완료] {vid} 자막 {len(text):,}자 · 조회수 {vals[1]} · {vals[5]}")
        else:
            print(f"[자막 없음] {vid} · {vals[5]}")
    save_videos(videos)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("search")
    s.add_argument("query")
    s.add_argument("--n", type=int, default=25)
    f = sub.add_parser("fetch")
    f.add_argument("ids", nargs="+")
    a = ap.parse_args()
    if a.cmd == "search":
        search(a.query, a.n)
    else:
        fetch(a.ids)


if __name__ == "__main__":
    main()
