"""Executable -> mapped Tk event-loop readiness, without UI-tool latency.

Use --rounds 5. 'first' is first invocation only; OS cold cache requires reboot
and is deliberately not claimed by this test. No cache flushing / system changes.
"""
from pathlib import Path
import argparse
import json
import os
import statistics
import subprocess
import tempfile
import time


def measure(executable, rounds):
    rows=[]
    with tempfile.TemporaryDirectory() as folder:
        for index in range(rounds):
            marker=Path(folder)/f'ready-{index}.txt'
            env=dict(os.environ,DAKE_STARTUP_READY_FILE=str(marker),DAKE_STARTUP_PROBE='1')
            start=time.perf_counter_ns()
            p=subprocess.Popen([str(executable)],env=env)
            try:
                p.wait(timeout=30)
                if p.returncode or not marker.exists(): raise RuntimeError(f'no readiness marker: {p.returncode}')
                seconds=(int(marker.read_text())-start)/1e9
                rows.append({'run':index+1,'kind':'first launch (not OS-cold)' if not index else 'warm','seconds':seconds})
            finally:
                if p.poll() is None:p.terminate();p.wait(3)
    return {'exe':str(executable),'runs':rows,'warm_median':statistics.median(r['seconds'] for r in rows[1:])}


if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('exe',type=Path);ap.add_argument('output',type=Path);ap.add_argument('--rounds',type=int,default=6);args=ap.parse_args()
    result=measure(args.exe.resolve(),args.rounds)
    args.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result))
