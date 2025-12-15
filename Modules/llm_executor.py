# Modules/llm_executor.py
import queue
import threading
import time
from openai import RateLimitError

_llm_queue = queue.Queue()
MAX_WORKERS = 2   # 👈 THIS IS THE KEY

def _llm_worker(worker_id):
    while True:
        fn, args, kwargs, result = _llm_queue.get()
        try:
            result["output"] = fn(*args, **kwargs)
        except RateLimitError as e:
            result["error"] = e
        except Exception as e:
            result["error"] = e
        finally:
            _llm_queue.task_done()

# 🔥 Start EXACTLY 2 workers
for i in range(MAX_WORKERS):
    threading.Thread(
        target=_llm_worker,
        args=(i,),
        daemon=True
    ).start()


def submit_llm_job(fn, *args, **kwargs):
    result = {}
    _llm_queue.put((fn, args, kwargs, result))
    return result


def wait_for_job(job, poll_interval=0.2):
    while "output" not in job and "error" not in job:
        time.sleep(poll_interval)

    if "error" in job:
        raise job["error"]

    return job["output"]
