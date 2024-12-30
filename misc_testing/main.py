import multiprocessing  
import time
import os

def worker():
    while True:
            print(f"waiting {os.getpid()}")
            time.sleep(0.5)

import multiprocessing
p = multiprocessing.Process(target=worker, name="customprocess")
os.getpid()

p.start()
