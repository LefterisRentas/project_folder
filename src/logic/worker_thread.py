# logic/worker_thread.py

import threading

class WorkerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        # Initialize any necessary attributes here

    def run(self):
        # This method will be executed when the thread is started
        while True:
            # Implement your worker thread's logic here
            print("Worker thread is running")
            # You can use sleep to simulate some work being done
            threading.Event().wait(5)  # Sleep for 5 seconds

