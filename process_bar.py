import time

def show_progress(duration):
    total_steps = 100
    interval = duration / total_steps

    for i in range(1, total_steps + 1):
        print(f"Đã xử lý: {i}%")
        time.sleep(interval)
    print("done!")
#show_progress(15)