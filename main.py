from datetime import datetime

def print_current_date_time():
    current_date_time = datetime.now()
    print(f"Current Date and Time: {current_date_time}")

if __name__ == "__main__":
    print_current_date_time()