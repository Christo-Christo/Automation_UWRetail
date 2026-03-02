from syntax.main import main
import time
import datetime

input_path = r"C:\Users\CRISTO.CRISTO\OneDrive - Zurich APAC\Input Sheet RRA_1M.xlsx"

if __name__ == "__main__":
    print("🚀 Starting program...")
    start = time.time()
    main(input_path)
    elapsed = time.time() - start
    formatted = str(datetime.timedelta(seconds=int(elapsed)))
    print(f"\n⏱️ Total runtime: {formatted}")