import pandas as pd
import random
from datetime import datetime
import time

file_path = "SimulatedStream.xlsx"  # Ensure this path matches your file location

def generate_new_row():
    return {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "region": random.choice(['East', 'West', 'North', 'South']),
        "temperature": round(random.uniform(25, 35), 1),
        "sales": random.randint(100, 1000)
    }

while True:
    try:
        # Load existing Excel data
        df = pd.read_excel(file_path)

        # Append a new row
        new_row = generate_new_row()
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        # Save back to Excel
        df.to_excel(file_path, index=False)

        print(f"Added new row at {new_row['timestamp']}")
    except Exception as e:
        print("Error:", e)

    time.sleep(30)  # Wait 30 seconds before adding next row
