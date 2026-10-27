
import os
import random
import string
from openpyxl import Workbook

folder = r"D:\Excel_Multithreading_50GB\excel_data"

target_size = 10 * 1024 * 1024 * 1024
rows_per_file = 50000

os.makedirs(folder, exist_ok=True)

existing_files = [
    f for f in os.listdir(folder)
    if f.startswith("data_") and f.endswith(".xlsx")]

total_size = sum(
    os.path.getsize(os.path.join(folder, f))
    for f in existing_files)

file_numbers = []
for filename in existing_files:
    try:
        number = int(filename.replace("data_", "").replace(".xlsx", ""))
        file_numbers.append(number)
    except ValueError:
        pass

file_number = max(file_numbers, default=0) + 1
record_id = len(existing_files) * rows_per_file + 1

print("Existing files:", len(existing_files))
print("Existing size:", round(total_size / (1024 ** 3), 2), "GB")

while total_size < target_size:

    path = os.path.join(folder, f"data_{file_number}.xlsx")

    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Data")

    sheet.append([
        "Record_ID", "Name", "Department",
        "City", "Amount", "Message"])

    for i in range(rows_per_file):
        name = "".join(random.choices(string.ascii_letters, k=10))
        message = "Hello " + "".join(
            random.choices(string.ascii_letters, k=80))

        sheet.append([
            record_id,
            name,
            random.choice(["IT", "HR", "Sales", "Finance"]),
            random.choice(["Delhi", "Mumbai", "Chennai", "Bengaluru"]),
            random.randint(1000, 100000),
            message])

        record_id += 1

    workbook.save(path)

    size = os.path.getsize(path)
    total_size += size

    print(
        f"Created: data_{file_number}.xlsx | "
        f"File size: {size / (1024 ** 2):.2f} MB | "
        f"Total: {total_size / (1024 ** 3):.2f} GB")

    file_number += 1

print("\nDataset generation completed.")
print("Total Excel files:", len([
    f for f in os.listdir(folder)
    if f.startswith("data_") and f.endswith(".xlsx")]))
print("Total size:", round(total_size / (1024 ** 3), 2), "GB")