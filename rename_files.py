import os

base_dir = r"<workspace>\mail-auto\attachments_20260531_132532"

for subdir in ["示例员工", "水彬"]:
    subdir_path = os.path.join(base_dir, subdir)
    if not os.path.exists(subdir_path):
        continue

    for filename in os.listdir(subdir_path):
        if "b'" in filename:
            old_path = os.path.join(subdir_path, filename)
            new_filename = filename.replace("b'", "").replace("'", "")
            new_path = os.path.join(subdir_path, new_filename)
            os.rename(old_path, new_path)
            print(f"Renamed: {filename} -> {new_filename}")

print("\nAll files renamed successfully!")
