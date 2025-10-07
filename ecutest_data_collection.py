import api
import os
import csv

folder = r"C:\_git_tm\ROOT_GEN6\xxx"  # <-- Your folder with .pkg files
output = r"C:\Users\Mootaz\Desktop\reboustness\reboustness.csv"

results = []
for f in os.listdir(folder):
    if f.endswith(".pkg"):
        path = os.path.join(folder, f)
        pkg = api.ObjectApi.OpenPackage(path)
        desc = pkg.GetDescription()
        id = pkg.GetId() if hasattr(pkg, "GetId") else f
        results.append([id, desc])

with open(output, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["ID", "Description"])
    writer.writerows(results)

print("Saved", len(results), "entries to", output)
