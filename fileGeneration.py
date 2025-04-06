import os
from reportlab.pdfgen import canvas

file_paths = []
directory = r"c:\Users\KarinaKresnadi\Documents\test"  # Update this path

n = 5 # Update this

# Ensure directory exists
os.makedirs(directory, exist_ok=True)

for i in range(1, n):
    file_path = os.path.normpath(f"{directory}\DMS Sample - 10 - TST - test_{i+1}.pdf")  # Ensure consistent separators
    c = canvas.Canvas(file_path)
    c.drawString(100, 750, f"This is file {i+1}.")
    c.save()
    file_paths.append(file_path)

print("Files created.")