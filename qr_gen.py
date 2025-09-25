import os
import openpyxl
import qrcode

# Load Excel data
data = openpyxl.load_workbook("data.xlsx")
sheet = data.active

# Create output folder if it does not exist
os.makedirs("QR", exist_ok=True)

ids = []
names = []

# Assuming first row contains data (not headers) and first 10 rows are required
for i in range(1, 12):  
    id_cell = sheet.cell(i, 1).value   # First column -> ID
    name_cell = sheet.cell(i, 2).value # Second column -> Name

    ids.append(id_cell)
    names.append(name_cell)

# Print to check
print("IDs:", ids)
print("Names:", names)

# Generate QR codes
for idx in range(len(ids)):
    qr_content = ids[idx]      # Embed ID inside QR
    filename = names[idx]      # Use Name as filename
    
    img = qrcode.make(qr_content)
    img.save(f"QR/{filename}.png")
