import qrcode
import os
import csv

# Define the output directory
output_dir = r"C:\Users\fatah\OneDrive\Desktop\MY Files\Job_Application\Interviews-Technical_Tasks\KSAPortfolia\portfolio_project\assets"

# URLs to generate QR codes for
urls = {
    "tableau_portfolio": "https://public.tableau.com/app/profile/fathallah.elmasri",
    "github": "https://github.com/Elmasri-Fathallah",
    "linkedin": "https://www.linkedin.com/in/fathallah-elmasri",
    "email": "mailto:fathallah.elmasri@gmail.com"
}

# Generate QR codes
qr_map = []
for name, url in urls.items():
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    filename = f"QR_{name}.png"
    filepath = os.path.join(output_dir, filename)
    img.save(filepath)
    
    qr_map.append({
        "name": name,
        "url": url,
        "filename": filename,
        "filepath": filepath
    })
    print(f"Generated QR code for {name}: {filename}")

# Create CSV map
csv_path = os.path.join(output_dir, "QR_map.csv")
with open(csv_path, 'w', newline='') as csvfile:
    fieldnames = ['name', 'url', 'filename', 'filepath']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(qr_map)

print(f"\nQR map saved to: {csv_path}")
print("\nAll QR codes generated successfully!")
print("\nTo run this script, ensure you have qrcode installed:")
print("pip install qrcode[pil]")
