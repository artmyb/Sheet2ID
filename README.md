# Sheet2ID – Automated ID Card Generator

**Sheet2ID** is a professional-grade Python application built with **Tkinter** that allows users to batch-generate ID cards from a spreadsheet.  
It features a real-time visual editor to map data fields, position photos, and export high-quality, print-ready PDFs.

---

## Features

- **Data Driven**  
  Import records directly from `.csv` or `.xlsx` files.

- **Live Preview**  
  See changes instantly on a virtual card as you adjust coordinates and font sizes.

- **Dynamic Photo Mapping**  
  Automatically matches photos from a folder to spreadsheet records based on a **Photo Specifier**  
  (e.g. matching a Student ID or Name to a filename).

- **Custom Templates**  
  Upload your own background image to serve as the card design.

- **Precision Layout**  
  Adjust dimensions in millimeters (mm) and use mouse-wheel scrolling for fine-tuning coordinates.

- **Unicode Support**  
  Built-in support for Arial / Liberation fonts to handle international characters.

- **Print-Ready Export**  
  Generates an A4 PDF with multiple cards per page, including cutting guides.

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/sheet2id.git
cd sheet2id
```

### 2. Install dependencies

Ensure you have **Python 3.x** installed:

```bash
pip install pandas pillow openpyxl reportlab
```

---

## 📖 How to Use

### Load Data
Click **📁 Load Spreadsheet** and select your CSV or Excel file.

### Select Background
Click **🖼️ Select Background** to upload your ID card design (the empty card).

### Setup Photos
- Click **📸 Select Photo Folder** to choose where your headshots are stored.
- Check the **Photo Specifier** box next to the column that matches image filenames  
  (e.g. if the image is `101.jpg`, your ID column should contain `101`).

### Map Fields
- Toggle checkboxes to enable or disable fields on the card.
- Adjust **X and Y coordinates** to position text.
- Use the **Anchor** dropdown (**Left**, **Center**, **Right**) for precise alignment.

### Export
- Navigate through records using **◀ Previous** and **Next ▶** to verify layout.
- When satisfied, click **GENERATE PDF**.

---

## Project Structure

- **IDCardGenerator**  
  Main class handling the GUI and application logic.

- **find_photo()**  
  Logic for fuzzy-matching spreadsheet data to local image files.

- **export_pdf()**  
  Uses ReportLab to calculate A4 grid positions and render the final document.

---

## Requirements

| Library | Purpose |
|------|--------|
| Pandas | Data manipulation and spreadsheet parsing |
| Pillow (PIL) | Image processing and real-time preview rendering |
| ReportLab | High-fidelity PDF generation |
| Tkinter | Standard Python GUI framework |

---

## 🤝 Contributing

Contributions are welcome!  
If you have ideas for new features (such as QR code generation or bulk photo cropping), feel free to fork the repository and submit a pull request.
