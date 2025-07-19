# 💡 Paper Invoice Filler
This Python project automates the process of filling in invoice data and payment confirmation stamps on physical invoices' reverse sides. Leveraging Python's automation capabilities, the project generates appropriately named PNG files and uses them as input for a physical printer.

## 🧰 Overview
The project scope involved operations on multiple files simultaneously, reading data from an Excel file, and creating files with strictly defined naming rules. It aimed to bridge digital data with paper output through Python automation.

## 🕒 Project Details
Duration: May 2023 - August 2023
Company: ART-COM Sp. z o.o.

## 🛠️ Tools & Technologies
### 🧑‍💻 Python and Excel integration
### 📥 Python libraries: pandas, PIL, os
Custom Python module: kwotaslownie # https://github.com/dowgird/pyliczba
Variables Used: list, string, integer, float, tuple
Usage of "for" loops for repetitive file operations
Integration with accounting software's invoice data
Customizable font styles, sizes, and colors for printable invoice details
Pillow library for background templates and font styles
Handling different invoice types ("FZ", "FZK", "FZKOR") with specific formatting
Saving generated invoice images in an "Output" folder

## ✅ Benefits
Improved accuracy: Reducing manual data entry errors
Time savings: Speeding up the task of generating printable invoice details
Consistency: Standardizing all generated invoice details
Scalability: Applicable to a large volume of invoices, enhancing efficiency
Flexibility: Customizable design elements for tailored invoices

## 🔁 Setup:
Ensure Python 3.x and required dependencies are installed.
Place invoice data in an Excel file named Invoice Data.xlsx.
Adjust settings within the Python script as needed.

## 🔤 Run:
Execute the Python script to process and print invoices.
Generated invoice images will be stored in the Output directory.

# 🧪 Example usage of print_invoice function:

**🕵️‍♂️ Example 1:**

```python:
# Import necessary libraries
from kwotaslownie import kwotaslownie
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import os
import pandas as pd
import time
import pyautogui as pag

print_invoice(invoice_total = 1000.0, invoice_date = '2023-12-31', invoice_id = 123456, invoice_type = 'FZ')
```

**🖱️ Example 2:**

``` python:
# Import necessary libraries
from kwotaslownie import kwotaslownie
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import os
import pandas as pd
import time
import pyautogui as pag

# Function definition (place the function code here)

# Sample data
sample_data = {
    'invoice_total': 1234.56,
    'invoice_date': '2024-01-28',
    'invoice_id': 1001,
    'invoice_type': 'FZ',
    'invoice_department': 'DGK',
}

# Create a DataFrame with the sample data
sample_df = pd.DataFrame([sample_data])

# Set the working directory
os.chdir(r"C:\Users\User\InvoiceProject")

# Iterate over rows in the DataFrame
for _, record in sample_df.iterrows():
    # Print the invoice using the print_invoice function
    print_invoice(
        invoice_total=record['invoice_total'],
        invoice_date=record['invoice_date'],
        invoice_id=record['invoice_id'],
        invoice_type=record['invoice_type'],
        invoice_department=record['invoice_department'],
        photo_input="STAMP1.jpg",
        font_input="Courier Prime.ttf",
        font_size=35,
        color=(0, 0, 0)
    )

    # Start printing the generated invoice
    os.startfile(os.path.join(os.getcwd(), "Output", f"Invoice {record['invoice_id']}.png"), "print")

    # Wait for the print to complete
    time.sleep(3)
    pag.click(pag.locateOnScreen("Drukuj.png"))
    time.sleep(10)

# Clean up: Remove the generated invoice files
for _, record in sample_df.iterrows():
    os.remove(os.path.join(os.getcwd(), "Output", f"Invoice {record['invoice_id']}.png"))
```

## 🧾 Notes
Ensure proper printer settings and image templates configuration before running the script.
Adjust time delays (time.sleep()) for system-specific printing speed.
