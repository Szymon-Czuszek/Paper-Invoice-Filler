#!/usr/bin/env python
# coding: utf-8

# In[2]:


from kwotaslownie import kwotaslownie

from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

import os

import pandas as pd

import time

import pyautogui as pag


# In[3]:


def print_invoice(invoice_total, invoice_date, invoice_id, invoice_type = "FZ", invoice_department = "DGK",
                  photo_input = "STAMP1.jpg", font_input = "Courier Prime.ttf", font_size = 35, color = (0, 0, 0)):
    
    """This function takes in invoice data (invoice_total, invoice_date, invoice_id, invoice_department), template photo 
    (photo_input), and font related spcifiers (font_input, font_size, color) to create ready to be printed invoice details.
    
    Args:
        invoice_total(float): Total amount found in accounting software's invoice data.
        invoice_date(string): Date found in the accounting software's invoice data.
        invoice_id(variant): Identification number found in the accounting software's invoice data.
        invoice_type(string): Type of the invoice. Invoice types = {'magazyn':'FZ', 'koszt':'FZK', 'korekta':'FZKOR'}
        invoice_department(string): Department name of the worker doing the job.
        photo_input(string): Name of the file containing the invoice background template in the image format.
        font_input(string): Name of the file containing the font type template in the TrueType font format.
        font_size(integer): Value which gets used to generate the font size used in the printed invoice details.
        color(tuple): Tuple containing three integers refering to the RGB color values.
        
    Returns:
        PIL.Image.Image: Image containing readily printable invoice details.
        
    """
    invoice_type = invoice_type.upper()
    background = Image.open(photo_input).convert("RGBA") # Loads the image template as background
    draw = ImageDraw.Draw(background) # Creates the object to draw on the background as draw
    x, y = background.size # Extracts background dimensions to use them for drawing as x and y
    font_type = ImageFont.truetype(font_input, font_size) # Loads the TrueType font and font size as font_type
    full_id = f'{invoice_type}/{str(invoice_id).rjust(6, "0")}/{invoice_date[-4::]}'
    
    if invoice_type == "FZ" or invoice_type == "FZK":
        amount_int = invoice_total
        amount = kwotaslownie(invoice_total, fmt = 1)
    elif invoice_type == "FZKOR":
        amount_int = - invoice_total
        amount = f"minus {kwotaslownie(invoice_total, fmt = 1)}"
    else:
        raise Exception("Wrong invoice type inputted! Please choose one in ['FZ', 'FZK', 'FZKOR'].")

    draw.text((0.060 * x, 0.060 * y), invoice_date, fill = color, font = font_type) # top left date
    draw.text((0.060 * x, 0.120 * y), invoice_date, fill = color, font = font_type) # mid left date

    draw.text((0.150 * x, 0.364 * y), f"{amount_int} ZŁ", fill = color, font = font_type) # int total

    len_amount = int(len(amount) / 3)
        
    pauza = "" if amount[:len_amount - 2:].endswith(" ") else "-"
    
    draw.text((0.150 * x, 0.387 * y), amount[:len_amount - 2:] + pauza, fill = color, font = font_type)
        
    pauza = "" if amount[len_amount - 2:2 * len_amount:].endswith(" ") else "-"
    
    draw.text((0.055 * x, 0.411 * y), amount[len_amount - 2:2 * len_amount:] + pauza, fill = color, font = font_type)
    draw.text((0.055 * x, 0.433 * y), amount[2 * len_amount::], fill = color, font = font_type)

    draw.text((0.222 * x, 0.666 * y), invoice_date, fill = color, font = font_type) # bot right date
    draw.text((0.125 * x, 0.709 * y), invoice_department, fill = color, font = font_type) # department = "DGK"
    draw.text((0.279 * x, 0.704 * y), full_id, fill = color, font = font_type) # full invoice id
    
    background.save(os.path.join(os.getcwd(), "Output", f"Invoice {invoice_id}.png"), format = "png")

    return background


# In[4]:


if __name__ == "__main__":
    os.chdir(r"C:\Users\User\InvoiceProject")
    
    data = pd.read_excel("Invoice Data.xlsx", sheet_name = "Sheet1", header = 0, index_col = "record")
    data.dropna(axis = 0, how = "any", inplace = True) # Drop all rows where at least one value is NaN
    data["total"] = data["total"].astype(float)
    data["date"] = data["date"].dt.strftime('%d.%m.%Y')
    data["id"] = data["id"].astype(int)
    invoice_data = data.values.tolist()
    print(data)
    
    for record in invoice_data:
    
        print(record)
    
        print_invoice(invoice_total = record[0],
                      invoice_date = record[1],
                      invoice_id = record[2],
                      invoice_type = record[3]
                     )
        
        os.startfile(os.path.join(
            os.getcwd(), "Output", f"Invoice {record[2]}.png"), "print")
        
        time.sleep(3)
        pag.click(pag.locateOnScreen("Drukuj.png"))
        time.sleep(10)
        
    for record in invoice_data:
        os.remove(os.path.join(
            os.getcwd(), "Output", f"Invoice {record[2]}.png"))


# ###### The End
