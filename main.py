import os
import re
import openpyxl
import webbrowser
import customtkinter as ctk
from tkinter import StringVar, IntVar
from tkinter.filedialog import askopenfilename
from pdfminer.high\_level import extract\_text

# Define the function to extract details from PDF

def extract\_details\_from\_pdf(pdf\_file):
text = extract\_text(pdf\_file)

```
# Extract CU IN No.
cu_in_pattern = re.compile(r"CU IN No\.:(.*?)\n")
cu_in_match = cu_in_pattern.search(text)
cu_in = cu_in_match.group(1).strip() if cu_in_match else ""

# Extract CU SN No.
cu_sn_pattern = re.compile(r"CU SN No\.:(.*?)\n")
cu_sn_match = cu_sn_pattern.search(text)
cu_sn = cu_sn_match.group(1).strip() if cu_sn_match else ""

# Document_Type
Document_Type = re.compile(r"Document_Type:(.*?)\n")
Document_Type_match = Document_Type.search(text)
Document_Type = Document_Type_match.group(1).strip() if Document_Type_match else ""

# PIN
PIN = re.compile(r"PIN:(.*?)\n")
PIN_match = PIN.search(text)
PIN = PIN_match.group(1).strip() if PIN_match else ""

# INVOICE_NO
INVOICE_NO_pattern = re.compile(r"INVOICE_NO\s*:\s*(.*?)\n")
INVOICE_NO_match = INVOICE_NO_pattern.search(text)
INVOICE_NO = INVOICE_NO_match.group(1).strip() if INVOICE_NO_match else ""

#Invoice Date
Invoice_Date_pattern = re.compile(r"Invoice\s+Date\s*:\s*(.*?)\n")
Invoice_Date_match = Invoice_Date_pattern.search(text)
Invoice_Date = Invoice_Date_match.group(1).strip() if Invoice_Date_match else ""

#PIN_No
PIN_No_pattern = re.compile(r"PIN_No\s*:\s*(.*?)\n")
PIN_No_match = PIN_No_pattern.search(text)
PIN_No = PIN_No_match.group(1).strip() if PIN_No_pattern else ""

# Extract TOTAL
total_match = re.search(r"TOTAL\s*:\s*([0-9,.]+)", text, re.IGNORECASE)
total = total_match.group(1).strip() if total_match else ""
if not total:
    total_match = re.search(r"([0-9,.]+)\s*TOTAL", text, re.IGNORECASE)
    total = total_match.group(1).strip() if total_match else ""

# Extract VAT percentage
vat_percentage_match = re.search(r"VAT\s*:\s*(\d+\.\d+)%", text, re.IGNORECASE)
vat_percentage = vat_percentage_match.group(1).strip() if vat_percentage_match else ""

# Extract customer using line matching
lines = text.split('\n')

# Find the line index containing "Customer:"
customer_index = None
customer_code = None
for i, line in enumerate(lines):
    if "Customer :" in line:
        customer_index = i
        match = re.search(r"(?<=Customer : )\w+", line)  # Extract customer code (e.g., lnk123)
        if match:
            customer_code = match.group()
        break

# Extract the customer name and address
customer = ""
if customer_index is not None and customer_code:
    # Find the indices of "Order_Date" and "Delivery_Note_No" lines
    order_date_index = None
    delivery_note_index = None
    for j, line in enumerate(lines[customer_index:]):
        if "Order_Date:" in line:
            order_date_index = customer_index + j
        elif "Delivery_Note_No:" in line:
            delivery_note_index = customer_index + j
            break  # Stop searching once Delivery_Note_No is found

    # Extract customer details between Order_Date and Delivery_Note_No, excluding dates
    if order_date_index is not None and delivery_note_index is not None:
        for line in lines[order_date_index + 1:delivery_note_index]:
            if line and "Order_No:" not in line and not re.match(r"\d{2}/\d{2}/\d{2}",
                                                                 line):  # Exclude lines with date format
                customer += line.strip() + " "

customer = f"{customer_code} {customer.strip()}"  # Add customer code and remove trailing whitespace

return cu_in, cu_sn, Document_Type, customer, PIN, INVOICE_NO, Invoice_Date, PIN_No, total, vat_percentage, pdf_file
```

# Define the function to write onto excel worksheet

def write\_to\_excel(details\_list, output\_file):
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(\["cu\_in","cu\_sn","Document\_Type","customer","PIN","INVOICE\_NO","Invoice\_Date","PIN\_No","total","vat\_percentage", "PDF\_File"])
for details in details\_list:
sheet.append(details)
workbook.save(output\_file)

# Define the function to browse the PDF file using file dialog

def browse\_files():
file\_names = ctk.CTkFileDialog.askopenfilenames(
title="Select PDF files",
filetypes=(("PDF files", "*.pdf"), ("All files", "*.\*"))
)
for file\_name in file\_names:
listbox.insert("end", file\_name)

# Define the function to extract text from PDF files

def extract\_text():
pdf\_files = list(listbox.get(0, "end"))
details\_list = \[]
for pdf\_file in pdf\_files:
details = extract\_details\_from\_pdf(pdf\_file)
details\_list.append(details)
output\_file = "TaxInvoice.xlsx"
write\_to\_excel(details\_list, output\_file)
ctk.CTkMessageBox.showinfo(
title="Text Extraction Completed",
message=f"Details extracted from PDFs and saved to {output\_file}."
)

# Create a customtkinter window

root = ctk.CTk()

# Set the window title

root.title("PDF Details Extractor")

# Create a customtkinter label to display the instruction

instruction\_label = ctk.CTkLabel(
root,
text="Select PDF Files:"
)
instruction\_label.pack(pady=10)

# Create a textbox to display selected PDF files

listbox = ctk.CTkTextbox(
root,
width=50,
height=10
)
listbox.pack()

# Create a button to browse for PDF files

browse\_button = ctk.CTkButton(
root,
text="Browse",
command=browse\_files
)
browse\_button.pack()

# Create radiobuttons to select details to include

selected\_option = StringVar()

cu\_in\_radiobutton = ctk.CTkRadiobutton(
root,
text="CU IN",
variable=selected\_option,
value="CU IN"
)
cu\_in\_radiobutton.pack()

cu\_sn\_radiobutton = ctk.CTkRadiobutton(
root,
text="CU SN",
variable=selected\_option,
value="CU SN"
)
cu\_sn\_radiobutton.pack()

document\_type\_radiobutton = ctk.CTkRadiobutton(
root,
text="Document Type",
variable=selected\_option,
value="Document Type"
)
document\_type\_radiobutton.pack()

customer\_radiobutton = ctk.CTkRadiobutton(
root,
text="Customer",
variable=selected\_option,
value="Customer"
)
customer\_radiobutton.pack()

pin\_radiobutton = ctk.CTkRadiobutton(
root,
text="PIN",
variable=selected\_option,
value="PIN"
)
pin\_radiobutton.pack()

invoice\_no\_radiobutton = ctk.CTkRadiobutton(
root,
text="Invoice No",
variable=selected\_option,
value="Invoice No"
)
invoice\_no\_radiobutton.pack()

invoice\_date\_radiobutton = ctk.CTkRadiobutton(
root,
text="Invoice Date",
variable=selected\_option,
value="Invoice Date"
)
invoice\_date\_radiobutton.pack()

pin\_no\_radiobutton = ctk.CTkRadiobutton(
root,
text="PIN No",
variable=selected\_option,
value="PIN No"
)
pin\_no\_radiobutton.pack()

total\_radiobutton = ctk.CTkRadiobutton(
root,
text="Total",
variable=selected\_option,
value="Total"
)
total\_radiobutton.pack()

vat\_percentage\_radiobutton = ctk.CTkRadiobutton(
root,
text="VAT Percentage",
variable=selected\_option,
value="VAT Percentage"
)
vat\_percentage\_radiobutton.pack()

# Create a button to extract text from PDF files

extract\_button = ctk.CTkButton(
root,
text="Extract Text",
command=extract\_text
)
extract\_button.pack(pady=10)

# Start the main loop of the window

root.mainloop()
