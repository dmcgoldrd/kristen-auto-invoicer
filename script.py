import pandas as pd
import numpy as np
from fpdf import FPDF
import os
import sys


class PDFInvoice(FPDF):
    def header(self):
        # Logo
        #self.image('logo.png', 10, 8, 33)
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'INVOICE', 0, 1, 'C')
        self.ln(5)  # Add a line break for spacing

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        # Print the page number
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')
        
        # Move below the page number for the thank you message
        self.set_y(self.get_y() + 5)  # Move down from the current position
        self.cell(0, 10, 'Thank you for your business!', 0, 0, 'C')

    # def add_business_info(self, dba, email, address, phone):
    #     self.set_xy(10, 40)  # Adjust the X and Y positions as needed
    #     self.set_font('Arial', 'B', 12)
    #     self.cell(0, 6, dba, 0, 2)  # '2' moves below after printing, with reduced line height for single-spacing
    def add_business_info(self, dba, email, address, phone):
        self.set_xy(10, 40)  # Adjust the X and Y positions as needed
        self.set_font('Arial', 'B', 12)
        self.cell(0, 6, dba, 0, 2)  # '2' moves below after printing, with reduced line height for single-spacing
        self.set_font('Arial', '', 10)
        self.cell(0, 6, email, 0, 2)  # Reduced line height for single-spacing
        # For multi_cell, adjust line height directly in the function call
        address_lines = address.split('\n')
        for line in address_lines:
            self.cell(0, 6, line, 0, 1)
        self.cell(0, 6, phone, 0, 2)  # Reduced line height for single-spacing
        self.ln(5)  # Add a small space before moving to invoice data

 
    def add_invoice_data(self, date, invoice_num, for_):
        right_column_x = 150  # X position for the right-aligned invoice data; adjust as needed
        self.set_xy(right_column_x, 40)  # Match Y position with business info, adjust X as needed
        
        # "Date" in bold
        self.set_font('Arial', 'B', 10)
        self.cell(40, 6, 'Date:', 0, 0)  # No border, no new line
        # Date value in normal font
        self.set_font('Arial', '', 10)
        self.cell(0, 6, date, 0, 1, 'R')  # Align right, with new line
        
        # Move slightly to the right for the next line to align with "Date:"
        self.set_x(right_column_x)
        # "Invoice #" in bold
        self.set_font('Arial', 'B', 10)
        self.cell(40, 6, 'Invoice #:', 0, 0)  # No border, no new line
        # Invoice number in normal font
        self.set_font('Arial', '', 10)
        self.cell(0, 6, invoice_num, 0, 1, 'R')  # Align right, with new line
        
        # Repeat for "For"
        self.set_x(right_column_x)
        self.set_font('Arial', 'B', 10)
        self.cell(40, 6, 'For:', 0, 0)  # No border, no new line
        self.set_font('Arial', '', 10)
        self.cell(0, 6, for_, 0, 1, 'R')  # Align right, with new line

        self.ln(5)  # Add some space before moving to the next section
    # Other methods like add_bill_to, header, footer, add_items_table...

    def add_bill_to(self, address):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, 'Bill To:', 0, 1)
        self.set_font('Arial', '', 10)
        # Convert the address to lines for single-spacing
        address_lines = address.split('\n')
        for line in address_lines:
            self.cell(0, 6, line, 0, 1)
        self.ln(5)  # Add spacing after address

    def add_items_table(self, items):
        line_height = 7

        self.set_font('Arial', 'B', 10)
        # Headers with borders but without fill color
        self.cell(40, line_height, 'Order Number', 1, 0, 'C')
        self.cell(50, line_height, 'Item Number', 1, 0, 'C')
        self.cell(40, line_height, 'SALES', 1, 0, 'C')
        self.cell(30, line_height, 'COMMISSION', 1, 0, 'C')
        self.cell(30, line_height, 'AMOUNT', 1, 1, 'C')
        self.set_font('Arial', '', 10)

        total = 0
        order_num = 0

        for item in items:
            amount = round(float(item['amount']), 2)
            # No borders for item rows
    
            if item['order_number'] == order_num:
                self.cell(40, line_height, '', 0, 0)
            else:
                self.cell(40, line_height, item['order_number'], 0, 0)
                order_num = item['order_number']
            self.cell(50, line_height, item['item_number'], 0, 0)
            self.cell(40, line_height, item['sales'], 0, 0)
            self.cell(30, line_height, item['commission'], 0, 0)
            self.cell(30, line_height, f"${amount:,.2f}", 0, 1, 'R')
            total += float(item['amount'])

        # Set font to bold for the "Total" label
        self.set_font('Arial', 'B', 10)
        self.cell(160, line_height, 'Total', 1, 0, 'R')  # The 'R' alignment for label, now in bold
        
        # Assuming you keep the total amount in the same font style
        # If you want the amount in bold as well, keep the font as is, otherwise set it back to normal
        self.cell(30, 10, f"${total:,.2f}", 1, 1, 'R')  # The 'R' alignment for total amount

        self.ln(10)  # Add some space after the table for aesthetics

ds = pd.Timestamp.now().strftime('%Y_%m_%d')



def prompt_for_filename(prompt='Enter Filename: '):
    filename = input(prompt)
    return filename


df = pd.read_csv(prompt_for_filename('Enter location of Looker Export: '))
directory = pd.read_csv(prompt_for_filename('Enter location of Affiliate Directory: '))

# Check if df and directory are populated
if df.empty or directory.empty:
    sys.exit("Dataframes df and directory are not populated. Exiting...")

increment_directory = input("Would you like an incremented version of the directory? (y/n): ")


df.Email = df.Email.str.lower()
df.fillna('', inplace=True)
df = df.sort_values('Customer Order No', ascending=True)

df['Commission Rate'] = df['Commission Rate'].str.replace('%', '').astype(float) / 100
df['Commission Amount'] = df['Total Gov'].str.replace('$', '').str.replace(',', '').astype(float) * df['Commission Rate']


directory['Order Email'] = directory['Order Email'].str.lower()
directory.fillna('', inplace=True)

# Fix Hand Carry
df['Shipment Delivery -  Date'].fillna(df['Order Created Date'], inplace=True)
df['Reward Points Redeemed'] = df['Reward Points Redeemed'].str.replace('$', '').str.replace(',', '').astype(float)


# Fix Ship Country Code
df = df[df.groupby('Email')['Ship Country Code'].transform(lambda x: x.eq('US').all())]

# Fix Rewards
df = df[(df['Reward Points Redeemed'] == 0) | \
        ((df['Shipment Delivery -  Date'] <= '2024-01-01') & (df['Reward Points Redeemed'] >= 0))]

# Fix Returns
df = df[df['Item Status'] == 'shipped']

merged_df = df.merge(directory, left_on='Email', right_on='Order Email', suffixes=('', '_directory'))
#merged_df['Email'] = merged_df['Invoice Email']

# Generate Groupings
grouped = merged_df.groupby(['Order Stylist Name', 'Email'])\
    [['Order Stylist Name', 'Email','Customer Order No', 'PID', 'Total Gov','Commission Rate', 'Commission Amount', 'First Name', 'Last Name', 'Address', 'Phone', 'Last Invoice #','DBA', 'Order Email', 'Invoice Email']]

for n, g in grouped:        
    # Create instance of FPDF class
    pdf = PDFInvoice()

    f = g.head(1).to_dict(orient='records')[0]

    title = f['First Name']+'_'+f['Last Name']+ '_'+ str(int(f['Last Invoice #']+1))
    # Add a page
    pdf.add_page()

    # Add invoice data
    pdf.add_invoice_data(pd.Timestamp.now().strftime('%B %d, %Y'), str(int(f['Last Invoice #'])+1), 'Affiliate Commissions')
    
    # Add business info
    pdf.add_business_info(str(f['DBA']), f['Invoice Email'], f['Address'],f['Phone'])
                          
    # Add bill to address
    pdf.add_bill_to("Customer Care Affiliate Commissions\nModa Operandi\n34 34th Street, Building 6, Suite 4A\nBrooklyn, NY 11232")

    # Add items
    items = [
        {'order_number': str(v['Customer Order No']), 'item_number': str(v['PID']), 'sales': str(v['Total Gov']), 'commission': str(v['Commission Rate']), 'amount': str(v['Commission Amount'])} for k,v in g.iterrows()]
        # Add more items...
    pdf.add_items_table(items)

    dir = './invoices_{0}/'.format(ds)+str(n[0]).lower().replace(' ', '_')
    if not os.path.exists(dir):
        os.makedirs(dir)
    

    # Save the pdf with name .pdf
    pdf.output('{0}/{1}.pdf'.format(dir, title))
    if increment_directory == 'n':
        continue
    directory.loc[directory['Order Email'] == f['Order Email'],'Last Invoice #'] = f['Last Invoice #'] + 1
if increment_directory == 'y':
    directory.to_csv('./data/affiliate_directory_{0}_incremented.csv'.format(ds), index=False)