import xml.etree.ElementTree as ET
import pandas as pd
import requests

# Reading XML File
file = "Input.xml"
tree = ET.parse(file)
roots = tree.getroot()

# Creating Array For Each Column
date = []
company_name = []
debtor = []
voucher_number = []
vcr_type = []
amount_verified = []
amount = []
actual_amount = []
ref_no = []
ref_date = []

# Loop TO get Hold Of each Attribute Text. Where Voucher Type is Receipt
for root in roots.findall(".//VOUCHER[@VCHTYPE='Receipt']"):
    try:
        date.append(root.find('DATE').text)
        company_name.append(root.find('PARTYLEDGERNAME').text)
        debtor.append(root.find('PARTYLEDGERNAME').text)
        voucher_number.append(root.find('VOUCHERNUMBER').text)
        amount_verified.append(root.find("AUDITED").text)
        vcr_type.append(root.find("VOUCHERTYPENAME").text)
        amount.append(root.find("AMOUNT").text)
        actual_amount.append(root.find("VATEXPAMOUNT").text)
        ref_no.append(root.find("REFERENCE").text)
        ref_date.append(root.find("REFERENCEDATE").text)
    except AttributeError:
        amount.append("Na")
        actual_amount.append("Na")
        ref_no.append("Na")
        ref_date.append("Na")

# Creating DataFrame Using Pandas
data = {'Date': date, 'Vch No.': voucher_number, 'Ref No': ref_no, 'Ref Date': ref_date, 'Debtor': debtor,
        "Ref Amount": amount, "Amount": actual_amount, "Particulars": company_name, "Vch Type": vcr_type,
        "Amount Verified": amount_verified}
# Creating Xlsx File From Dataframe
dataframe = pd.DataFrame(data)
dataframe.to_excel('output.xlsx')

# Creating Google sheet File using Sheety Api
# API ENDPOINT
sheet_endpoint = "https://api.sheety.co/6c47ce8ae18b53664d39007a9924b720/data/sheet1"

# SENDING EACH ARRAY AS COLUMN VIA POST Request
sheet_input = {
    "sheet1": {
        "Date": date,
        "Vch No.": voucher_number,
        "Ref No": ref_no,
        "Ref Date": ref_date,
        "Debtor": debtor,
        "Ref Amount": amount,
        "Amount": actual_amount,
        "Particulars": company_name,
        "Vch Type": vcr_type,
        "Amount Verified": amount_verified
    }
}
# Sending API Request
sheet_response = requests.post(sheet_endpoint, json=sheet_input)
print(sheet_response.text)