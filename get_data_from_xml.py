import xml.etree.ElementTree as ET
import pandas as pd
import requests

# Reading XML File
file = "Input.xml"
tree = ET.parse(file)
roots = tree.getroot()

data_set = []
for root in roots.findall(".//VOUCHER[@VCHTYPE='Receipt']"):
    try:
        data_set.append({
            "Date": root.find('DATE').text,
            "Vch No.": root.find('VOUCHERNUMBER').text,
            "Particulars": root.find('PARTYLEDGERNAME').text,
            "Amount Verified": root.find("AUDITED").text,
            "Debtor": root.find('PARTYLEDGERNAME').text,
            "Vch Type": root.find("VOUCHERTYPENAME").text,
            "Ref No": root.find("REFERENCE").text,
            "Amount": root.find("AMOUNT").text,
            "Ref Amount": root.find("VATEXPAMOUNT").text,
            "ref_date": root.find("REFERENCEDATE").text,
        })
    except AttributeError:
        data_set.append({
            "Date": root.find('DATE').text,
            "Vch No.": root.find('VOUCHERNUMBER').text,
            "Particulars": root.find('PARTYLEDGERNAME').text,
            "Amount Verified": root.find("AUDITED").text,
            "Debtor": root.find('PARTYLEDGERNAME').text,
            "Vch Type": root.find("VOUCHERTYPENAME").text,
            "Ref No": root.find("REFERENCE"),
            "Amount": root.find("AMOUNT"),
            "Ref Amount": root.find("VATEXPAMOUNT"),
            "ref_date": root.find("REFERENCEDATE"),
        })

# Creating Xlsx File From Dataframe
dataframe = pd.DataFrame(data_set)
dataframe.to_excel('output.xlsx')


# Creating Google sheet File using Sheety Api
# API ENDPOINT
sheet_endpoint = "https://api.sheety.co/6c47ce8ae18b53664d39007a9924b720/data/sheet1"

for root in roots.findall(".//VOUCHER[@VCHTYPE='Receipt']"):
    try:
        sheet_input = {
            "sheet_1":
            {
                "Date": root.find('DATE').text,
                "Vch No.": root.find('VOUCHERNUMBER').text,
                "Particulars": root.find('PARTYLEDGERNAME').text,
                "Amount Verified": root.find("AUDITED").text,
                "Debtor": root.find('PARTYLEDGERNAME').text,
                "Vch Type": root.find("VOUCHERTYPENAME").text,
                "Ref No": root.find("REFERENCE").text,
                "Amount": root.find("AMOUNT").text,
                "Ref Amount": root.find("VATEXPAMOUNT").text,
                "ref_date": root.find("REFERENCEDATE").text,
            }
        }
    except AttributeError:
        sheet_input = {
            "sheet_1":
                {
                    "Date": root.find('DATE').text,
                    "Vch No.": root.find('VOUCHERNUMBER').text,
                    "Particulars": root.find('PARTYLEDGERNAME').text,
                    "Amount Verified": root.find("AUDITED").text,
                    "Debtor": root.find('PARTYLEDGERNAME').text,
                    "Vch Type": root.find("VOUCHERTYPENAME").text,
                    "Ref No": root.find("REFERENCE"),
                    "Amount": root.find("AMOUNT"),
                    "Ref Amount": root.find("VATEXPAMOUNT"),
                    "ref_date": root.find("REFERENCEDATE"),
                }
        }
    # Sending API Request
    sheet_response = requests.post(sheet_endpoint, json=sheet_input)
