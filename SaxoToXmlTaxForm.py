import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from bisect import bisect_left

def parse_exchange_rates(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    ns = {"ns": "http://www.ecb.europa.eu/vocabulary/stats/exr/1"}
    exchange_rates = {}
    
    for obs in root.findall(".//ns:Obs", ns):
        date_str = obs.get("TIME_PERIOD")
        rate = float(obs.get("OBS_VALUE"))
        exchange_rates[date_str] = rate
    
    sorted_dates = sorted(exchange_rates.keys())
    return exchange_rates, sorted_dates

def get_nearest_rate(exchange_rates, sorted_dates, target_date):
    if target_date in exchange_rates:
        return exchange_rates[target_date]
    pos = bisect_left(sorted_dates, target_date)
    if pos == 0:
        return exchange_rates[sorted_dates[0]]
    else:
        return exchange_rates[sorted_dates[pos - 1]]

def format_date(date_value):
    if isinstance(date_value, str):
        try:
            return datetime.strptime(date_value, "%d-%b-%Y").strftime("%Y-%m-%d")
        except ValueError:
            return date_value
    return date_value.strftime("%Y-%m-%d") if isinstance(date_value, datetime) else date_value

def generate_xml_from_excel(file_path, exchange_rate_file, taxpayer_info, kdvp_info):
    df = pd.read_excel(file_path, header=None)
    exchange_rates, sorted_dates = parse_exchange_rates(exchange_rate_file)

    envelope = ET.Element("Envelope", {
        "xmlns": "http://edavki.durs.si/Documents/Schemas/Doh_KDVP_9.xsd",
        "xmlns:edp": "http://edavki.durs.si/Documents/Schemas/EDP-Common-1.xsd"
    })
    
    header = ET.SubElement(envelope, "edp:Header")
    taxpayer = ET.SubElement(header, "edp:taxpayer")
    ET.SubElement(taxpayer, "edp:taxNumber").text = taxpayer_info["taxNumber"]
    ET.SubElement(taxpayer, "edp:taxpayerType").text = taxpayer_info["taxpayerType"]
    ET.SubElement(taxpayer, "edp:name").text = taxpayer_info["name"]
    ET.SubElement(taxpayer, "edp:address1").text = taxpayer_info["address1"]
    ET.SubElement(taxpayer, "edp:city").text = taxpayer_info["city"]
    ET.SubElement(taxpayer, "edp:postNumber").text = taxpayer_info["postNumber"]
    ET.SubElement(taxpayer, "edp:birthDate").text = taxpayer_info["birthDate"]
    
    body = ET.SubElement(envelope, "body")
    doh_kdvp = ET.SubElement(body, "Doh_KDVP")
    kdvp = ET.SubElement(doh_kdvp, "KDVP")
    
    ET.SubElement(kdvp, "DocumentWorkflowID").text = "I"
    ET.SubElement(kdvp, "Year").text = kdvp_info["Year"]
    ET.SubElement(kdvp, "PeriodStart").text = kdvp_info["PeriodStart"]
    ET.SubElement(kdvp, "PeriodEnd").text = kdvp_info["PeriodEnd"]
    ET.SubElement(kdvp, "IsResident").text = "true"
    ET.SubElement(kdvp, "SecurityCount").text = str(len(df[6].unique()))
    
    item_id = 1
    for _, row in df.iterrows():
        if row.isnull().all():
            break
        date_open = format_date(row[1])
        date_close = format_date(row[0])
        rate_open = get_nearest_rate(exchange_rates, sorted_dates, date_open)
        rate_close = get_nearest_rate(exchange_rates, sorted_dates, date_close)
        price_open_eur = row[12] / rate_open
        price_close_eur = row[13] / rate_close
        
        kdvp_item = ET.SubElement(doh_kdvp, "KDVPItem")
        ET.SubElement(kdvp_item, "ItemID").text = str(item_id)
        ET.SubElement(kdvp_item, "InventoryListType").text = "PLVP"
        ET.SubElement(kdvp_item, "Name").text = row[6]
        
        securities = ET.SubElement(kdvp_item, "Securities")
        ET.SubElement(securities, "Code").text = row[6]
        
        row_purchase = ET.SubElement(securities, "Row")
        ET.SubElement(row_purchase, "ID").text = "0"
        purchase = ET.SubElement(row_purchase, "Purchase")
        ET.SubElement(purchase, "F1").text = date_open
        ET.SubElement(purchase, "F3").text = f"{abs(row[11]):.4f}"
        ET.SubElement(purchase, "F4").text = f"{price_open_eur:.4f}"
        
        row_sale = ET.SubElement(securities, "Row")
        ET.SubElement(row_sale, "ID").text = "2"
        sale = ET.SubElement(row_sale, "Sale")
        ET.SubElement(sale, "F6").text = date_close
        ET.SubElement(sale, "F7").text = f"{abs(row[11]):.4f}"
        ET.SubElement(sale, "F9").text = f"{price_close_eur:.4f}"
        
        item_id += 1
    
    tree = ET.ElementTree(envelope)
    tree.write("output.xml", encoding="utf-8", xml_declaration=True)

taxpayer_info = {
    "taxNumber": input("Enter tax number: "),
    "taxpayerType": input("Enter taxpayer type (FO or other): "),
    "name": input("Enter name: "),
    "address1": input("Enter address: "),
    "city": input("Enter city: "),
    "postNumber": input("Enter post number: "),
    "birthDate": input("Enter birth date (YYYY-MM-DD): ")
}

kdvp_info = {
    "Year": input("Enter year: "),
    "PeriodStart": input("Enter period start date (YYYY-MM-DD): "),
    "PeriodEnd": input("Enter period end date (YYYY-MM-DD): ")
}

generate_xml_from_excel("taxexample.xlsx", "usdtoeur.xml", taxpayer_info, kdvp_info)
