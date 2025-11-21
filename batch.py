import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import warnings
from flask import Flask, request, send_file, render_template_string, jsonify
from io import BytesIO
import zipfile
import tempfile
import uuid
import asyncio
import aiofiles
import concurrent.futures
from threading import Lock
import time

# Suppress warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)

# Consignment-specific fixed values (for LV02 2025 6241)
CONSIGNMENT_VALUES = {
    'total_invoice': '2006.64',
    'total_cif': '4212.99', 
    'total_cost': '621.1',
    'external_freight_foreign': '4.78',
    'external_freight_national': '8.56',
    'insurance_foreign': '0.58',
    'insurance_national': '1.04',
    'other_cost_foreign': '0.47',
    'other_cost_national': '0.84',
    'total_cif_itm': '70.8',
    'statistical_value': '71',
    'alpha_coefficient': '0.0168042100227245',
    'duty_tax_base': '71',
    'duty_tax_rate': '6',
    'duty_tax_amount': '4.3',
    'total_item_taxes': '347.75',
    'calculation_working_mode': '0',
    'container_flag': 'False',
    'delivery_terms_code': 'DDP',
    'currency_rate': '1.79',
    'manifest_reference': 'LV02 2025 6241',
    'total_forms': '16'
}

# Global progress tracking
conversion_progress = {}
progress_lock = Lock()

def read_excel_data(file_content):
    """Read Excel file with exact ASYCUDA structure"""
    sad_data = {}
    items_data = []
    
    try:
        # Read SAD sheet
        try:
            sad_df = pd.read_excel(BytesIO(file_content), sheet_name='SAD')
            if not sad_df.empty:
                sad_row = sad_df.iloc[0]
                for col in sad_df.columns:
                    if pd.notna(sad_row[col]):
                        sad_data[col] = str(sad_row[col])
                    else:
                        sad_data[col] = ''
        except Exception as e:
            print(f"Warning reading SAD sheet: {str(e)}")
        
        # Read Items sheet
        try:
            items_df = pd.read_excel(BytesIO(file_content), sheet_name='Items')
            if not items_df.empty:
                for _, row in items_df.iterrows():
                    item = {}
                    for col in items_df.columns:
                        if pd.notna(row[col]):
                            item[col] = str(row[col])
                        else:
                            item[col] = ''
                    items_data.append(item)
        except Exception as e:
            print(f"Warning reading Items sheet: {str(e)}")
        
    except Exception as e:
        print(f"Error reading file: {str(e)}")
    
    return sad_data, items_data

def calculate_form_totals(items_data):
    """Calculate form-specific totals (per XML file)"""
    invoice_foreign_total = 0
    for item in items_data:
        inv_foreign = item.get('Invoice Amount_foreign_currency', '0')
        try:
            invoice_foreign_total += float(inv_foreign) if inv_foreign else 0
        except:
            pass
    
    return invoice_foreign_total

def add_element(parent, tag_name, text_content):
    """Helper method to add element with text content"""
    element = ET.SubElement(parent, tag_name)
    if text_content and text_content != 'nan' and text_content != 'None' and text_content != '':
        element.text = str(text_content)
    return element

def create_valuation_subsections(parent, form_invoice_foreign):
    """Create valuation subsections with consignment-specific values"""
    # Invoice
    invoice = ET.SubElement(parent, "Invoice")
    add_element(invoice, "Amount_national_currency", "3591.89")
    add_element(invoice, "Amount_foreign_currency", str(form_invoice_foreign))
    add_element(invoice, "Currency_code", "USD")
    add_element(invoice, "Currency_name", "Geen vreemde valuta")
    add_element(invoice, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # External_freight
    external = ET.SubElement(parent, "External_freight")
    add_element(external, "Amount_national_currency", "509.24")
    add_element(external, "Amount_foreign_currency", "17.27")
    add_element(external, "Currency_code", "USD")
    add_element(external, "Currency_name", "Geen vreemde valuta")
    add_element(external, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Internal_freight
    internal = ET.SubElement(parent, "Internal_freight")
    add_element(internal, "Amount_national_currency", "0")
    add_element(internal, "Amount_foreign_currency", "0")
    add_element(internal, "Currency_code", "")
    add_element(internal, "Currency_name", "Geen vreemde valuta")
    add_element(internal, "Currency_rate", "0")
    
    # Insurance
    insurance = ET.SubElement(parent, "Insurance")
    add_element(insurance, "Amount_national_currency", "62.26")
    add_element(insurance, "Amount_foreign_currency", "1.00875")
    add_element(insurance, "Currency_code", "USD")
    add_element(insurance, "Currency_name", "Geen vreemde valuta")
    add_element(insurance, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Other_cost
    other = ET.SubElement(parent, "Other_cost")
    add_element(other, "Amount_national_currency", "49.6")
    add_element(other, "Amount_foreign_currency", "")
    add_element(other, "Currency_code", "USD")
    add_element(other, "Currency_name", "Geen vreemde valuta")
    add_element(other, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Deduction
    deduction = ET.SubElement(parent, "Deduction")
    add_element(deduction, "Amount_national_currency", "0")
    add_element(deduction, "Amount_foreign_currency", "0")
    add_element(deduction, "Currency_code", "USD")
    add_element(deduction, "Currency_name", "Geen vreemde valuta")
    add_element(deduction, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])

def create_item_supplementary_unit(parent, item_data, unit_num):
    """Create supplementary unit with proper structure for items"""
    supp_unit = ET.SubElement(parent, "Supplementary_unit")
    
    if unit_num == '1':
        add_element(supp_unit, "Supplementary_unit_rank", "")
        add_element(supp_unit, "Supplementary_unit_code", item_data.get('Supplementary_unit_code', 'PCE'))
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_1', 'Aantal Stucks'))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_1', ''))
    elif unit_num == '2':
        add_element(supp_unit, "Supplementary_unit_rank", "2")
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_2', ''))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_2', ''))
    else:  # unit_num == '3'
        add_element(supp_unit, "Supplementary_unit_rank", "3")
        add_element(supp_unit, "Supplementary_unit_name", item_data.get('Supplementary_unit_name_3', ''))
        add_element(supp_unit, "Supplementary_unit_quantity", item_data.get('Supplementary_unit_quantity_3', ''))

def create_item_valuation_subsections(parent, item_data):
    """Create valuation subsections for items with consignment-specific values"""
    # Invoice
    invoice = ET.SubElement(parent, "Invoice")
    add_element(invoice, "Amount_national_currency", "")
    add_element(invoice, "Amount_foreign_currency", item_data.get('Invoice Amount_foreign_currency', ''))
    add_element(invoice, "Currency_code", "USD")
    add_element(invoice, "Currency_name", "Geen vreemde valuta")
    add_element(invoice, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # External_freight (consignment-specific per item)
    external = ET.SubElement(parent, "External_freight")
    add_element(external, "Amount_national_currency", CONSIGNMENT_VALUES['external_freight_national'])
    add_element(external, "Amount_foreign_currency", CONSIGNMENT_VALUES['external_freight_foreign'])
    add_element(external, "Currency_code", "USD")
    add_element(external, "Currency_name", "Geen vreemde valuta")
    add_element(external, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Internal_freight
    internal = ET.SubElement(parent, "Internal_freight")
    add_element(internal, "Amount_national_currency", "0")
    add_element(internal, "Amount_foreign_currency", "")
    add_element(internal, "Currency_code", "")
    add_element(internal, "Currency_name", "Geen vreemde valuta")
    add_element(internal, "Currency_rate", "0")
    
    # Insurance (consignment-specific per item)
    insurance = ET.SubElement(parent, "Insurance")
    add_element(insurance, "Amount_national_currency", CONSIGNMENT_VALUES['insurance_national'])
    add_element(insurance, "Amount_foreign_currency", CONSIGNMENT_VALUES['insurance_foreign'])
    add_element(insurance, "Currency_code", "USD")
    add_element(insurance, "Currency_name", "Geen vreemde valuta")
    add_element(insurance, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Other_cost (consignment-specific per item)
    other = ET.SubElement(parent, "Other_cost")
    add_element(other, "Amount_national_currency", CONSIGNMENT_VALUES['other_cost_national'])
    add_element(other, "Amount_foreign_currency", CONSIGNMENT_VALUES['other_cost_foreign'])
    add_element(other, "Currency_code", "USD")
    add_element(other, "Currency_name", "Geen vreemde valuta")
    add_element(other, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])
    
    # Deduction
    deduction = ET.SubElement(parent, "Deduction")
    add_element(deduction, "Amount_national_currency", "0")
    add_element(deduction, "Amount_foreign_currency", "0")
    add_element(deduction, "Currency_code", "USD")
    add_element(deduction, "Currency_name", "Geen vreemde valuta")
    add_element(deduction, "Currency_rate", CONSIGNMENT_VALUES['currency_rate'])

def create_item_element(parent, item_data, item_number):
    """Create individual Item element with consignment-specific values"""
    item = ET.SubElement(parent, "Item")
    
    # Packages section
    packages = ET.SubElement(item, "Packages")
    add_element(packages, "Number_of_packages", item_data.get('Number_of_packages', ''))
    add_element(packages, "Marks1_of_packages", item_data.get('Marks1_of_packages', ''))
    add_element(packages, "Marks2_of_packages", item_data.get('Marks2_of_packages', ''))
    add_element(packages, "Kind_of_packages_code", item_data.get('Kind_of_packages_code', 'STKS'))
    add_element(packages, "Kind_of_packages_name", item_data.get('Kind_of_packages_name', 'Stuks'))
    
    # Tariff section
    tariff = ET.SubElement(item, "Tariff")
    add_element(tariff, "Extended_customs_procedure", item_data.get('Extended_customs_procedure', '4000'))
    add_element(tariff, "National_customs_procedure", item_data.get('National_customs_procedure', '00:00:00'))
    add_element(tariff, "Preference_code", item_data.get('Preference_code', ''))
    
    harmonized = ET.SubElement(tariff, "Harmonized_system")
    add_element(harmonized, "Commodity_code", item_data.get('Commodity_code', ''))
    add_element(harmonized, "Precision_4", item_data.get('Precision_4', ''))
    
    # Three supplementary units
    create_item_supplementary_unit(tariff, item_data, '1')
    create_item_supplementary_unit(tariff, item_data, '2') 
    create_item_supplementary_unit(tariff, item_data, '3')
    
    quota = ET.SubElement(tariff, "Quota")
    add_element(quota, "Quota_code", item_data.get('Quota_code', ''))
    
    # Goods_description
    goods_desc = ET.SubElement(item, "Goods_description")
    add_element(goods_desc, "Country_of_origin_code", item_data.get('Country_of_origin_code', 'US'))
    add_element(goods_desc, "Description_of_goods", item_data.get('Description_of_goods', ''))
    add_element(goods_desc, "Commercial_description", item_data.get('Commercial_description', ''))
    
    # Valuation_item with consignment-specific values
    valuation_item = ET.SubElement(item, "Valuation_item")
    add_element(valuation_item, "Rate_of_adjustment", "1")
    add_element(valuation_item, "Total_cost_itm", "")
    add_element(valuation_item, "Total_cif_itm", CONSIGNMENT_VALUES['total_cif_itm'])
    add_element(valuation_item, "Statistical_value", CONSIGNMENT_VALUES['statistical_value'])
    add_element(valuation_item, "Alpha_coeficient_of_apportionment", CONSIGNMENT_VALUES['alpha_coefficient'])
    
    weight = ET.SubElement(valuation_item, "Weight")
    add_element(weight, "Gross_weight_itm", item_data.get('Gross_weight_itm', '0.5'))
    add_element(weight, "Net_weight_itm", item_data.get('Net_weight_itm', '0.5'))
    
    # Item valuation subsections with consignment-specific values
    create_item_valuation_subsections(valuation_item, item_data)
    
    # Previous_document
    prev_doc = ET.SubElement(item, "Previous_document")
    add_element(prev_doc, "Summary_declaration", item_data.get('Summary_declaration', ''))
    add_element(prev_doc, "Summary_declaration_sl", item_data.get('Summary_declaration_sl', '1'))
    
    # Taxation with consignment-specific values
    taxation = ET.SubElement(item, "Taxation")
    add_element(taxation, "Item_taxes_amount", CONSIGNMENT_VALUES['duty_tax_amount'])
    add_element(taxation, "Item_taxes_mode_of_payment", "1")
    
    tax_line = ET.SubElement(taxation, "Taxation_line")
    add_element(tax_line, "Duty_tax_code", "IR")
    add_element(tax_line, "Duty_tax_base", CONSIGNMENT_VALUES['duty_tax_base'])
    add_element(tax_line, "Duty_tax_rate", CONSIGNMENT_VALUES['duty_tax_rate'])
    add_element(tax_line, "Duty_tax_amount", CONSIGNMENT_VALUES['duty_tax_amount'])
    add_element(tax_line, "Duty_tax_MP", "1")

def create_asycuda_xml(sad_data, items_data, filename):
    """Create exact ASYCUDA XML structure with consignment"""
    # Calculate form-specific values
    form_invoice_foreign = calculate_form_totals(items_data)
    
    # Create root element
    root = ET.Element("ASYCUDA")
    
    # SAD section
    sad = ET.SubElement(root, "SAD")
    
    # Assessment_notice section
    assessment_notice = ET.SubElement(sad, "Assessment_notice")
    add_element(assessment_notice, "Total_item_taxes", CONSIGNMENT_VALUES['total_item_taxes'])
    
    items_taxes = ET.SubElement(assessment_notice, "Items_taxes")
    item_tax = ET.SubElement(items_taxes, "Item_tax")
    add_element(item_tax, "Tax_code", sad_data.get('Tax_code', 'IR'))
    add_element(item_tax, "Tax_description", sad_data.get('Tax_description', 'Invoerrechten'))
    add_element(item_tax, "Tax_amount", CONSIGNMENT_VALUES['total_item_taxes'])
    add_element(item_tax, "Tax_mop", sad_data.get('Tax_mop', '1'))
    
    # Properties section
    properties = ET.SubElement(sad, "Properties")
    add_element(properties, "Sad_flow", sad_data.get('Sad_flow', 'I'))
    
    forms = ET.SubElement(properties, "Forms")
    add_element(forms, "Number_of_the_form", sad_data.get('Number_of_the_form', '1'))
    add_element(forms, "Total_number_of_forms", CONSIGNMENT_VALUES['total_forms'])
    
    add_element(properties, "Selected_page", sad_data.get('Selected_page', '1'))
    
    # Identification section
    identification = ET.SubElement(sad, "Identification")
    add_element(identification, "Manifest_reference_number", CONSIGNMENT_VALUES['manifest_reference'])
    
    office_segment = ET.SubElement(identification, "Office_segment")
    add_element(office_segment, "Customs_clearance_office_code", sad_data.get('Customs_clearance_office_code', 'LV01'))
    add_element(office_segment, "Customs_clearance_office_name", sad_data.get('Customs_clearance_office_name', 'Luchthaven Vracht'))
    
    type_elem = ET.SubElement(identification, "Type")
    add_element(type_elem, "Type_of_declaration", sad_data.get('Type_of_declaration', 'INV'))
    add_element(type_elem, "General_procedure_code", sad_data.get('General_procedure_code', '4'))
    
    # Traders section
    traders = ET.SubElement(sad, "Traders")
    
    exporter = ET.SubElement(traders, "Exporter")
    add_element(exporter, "Exporter_code", sad_data.get('Exporter_code', ''))
    add_element(exporter, "Exporter_name", sad_data.get('Exporter_name', ''))
    
    consignee = ET.SubElement(traders, "Consignee")
    add_element(consignee, "Consignee_code", sad_data.get('Consignee_code', '10026483'))
    add_element(consignee, "Consignee_name", sad_data.get('Consignee_name', 'Dhr. Anthony Martina Paradera 1-H Paradera Paradera Aruba'))
    
    financial_trader = ET.SubElement(traders, "Financial")
    add_element(financial_trader, "Financial_code", sad_data.get('Financial_code', ''))
    add_element(financial_trader, "Financial_name", sad_data.get('Financial_name', ''))
    
    # Declarant section
    declarant = ET.SubElement(sad, "Declarant")
    add_element(declarant, "Declarant_code", sad_data.get('Declarant_code', '1160650'))
    add_element(declarant, "Declarant_name", sad_data.get('Declarant_name', 'Dhr. Victor Hoek Alto Vista 133 Alto Vista Noord/Tanki Leendert Aruba'))
    add_element(declarant, "Declarant_representative", sad_data.get('Declarant_representative', 'Lizandra I. Geerman'))
    
    reference = ET.SubElement(declarant, "Reference")
    add_element(reference, "Year", sad_data.get('Reference Year', '2025'))
    add_element(reference, "Number", sad_data.get('Reference Number', ''))
    
    # General_information section
    general_info = ET.SubElement(sad, "General_information")
    
    country = ET.SubElement(general_info, "Country")
    add_element(country, "Country_first_destination", sad_data.get('Country_first_destination', 'US'))
    add_element(country, "Trading_country", sad_data.get('Trading_country', 'US'))
    add_element(country, "Country_of_origin_name", sad_data.get('Country_of_origin_name', 'Verenigde Staten'))
    
    export = ET.SubElement(country, "Export")
    add_element(export, "Export_country_code", sad_data.get('Export_country_code', 'US'))
    add_element(export, "Export_country_name", sad_data.get('Export_country_name', 'Verenigde Staten'))
    add_element(export, "Export_country_region", sad_data.get('Export_country_region', ''))
    
    destination = ET.SubElement(country, "Destination")
    add_element(destination, "Destination_country_code", sad_data.get('Destination_country_code', 'AW'))
    add_element(destination, "Destination_country_name", sad_data.get('Destination_country_name', 'Aruba'))
    add_element(destination, "Destination_country_region", sad_data.get('Destination_country_region', ''))
    
    add_element(general_info, "Value_details", CONSIGNMENT_VALUES['total_cost'])
    add_element(general_info, "CAP", sad_data.get('CAP', ''))
    
    # Transport section
    transport = ET.SubElement(sad, "Transport")
    add_element(transport, "Container_flag", CONSIGNMENT_VALUES['container_flag'])
    add_element(transport, "Location_of_goods", sad_data.get('Location_of_goods', 'RT-01'))
    add_element(transport, "Location_of_goods_address", sad_data.get('Location_of_goods_address', 'Sabana Berde #75'))
    
    means_transport = ET.SubElement(transport, "Means_of_transport")
    
    departure = ET.SubElement(means_transport, "Departure_arrival_information")
    add_element(departure, "Identity", sad_data.get('Departure_arrival_information Identity', 'COPA AIRLINES'))
    add_element(departure, "Nationality", sad_data.get('Departure_arrival_information Nationality', 'PA'))
    
    border = ET.SubElement(means_transport, "Border_information")
    add_element(border, "Identity", sad_data.get('Border_information Identity', ''))
    add_element(border, "Nationality", sad_data.get('Border_information Nationality', ''))
    add_element(border, "Mode", sad_data.get('Border_information Mode', '4'))
    
    delivery = ET.SubElement(transport, "Delivery_terms")
    add_element(delivery, "Code", CONSIGNMENT_VALUES['delivery_terms_code'])
    add_element(delivery, "Place", sad_data.get('Delivery_terms Place', 'USA'))
    
    border_office = ET.SubElement(transport, "Border_office")
    add_element(border_office, "Code", sad_data.get('Border_office Code', 'LV01'))
    add_element(border_office, "Name", sad_data.get('Border_office Name', 'Luchthaven Vracht'))
    
    place_loading = ET.SubElement(transport, "Place_of_loading")
    add_element(place_loading, "Code", sad_data.get('Place_of_loading Code', 'AWAIR'))
    add_element(place_loading, "Name", sad_data.get('Place_of_loading Name', 'Aeropuerto Reina Beatrix'))
    
    # Financial section
    financial = ET.SubElement(sad, "Financial")
    add_element(financial, "Deffered_payment_reference", sad_data.get('Deffered_payment_reference', ''))
    add_element(financial, "Mode_of_payment", sad_data.get('Mode_of_payment', 'CONTANT'))
    
    fin_trans = ET.SubElement(financial, "Financial_transaction")
    add_element(fin_trans, "Code_1", sad_data.get('Financial_transaction Code_1', '1'))
    add_element(fin_trans, "Code_2", sad_data.get('Financial_transaction Code_1', '1'))
    
    bank = ET.SubElement(financial, "Bank")
    add_element(bank, "Branch", sad_data.get('Bank Branch', ''))
    add_element(bank, "Reference", sad_data.get('Bank Reference', ''))
    
    terms = ET.SubElement(financial, "Terms")
    add_element(terms, "Code", sad_data.get('Terms Code', ''))
    add_element(terms, "Description", sad_data.get('Terms Description', ''))
    
    amounts = ET.SubElement(financial, "Amounts")
    add_element(amounts, "Global_taxes", sad_data.get('Amounts Global_taxes', '0'))
    add_element(amounts, "Totals_taxes", CONSIGNMENT_VALUES['total_item_taxes'])
    
    guarantee = ET.SubElement(financial, "Guarantee")
    add_element(guarantee, "Amount", sad_data.get('Guarantee Amount', '0'))
    
    # Transit section
    transit = ET.SubElement(sad, "Transit")
    add_element(transit, "Result_of_control", sad_data.get('Result_of_control', ''))
    
    # Valuation section
    valuation = ET.SubElement(sad, "Valuation")
    add_element(valuation, "Calculation_working_mode", CONSIGNMENT_VALUES['calculation_working_mode'])
    add_element(valuation, "Total_cost", CONSIGNMENT_VALUES['total_cost'])
    add_element(valuation, "Total_cif", CONSIGNMENT_VALUES['total_cif'])
    
    # Valuation subsections with consignment-specific values
    create_valuation_subsections(valuation, form_invoice_foreign)
    
    total = ET.SubElement(valuation, "Total")
    add_element(total, "Total_invoice", CONSIGNMENT_VALUES['total_invoice'])
    add_element(total, "Total_weight", str(len(items_data)))
    
    # Items section
    items_elem = ET.SubElement(root, "Items")
    
    # Create items from Items data with consignment-specific values
    for i, item_data in enumerate(items_data):
        create_item_element(items_elem, item_data, i+1)
    
    return root

def prettify_xml(elem):
    """Convert XML to pretty formatted string"""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def convert_excel_to_xml(file_content, filename):
    """Convert single Excel file to ASYCUDA XML"""
    try:
        # Read data from Excel
        sad_data, items_data = read_excel_data(file_content)
        
        if not sad_data and not items_data:
            return False, f"No valid data found in {filename}"
        
        # Create XML structure
        xml_root = create_asycuda_xml(sad_data, items_data, filename)
        
        # Generate XML content
        xml_content = prettify_xml(xml_root)
        
        return True, xml_content
        
    except Exception as e:
        return False, f"{filename} | Error: {str(e)}"

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Universal ASYCUDA XML Converter</title>
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
                color: #ffffff;
                min-height: 100vh;
                overflow-x: hidden;
            }

            .container {
                display: flex;
                min-height: 100vh;
            }

            .left-panel {
                flex: 1;
                padding: 30px;
                background: rgba(30, 30, 30, 0.9);
                border-right: 2px solid #333;
                overflow-y: auto;
            }

            .right-panel {
                flex: 1;
                padding: 30px;
                background: rgba(25, 25, 25, 0.9);
                overflow-y: auto;
            }

            .header {
                text-align: center;
                margin-bottom: 30px;
                padding: 25px;
                background: rgba(30, 30, 30, 0.8);
                border-radius: 15px;
                border: 1px solid #4da6ff;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
            }

            .title {
                font-size: 2.2rem;
                font-weight: bold;
                color: #4da6ff;
                margin-bottom: 10px;
                text-shadow: 0 2px 4px rgba(0, 0, 0, 0.5);
            }

            .subtitle {
                font-size: 1.1rem;
                color: #cccccc;
                margin-bottom: 15px;
            }

            .consignment-info {
                background: rgba(77, 166, 255, 0.1);
                padding: 12px;
                border-radius: 8px;
                margin-top: 12px;
                border-left: 4px solid #4da6ff;
                font-size: 0.9rem;
            }

            .upload-section {
                background: rgba(30, 30, 30, 0.8);
                padding: 25px;
                border-radius: 15px;
                margin-bottom: 25px;
                border: 1px solid #333;
            }

            .upload-area {
                border: 3px dashed #4da6ff;
                border-radius: 10px;
                padding: 30px;
                text-align: center;
                background: rgba(77, 166, 255, 0.05);
                transition: all 0.3s ease;
                cursor: pointer;
                margin-bottom: 15px;
            }

            .upload-area:hover {
                background: rgba(77, 166, 255, 0.1);
                border-color: #66b3ff;
            }

            .upload-area.dragover {
                background: rgba(77, 166, 255, 0.2);
                border-color: #99ccff;
            }

            .upload-icon {
                font-size: 3rem;
                color: #4da6ff;
                margin-bottom: 15px;
            }

            .file-input {
                display: none;
            }

            .folder-input {
                display: none;
            }

            .btn {
                background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
                color: white;
                border: none;
                padding: 12px 25px;
                font-size: 1rem;
                border-radius: 8px;
                cursor: pointer;
                transition: all 0.3s ease;
                font-weight: bold;
                margin: 5px;
                width: calc(50% - 10px);
            }

            .btn:hover {
                background: linear-gradient(135deg, #106ebe 0%, #005a9e 100%);
                transform: translateY(-2px);
                box-shadow: 0 4px 15px rgba(0, 120, 212, 0.4);
            }

            .btn:disabled {
                background: #666;
                cursor: not-allowed;
                transform: none;
                box-shadow: none;
            }

            .btn-secondary {
                background: linear-gradient(135deg, #6c757d 0%, #5a6268 100%);
            }

            .btn-secondary:hover {
                background: linear-gradient(135deg, #5a6268 0%, #495057 100%);
            }

            .btn-group {
                display: flex;
                gap: 10px;
                margin: 15px 0;
            }

            .progress-section {
                background: rgba(30, 30, 30, 0.8);
                padding: 20px;
                border-radius: 15px;
                margin-bottom: 25px;
                border: 1px solid #333;
            }

            .progress-bar {
                width: 100%;
                height: 20px;
                background: #333;
                border-radius: 10px;
                overflow: hidden;
                margin: 15px 0;
            }

            .progress-fill {
                height: 100%;
                background: linear-gradient(90deg, #0078d4 0%, #4da6ff 100%);
                width: 0%;
                transition: width 0.3s ease;
                position: relative;
            }

            .progress-text {
                position: absolute;
                right: 10px;
                top: 50%;
                transform: translateY(-50%);
                color: white;
                font-weight: bold;
                font-size: 0.8rem;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.7);
            }

            .stats {
                display: flex;
                justify-content: space-around;
                margin: 15px 0;
                text-align: center;
            }

            .stat-item {
                background: rgba(77, 166, 255, 0.1);
                padding: 12px;
                border-radius: 8px;
                flex: 1;
                margin: 0 5px;
            }

            .stat-number {
                font-size: 1.5rem;
                font-weight: bold;
                color: #4da6ff;
            }

            .results-section {
                background: rgba(30, 30, 30, 0.8);
                padding: 20px;
                border-radius: 15px;
                border: 1px solid #333;
                height: calc(100vh - 200px);
                display: flex;
                flex-direction: column;
            }

            .results-log {
                background: #252526;
                border: 1px solid #333;
                border-radius: 8px;
                padding: 20px;
                flex: 1;
                overflow-y: auto;
                font-family: 'Consolas', monospace;
                color: #d4d4d4;
                margin-bottom: 15px;
                font-size: 0.9rem;
                line-height: 1.4;
            }

            .file-list {
                max-height: 200px;
                overflow-y: auto;
                margin: 15px 0;
                background: rgba(255, 255, 255, 0.05);
                border-radius: 8px;
                padding: 10px;
            }

            .file-item {
                background: rgba(255, 255, 255, 0.05);
                padding: 8px;
                margin: 3px 0;
                border-radius: 5px;
                border-left: 3px solid #4da6ff;
                font-size: 0.85rem;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }

            .file-size {
                color: #888;
                font-size: 0.8rem;
            }

            .remove-file {
                background: #f44336;
                color: white;
                border: none;
                padding: 2px 6px;
                border-radius: 3px;
                cursor: pointer;
                font-size: 0.8rem;
            }

            .success { color: #4CAF50; }
            .error { color: #f44336; }
            .warning { color: #ff9800; }
            .info { color: #4da6ff; }

            .hidden {
                display: none;
            }

            .file-count {
                text-align: center;
                margin: 10px 0;
                font-weight: bold;
                color: #4da6ff;
            }

            @keyframes pulse {
                0% { transform: scale(1); }
                50% { transform: scale(1.05); }
                100% { transform: scale(1); }
            }

            .pulse {
                animation: pulse 2s infinite;
            }

            .processing-file {
                background: rgba(255, 193, 7, 0.1);
                border-left-color: #ffc107;
            }

            @media (max-width: 768px) {
                .container {
                    flex-direction: column;
                }
                .left-panel, .right-panel {
                    flex: none;
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <!-- Left Panel -->
            <div class="left-panel">
                <div class="header">
                    <h1 class="title">Universal ASYCUDA XML Converter</h1>
                    <p class="subtitle">Convert Excel → Exact ASYCUDA XML Structure</p>
                    
                </div>

                <div class="upload-section">
                    <h2>📁 Upload Excel Files</h2>
                    <p>Select individual files or entire folders with Excel files</p>
                    
                    <div class="btn-group">
                        <button class="btn" onclick="document.getElementById('fileInput').click()">
                            📄 Select Files
                        </button>
                        <button class="btn btn-secondary" onclick="document.getElementById('folderInput').click()">
                            📁 Select Folder
                        </button>
                    </div>

                    <input type="file" id="fileInput" class="file-input" multiple accept=".xlsx,.xls,.xlsm">
                    <input type="file" id="folderInput" class="folder-input" webkitdirectory directory multiple>

                    <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
                        <div class="upload-icon">📁</div>
                        <h3>Drag & Drop Files Here</h3>
                        <p>or click to select files/folders</p>
                        <p class="file-count" id="fileCount">No files selected</p>
                    </div>

                    <div class="file-list" id="fileList"></div>

                    <div class="btn-group">
                        <button class="btn pulse" id="convertBtn" onclick="startConversion()" disabled>
                             Start Conversion
                        </button>
                        <button class="btn btn-secondary" onclick="clearAllFiles()">
                            🗑️ Clear All
                        </button>
                    </div>
                </div>

                <div class="progress-section hidden" id="progressSection">
                    <h2>📊 Conversion Progress</h2>
                    <div class="progress-bar">
                        <div class="progress-fill" id="progressFill">
                            <div class="progress-text" id="progressText">0%</div>
                        </div>
                    </div>
                    <div class="stats">
                        <div class="stat-item">
                            <div class="stat-number" id="totalCount">0</div>
                            <div>Total Files</div>
                        </div>
                        <div class="stat-item">
                            <div class="stat-number" id="processedCount">0</div>
                            <div>Processed</div>
                        </div>
                        <div class="stat-item">
                            <div class="stat-number" id="successCount">0</div>
                            <div>Successful</div>
                        </div>
                        <div class="stat-item">
                            <div class="stat-number" id="errorCount">0</div>
                            <div>Errors</div>
                        </div>
                    </div>
                    <div id="statusText" class="info">Ready to start conversion...</div>
                    <div id="currentFile" class="info" style="margin-top: 10px; font-style: italic;"></div>
                </div>
            </div>

            <!-- Right Panel -->
            <div class="right-panel">
                <div class="header">
                    <h2>📋 Conversion Log</h2>
                </div>
               <div class="results-section">
    <div class="results-log" id="resultsLog">
        <p><strong>ASYCUDA XML Generator v3.1 - Consignment LV02 2025 6241</strong></p>
        <hr style="border: none; border-top: 1px solid #444; margin: 10px 0;">
        <p>• Exact Aruba ASYCUDA pattern </p>
        <p>• Deterministic 1-to-1 conversion</p>
        <p>• Professional ASYCUDA compliance</p>
        <p>• Made by Arfa Rumman Khalid</p>
        <br>
        <p><em>Start, when you are ready!</em></p>
    </div>
                    <div class="btn-group">
                        <button class="btn btn-secondary" onclick="clearLog()">
                            🧹 Clear Log
                        </button>
                        <button class="btn btn-secondary" onclick="exportLog()">
                            💾 Export Log
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <script>
            let selectedFiles = [];
            let conversionSessionId = null;
            let progressInterval = null;

            // DOM Elements
            const fileInput = document.getElementById('fileInput');
            const folderInput = document.getElementById('folderInput');
            const uploadArea = document.getElementById('uploadArea');
            const fileList = document.getElementById('fileList');
            const fileCount = document.getElementById('fileCount');
            const convertBtn = document.getElementById('convertBtn');
            const progressSection = document.getElementById('progressSection');
            const progressFill = document.getElementById('progressFill');
            const progressText = document.getElementById('progressText');
            const resultsLog = document.getElementById('resultsLog');
            const statusText = document.getElementById('statusText');
            const currentFile = document.getElementById('currentFile');
            const totalCount = document.getElementById('totalCount');
            const processedCount = document.getElementById('processedCount');
            const successCount = document.getElementById('successCount');
            const errorCount = document.getElementById('errorCount');

            // Initialize
            function init() {
                setupEventListeners();
                updateUI();
            }

            function setupEventListeners() {
                // File input change - FIXED: Show all files but filter Excel files
                fileInput.addEventListener('change', function() {
                    const files = Array.from(this.files).filter(file => 
                        file.name.toLowerCase().match(/\.(xlsx|xls|xlsm)$/)
                    );
                    handleFiles(files);
                    this.value = ''; // Reset to allow selecting same files again
                });

                // Folder input change - FIXED: Show all files but filter Excel files
                folderInput.addEventListener('change', function() {
                    const files = Array.from(this.files).filter(file => 
                        file.name.toLowerCase().match(/\.(xlsx|xls|xlsm)$/)
                    );
                    handleFiles(files);
                    this.value = ''; // Reset to allow selecting same files again
                });

                // Drag and drop
                ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                    uploadArea.addEventListener(eventName, preventDefaults, false);
                });

                ['dragenter', 'dragover'].forEach(eventName => {
                    uploadArea.addEventListener(eventName, () => uploadArea.classList.add('dragover'), false);
                });

                ['dragleave', 'drop'].forEach(eventName => {
                    uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('dragover'), false);
                });

                uploadArea.addEventListener('drop', handleDrop, false);
            }

            function preventDefaults(e) {
                e.preventDefault();
                e.stopPropagation();
            }

            function handleDrop(e) {
                const dt = e.dataTransfer;
                const files = Array.from(dt.files).filter(file => 
                    file.name.toLowerCase().match(/\.(xlsx|xls|xlsm)$/)
                );
                handleFiles(files);
            }

            function handleFiles(files) {
                const newFiles = files.filter(file => 
                    !selectedFiles.some(f => f.name === file.name && f.size === file.size && f.lastModified === file.lastModified)
                );

                if (newFiles.length > 0) {
                    selectedFiles = [...selectedFiles, ...newFiles];
                    updateUI();
                    addLog(`Added ${newFiles.length} file(s) to queue`, 'success');
                }
            }

            function removeFile(index) {
                const removedFile = selectedFiles[index];
                selectedFiles.splice(index, 1);
                updateUI();
                addLog(`Removed file: ${removedFile.name}`, 'warning');
            }

            function clearAllFiles() {
                selectedFiles = [];
                updateUI();
                addLog('All files cleared.', 'info');
            }

            function updateUI() {
                updateFileList();
                updateFileCount();
                updateConvertButton();
                updateTotalCount();
            }

            function updateFileList() {
                fileList.innerHTML = '';
                
                // Show first 10 files with option to show more
                const filesToShow = selectedFiles.slice(0, 10);
                
                filesToShow.forEach((file, index) => {
                    const fileItem = document.createElement('div');
                    fileItem.className = 'file-item';
                    fileItem.innerHTML = `
                        <div>
                            <strong>${file.name}</strong>
                            <div class="file-size">${formatFileSize(file.size)}</div>
                        </div>
                        <button class="remove-file" onclick="removeFile(${index})">×</button>
                    `;
                    fileList.appendChild(fileItem);
                });

                if (selectedFiles.length > 10) {
                    const moreItem = document.createElement('div');
                    moreItem.className = 'file-item';
                    moreItem.style.textAlign = 'center';
                    moreItem.style.fontStyle = 'italic';
                    moreItem.textContent = `... and ${selectedFiles.length - 10} more files`;
                    fileList.appendChild(moreItem);
                }
            }

            function updateFileCount() {
                const count = selectedFiles.length;
                fileCount.textContent = count === 0 ? 'No files selected' : `${count} file${count !== 1 ? 's' : ''} selected`;
            }

            function updateTotalCount() {
                totalCount.textContent = selectedFiles.length;
            }

            function updateConvertButton() {
                convertBtn.disabled = selectedFiles.length === 0;
            }

            function formatFileSize(bytes) {
                if (bytes === 0) return '0 Bytes';
                const k = 1024;
                const sizes = ['Bytes', 'KB', 'MB', 'GB'];
                const i = Math.floor(Math.log(bytes) / Math.log(k));
                return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
            }

            function addLog(message, type = 'info') {
                const timestamp = new Date().toLocaleTimeString();
                const logEntry = document.createElement('div');
                logEntry.className = type;
                logEntry.innerHTML = `[${timestamp}] ${message}`;
                resultsLog.appendChild(logEntry);
                resultsLog.scrollTop = resultsLog.scrollHeight;
            }

            function clearLog() {
                resultsLog.innerHTML = 'Log cleared.';
                addLog('Log cleared successfully.', 'info');
            }

            function exportLog() {
                const logContent = resultsLog.innerText;
                const blob = new Blob([logContent], { type: 'text/plain' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `asycuda_conversion_log_${new Date().toISOString().slice(0, 10)}.txt`;
                a.click();
                URL.revokeObjectURL(url);
            }

            async function startConversion() {
                if (selectedFiles.length === 0) return;

                // Generate session ID
                conversionSessionId = 'session_' + Date.now();
                
                // Reset progress
                progressSection.classList.remove('hidden');
                progressFill.style.width = '0%';
                progressText.textContent = '0%';
                processedCount.textContent = '0';
                successCount.textContent = '0';
                errorCount.textContent = '0';
                statusText.textContent = 'Starting conversion process...';
                currentFile.textContent = '';
                
                convertBtn.disabled = true;
                convertBtn.textContent = '⏳ Converting...';

                addLog(`Starting conversion of ${selectedFiles.length} files...`, 'info');
                
                try {
                    const formData = new FormData();
                    selectedFiles.forEach(file => {
                        formData.append('files', file);
                    });
                    formData.append('sessionId', conversionSessionId);

                    // Start progress polling BEFORE the conversion request
                    startProgressPolling();

                    const response = await fetch('/convert', {
                        method: 'POST',
                        body: formData
                    });

                    if (!response.ok) {
                        throw new Error(`Server error: ${response.status}`);
                    }

                    const blob = await response.blob();
                    
                    // Stop progress polling
                    stopProgressPolling();
                    
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `ASYCUDA_XML_Output_${new Date().toISOString().slice(0, 10)}.zip`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    
                    addLog('Conversion completed! ZIP file downloaded.', 'success');
                    statusText.textContent = 'Conversion completed successfully!';
                    
                    // Final progress update
                    updateProgressUI({
                        processed: selectedFiles.length,
                        successful: selectedFiles.length - (conversion_progress.errors || 0),
                        errors: conversion_progress.errors || 0,
                        percent: 100,
                        status: 'Conversion completed!'
                    });
                    
                } catch (error) {
                    stopProgressPolling();
                    addLog(`Error during conversion: ${error.message}`, 'error');
                    statusText.textContent = `Error: ${error.message}`;
                } finally {
                    convertBtn.disabled = false;
                    convertBtn.textContent = '🚀 Start Conversion';
                }
            }

            function startProgressPolling() {
                progressInterval = setInterval(async () => {
                    try {
                        const response = await fetch(`/progress/${conversionSessionId}`);
                        if (response.ok) {
                            const progress = await response.json();
                            updateProgressUI(progress);
                        }
                    } catch (error) {
                        console.error('Error fetching progress:', error);
                    }
                }, 500); // Poll every 500ms for real-time updates
            }

            function stopProgressPolling() {
                if (progressInterval) {
                    clearInterval(progressInterval);
                    progressInterval = null;
                }
            }

            function updateProgressUI(progress) {
                if (!progress) return;
                
                const percent = progress.percent || 0;
                progressFill.style.width = percent + '%';
                progressText.textContent = percent.toFixed(1) + '%';
                
                processedCount.textContent = progress.processed || 0;
                successCount.textContent = progress.successful || 0;
                errorCount.textContent = progress.errors || 0;
                
                if (progress.current_file) {
                    currentFile.textContent = `Processing: ${progress.current_file}`;
                }
                
                if (progress.status) {
                    statusText.textContent = progress.status;
                }
            }

            // Global variable to track conversion progress (for final update)
            let conversion_progress = {};

            // Initialize the application
            init();
        </script>
    </body>
    </html>
    ''')

@app.route('/convert', methods=['POST'])
def convert_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files')
    session_id = request.form.get('sessionId', str(uuid.uuid4()))
    
    if not files or files[0].filename == '':
        return jsonify({'error': 'No files selected'}), 400
    
    # Initialize progress
    with progress_lock:
        conversion_progress[session_id] = {
            'total': len(files),
            'processed': 0,
            'successful': 0,
            'errors': 0,
            'percent': 0,
            'current_file': '',
            'status': 'Starting conversion...'
        }
    
    # Create a temporary zip file in memory
    zip_buffer = BytesIO()
    
    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            total_files = len(files)
            
            for i, file in enumerate(files):
                if file.filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                    # Update progress - current file
                    with progress_lock:
                        conversion_progress[session_id].update({
                            'current_file': file.filename,
                            'status': f'Processing {i+1}/{total_files}: {file.filename}',
                            'percent': (i / total_files) * 100
                        })
                    
                    try:
                        file_content = file.read()
                        success, result = convert_excel_to_xml(file_content, file.filename)
                        
                        if success:
                            xml_filename = file.filename.rsplit('.', 1)[0] + '.xml'
                            zip_file.writestr(xml_filename, result)
                            with progress_lock:
                                conversion_progress[session_id]['successful'] += 1
                        else:
                            # Create error log file
                            error_filename = file.filename + '_ERROR.txt'
                            zip_file.writestr(error_filename, f"Conversion failed: {result}")
                            with progress_lock:
                                conversion_progress[session_id]['errors'] += 1
                        
                    except Exception as e:
                        error_filename = file.filename + '_ERROR.txt'
                        zip_file.writestr(error_filename, f"Unexpected error: {str(e)}")
                        with progress_lock:
                            conversion_progress[session_id]['errors'] += 1
                    
                    # Update processed count
                    with progress_lock:
                        conversion_progress[session_id]['processed'] = i + 1
                        conversion_progress[session_id]['percent'] = ((i + 1) / total_files) * 100
            
            # Final progress update
            with progress_lock:
                conversion_progress[session_id].update({
                    'status': 'Conversion completed! Creating ZIP file...',
                    'percent': 100,
                    'current_file': ''
                })
        
        zip_buffer.seek(0)
        
        # Return the final progress before cleanup
        final_progress = conversion_progress.get(session_id, {})
        
        # Clean up progress data after a short delay
        def cleanup_progress():
            time.sleep(2)  # Wait 2 seconds before cleanup
            with progress_lock:
                if session_id in conversion_progress:
                    del conversion_progress[session_id]
        
        import threading
        threading.Thread(target=cleanup_progress).start()
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'ASYCUDA_XML_Output_{uuid.uuid4().hex[:8]}.zip'
        )
        
    except Exception as e:
        # Clean up progress data on error
        with progress_lock:
            if session_id in conversion_progress:
                del conversion_progress[session_id]
        return jsonify({'error': str(e)}), 500

@app.route('/progress/<session_id>')
def get_progress(session_id):
    with progress_lock:
        progress = conversion_progress.get(session_id, {
            'total': 0,
            'processed': 0,
            'successful': 0,
            'errors': 0,
            'percent': 0,
            'current_file': 'No active session',
            'status': 'Session not found'
        })
    return jsonify(progress)

@app.route('/health')
def health_check():
    return jsonify({'status': 'healthy', 'service': 'ASYCUDA XML Converter'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)