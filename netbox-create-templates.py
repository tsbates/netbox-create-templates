import os
from openpyxl import Workbook, workbook
from openpyxl.reader.excel import load_workbook, load_workbook

cwd = os.getcwd()
print('[+] Working in %s' % (cwd))

file = 'IPAM.xlsx'

wb = load_workbook(filename=file)

def delete_sheets_from_workbook(): # Create a fresh start with a single sheet
    allSheets = wb.sheetnames
    print("[-] Removing all existing sheets")
    for sheet in allSheets:
        del wb[sheet]

def create_ipam_sheets(): 
    sheets = ['IP Addresses', 'IP Ranges', 'Prefixes', 'Prefix and VLAN Roles', 'Aggregates', 'RIRs', 'VRFs', 'Route Targets', 'VLANs', 'VLAN Groups', 'Services']
    for sheet in sheets:
        sheet = wb.create_sheet(sheet)

def create_template_ip_addresses():
    sheet = wb['IP Addresses']
    options = ['address', 'vrf', 'tenant', 'status', 'role', 'device', 'virtual_machine', 'interface', 'is_primary','dns_name', 'description']
    for row in range(1):
        sheet.append(options)

def create_template_ip_ranges():
    sheet = wb['IP Ranges']
    options = ['start_address', 'end_address', 'vrf', 'tenant', 'status', 'role', 'description']
    for row in range(1):
        sheet.append(options)

def create_template_prefixes():
    sheet = wb['Prefixes']
    options = ['prefix', 'vrf', 'tenant', 'site', 'vlan_group', 'vlan', 'status', 'role', 'is_pool', 'mark_utilized', 'description']
    for row in range(1):
        sheet.append(options)

def create_template_prefixes_and_vlan_roles():
    sheet = wb['Prefix and VLAN Roles']
    options = ['name', 'slug', 'weight', 'description']
    for row in range(1):
        sheet.append(options)

# TODO
# Complete template creation for aggregation, RIRs, VRFs, Route Targets etc..

def save_workbook():
    wb.save(file)

if __name__ == '__main__':
    delete_sheets_from_workbook()
    create_ipam_sheets()
    create_template_ip_addresses()
    create_template_ip_ranges()
    create_template_prefixes()
    create_template_prefixes_and_vlan_roles()
    save_workbook()
