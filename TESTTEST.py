import xlwings as xw
from pyModbusTCP.client import ModbusClient
import time

# --- CONFIG ---
DEIF_IP = "192.168.0.253"
client = ModbusClient(host=DEIF_IP, port=502, unit_id=1, auto_open=True)

# Connect to the open Excel file
try:
    wb = xw.Book('SolarMonitor.xlsx')
    sheet = wb.sheets[0]
    print("Connected to Excel. Writing data...")
except Exception as e:
    print(f"Error: Make sure 'SolarMonitor.xlsx' is open! {e}")
    exit()

# Setup Excel Header
sheet.range('A1').value = ['Parameter', 'Value', 'Unit']
sheet.range('A2').value = 'Internal Temp'
sheet.range('A3').value = 'Mains P total'
sheet.range('A4').value = 'PV P Reference'

while True:
    # Read the block 590 to 594 (from your image)
    regs = client.read_input_registers(1, 50)
    
    if regs:
        # Update Excel Cells in real-time
        sheet.range('B2').value = regs[0] / 10.0  # Temp
        sheet.range('B3').value = regs[2]         # Mains P
        sheet.range('B4').value = regs[3]         # PV Reference
        
        # Add a "Last Updated" timestamp so you know it's live
        sheet.range('D1').value = f"Last Update: {time.strftime('%H:%M:%S')}"
    else:
        sheet.range('D1').value = "OFFLINE - Check DEIF"
        
    time.sleep(1) # Update every 1 second
