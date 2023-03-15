import datetime

input_file = r'C:\Users\James\Desktop\input.xlsx'
current_date = datetime.datetime.now().strftime('%m.%d.%Y')
output_file = fr'C:\Users\James\Desktop\OptumRx Accumulations {current_date}.xlsx'
image_file = fr'C:\Users\James\Desktop\Vendor_Breakdown_No_Optum_{current_date}.png'
