from report_oop import ExcelReportPlugin
import os

# base_path = os.sep.join(os.getcwd().split(os.sep)[:-2])
input_file = r'C:\Users\PutuAndika\OneDrive - Migo\Desktop\Data Engineer Project\Bootcamp Digital Skola\project_1\automate_report\digitalskola\report_putu\input_data\supermarket_sales.xlsx'
output_file = r'C:\Users\PutuAndika\OneDrive - Migo\Desktop\Data Engineer Project\Bootcamp Digital Skola\project_1\automate_report\digitalskola\report_putu\output_data\report_penjualan_2019.xlsx'

automate = ExcelReportPlugin(
    input_file= input_file,
    output_file=output_file
)
automate.main()