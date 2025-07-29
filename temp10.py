bidsheet_file = "new/bidsheet_master_consolidate.xlsx"
wapp_file = "wapp2.xlsx"
p21_file = "P21 supplier bid supplier norm 070725v3.xlsx"
supplier_port_file = "Supplier Port per Part table 070925.csv"
frieght_file = "Freight cost mutipliers table 071025v2.csv"
port_country_map = {
    'DALIAN': 'China', 
    'NINGBO': 'China', 
    'QINGDAO': 'China', 
    'SHANGHAI': 'China',
    'SHENZHEN': 'China', 
    'TIANJIN': 'China', 
    'XINGANG': 'China', 
    'XIAMEN': 'China',
    'AHMEDABAD': 'India', 
    'CHENNAI': 'India', 
    'DADRI': 'India', 
    'MUMBAI': 'India',
    'MUNDRA': 'India', 
    'NHAVA SHEVA': 'India',
    'SURABAYA': 'Indonesia',
    'PORT KLANG': 'Malaysia', 
    'PASIR GUDANG': 'Malaysia', 
    'TANJUNG PELAPAS': 'Malaysia',
    'BUSAN': 'South Korea',
    'KAOHSIUNG': 'Taiwan', 
    'KEELUNG': 'Taiwan', 
    'TAICHUNG': 'Taiwan', 
    'TAIPEI': 'Taiwan',
    'BANGKOK': 'Thailand',
    'LAEM CHABANG': 'Thailand',
    'HO CHI MINH CITY': 'Vietnam', 
    'VUNG TAU': 'Vietnam', 
    'HAI PHONG': 'Vietnam'
}
import pandas as pd
supplier_port_df = pd.read_csv(supplier_port_file)

supplier_port_long = supplier_port_df.melt(
    id_vars=['ROW ID #', 'Division', 'Part #'],
    var_name='Supplier',
    value_name='Port'
)

supplier_port_long['Country'] = supplier_port_long['Port'].map(port_country_map)

supplier_port_long.to_csv('supplier_country_mapping.csv', index=False)