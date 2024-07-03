import pandas as pd
import numpy as np
import openpyxl

# Load the Excel workbook
workbook_path = 'BLS - Industry By County Data.xlsx'
mapping_csv_file_path = 'mapping market to county.csv'

# Read all sheets at once
with pd.ExcelFile(workbook_path) as excel_file:
    sheet_names = [sheet for sheet in excel_file.sheet_names if sheet.lower() != "are codes"]
    dfs = {sheet: excel_file.parse(sheet) for sheet in sheet_names}

# Initialize dictionaries to store results and total employment
results = {}
total_employment = {}

# Process each sheet
for sheet, df in dfs.items():
    total_emp = df[(df['agglvl_code'] == 70) | (df['agglvl_code'] == 40)]['annual_avg_emplvl'].sum()
    total_employment[sheet] = total_emp
    df_filtered = df[(df['agglvl_code'] == 75) | (df['agglvl_code'] == 45)]
    industry_percentages = df_filtered.groupby('industry_code')['annual_avg_emplvl'].sum() / total_emp * 100
    results[sheet] = industry_percentages.to_dict()

output_df = pd.DataFrame(results)

# Create mapping data and DataFrames
mapping_data = {
    'NAICS': [111, 112, 113, 114, 115, 211, 212, 213, 221, 236, 237, 238, 311, 312, 313, 314, 315, 316, 321, 322, 323, 324, 325, 326, 327, 331, 332, 333, 334, 335, 336, 337, 339, 423, 424, 425, 441, 444, 445, 449, 455, 456, 457, 458, 459, 481, 482, 483, 484, 485, 486, 487, 488, 491, 492, 493, 512, 513, 516, 517, 518, 519, 521, 522, 523, 524, 525, 531, 532, 533, 541, 551, 561, 562, 611, 621, 622, 623, 624, 711, 712, 713, 721, 722, 811, 812, 813, 814, 921, 922, 923, 924, 925, 926, 928, 999, 927],
    'Job Description': ['Crop Production Worker', 'Animal Production and Aquaculture Worker', 'Forestry and Logging Worker', 'Fishing and Hunting Worker', 'Agriculture and Forestry Support Worker', 'Oil and Gas Extraction Worker', 'Mining Worker (except Oil and Gas)', 'Mining Support Activities Worker', 'Utilities Worker', 'Construction of Buildings Worker', 'Heavy and Civil Engineering Construction Worker', 'Specialty Trade Contractor', 'Food Manufacturing Worker', 'Beverage and Tobacco Product Manufacturing Worker', 'Textile Mills Worker', 'Textile Product Mills Worker', 'Apparel Manufacturing Worker', 'Leather and Allied Product Manufacturing Worker', 'Wood Product Manufacturing Worker', 'Paper Manufacturing Worker', 'Printing and Related Support Activities Worker', 'Petroleum and Coal Products Manufacturing Worker', 'Chemical Manufacturing Worker', 'Plastics and Rubber Products Manufacturing Worker', 'Nonmetallic Mineral Product Manufacturing Worker', 'Primary Metal Manufacturing Worker', 'Fabricated Metal Product Manufacturing Worker', 'Machinery Manufacturing Worker', 'Computer and Electronic Product Manufacturing Worker', 'Electrical Equipment and Appliance Manufacturing Worker', 'Transportation Equipment Manufacturing Worker', 'Furniture and Related Product Manufacturing Worker', 'Miscellaneous Manufacturing Worker', 'Merchant Wholesaler (Durable Goods)', 'Merchant Wholesaler (Nondurable Goods)', 'Wholesale Electronic Markets and Agents and Brokers', 'Motor Vehicle and Parts Dealer', 'Building Material and Garden Equipment Supplier', 'Food and Beverage Store Worker', 'Furniture and Home Furnishings Store Worker', 'General Merchandise Store Worker', 'Health and Personal Care Store Worker', 'Gasoline Station Worker', 'Clothing and Clothing Accessories Store Worker', 'Sporting Goods, Hobby, Book, and Music Store Worker', 'Air Transportation Worker', 'Rail Transportation Worker', 'Water Transportation Worker', 'Truck Transportation Worker', 'Transit and Ground Passenger Transportation Worker', 'Pipeline Transportation Worker', 'Scenic and Sightseeing Transportation Worker', 'Transportation Support Activities Worker', 'Postal Service Worker', 'Courier and Messenger', 'Warehousing and Storage Worker', 'Motion Picture and Sound Recording Industries Worker', 'Broadcasting Worker (except Internet)', 'Internet Publishing and Broadcasting Worker', 'Telecommunications Worker', 'Data Processing, Hosting, and Related Services Worker', 'Other Information Services Worker', 'Monetary Authorities-Central Bank Worker', 'Credit Intermediation and Related Activities Worker', 'Securities, Commodity Contracts, and Other Financial Investments Worker', 'Insurance Carriers and Related Activities Worker', 'Funds, Trusts, and Other Financial Vehicles Worker', 'Real Estate Worker', 'Rental and Leasing Services Worker', 'Lessors of Nonfinancial Intangible Assets Worker', 'Professional, Scientific, and Technical Services Worker', 'Management of Companies and Enterprises Worker', 'Administrative and Support Services Worker', 'Waste Management and Remediation Services Worker', 'Educational Services Worker', 'Ambulatory Health Care Services Worker', 'Hospital Worker', 'Nursing and Residential Care Facilities Worker', 'Social Assistance Worker', 'Performing Arts, Spectator Sports, and Related Industries Worker', 'Museums, Historical Sites, and Similar Institutions Worker', 'Amusement, Gambling, and Recreation Industries Worker', 'Accommodation Worker', 'Food Services and Drinking Places Worker', 'Repair and Maintenance Worker', 'Personal and Laundry Services Worker', 'Religious, Grantmaking, Civic, Professional Organizations Worker', 'Private Households Worker', 'Executive, Legislative, and General Government Worker', 'Justice, Public Order, and Safety Activities Worker', 'Administration of Human Resource Programs Worker', 'Administration of Environmental Quality Programs Worker', 'Administration of Housing Programs, Urban Planning Worker', 'Administration of Economic Programs Worker', 'National Security and International Affairs Worker', 'Unclassified Establishments Worker', 'Space Research and Technology Worker']
}

mapping_df = pd.DataFrame(mapping_data)

naics_to_ifr = {
    111: 'Agriculture', 112: 'Agriculture', 113: 'Agriculture', 114: 'Agriculture', 115: 'Agriculture',
    211: 'Mining', 212: 'Mining', 213: 'Mining',
    221: 'Utilities',
    236: 'Construction', 237: 'Construction', 238: 'Construction',
    311: 'Food and Beverages', 312: 'Food and Beverages',
    313: 'Textiles', 314: 'Textiles', 315: 'Textiles', 316: 'Textiles',
    321: 'Wood and Furniture', 322: 'Paper', 323: 'Paper',
    324: 'Plastic and Chemicals', 325: 'Plastic and Chemicals', 326: 'Plastic and Chemicals',
    327: 'Glass and Ceramics',
    331: 'Basic Metals', 332: 'Metal Products', 333: 'Metal Machinery',
    334: 'Electronics', 335: 'Electronics',
    336: 'Automotive',
    337: 'Wood and Furniture',
    339: 'Other Manufacturing',
    423: 'Other Non-Manufacturing', 424: 'Other Non-Manufacturing', 425: 'Other Non-Manufacturing',
    441: 'Other Non-Manufacturing', 444: 'Other Non-Manufacturing', 445: 'Other Non-Manufacturing',
    449: 'Other Non-Manufacturing', 455: 'Other Non-Manufacturing', 456: 'Other Non-Manufacturing',
    457: 'Other Non-Manufacturing', 458: 'Other Non-Manufacturing', 459: 'Other Non-Manufacturing',
    481: 'Other Non-Manufacturing', 482: 'Other Non-Manufacturing', 483: 'Other Non-Manufacturing',
    484: 'Other Non-Manufacturing', 485: 'Other Non-Manufacturing', 486: 'Other Non-Manufacturing',
    487: 'Other Non-Manufacturing', 488: 'Other Non-Manufacturing', 491: 'Other Non-Manufacturing',
    492: 'Other Non-Manufacturing', 493: 'Other Non-Manufacturing', 512: 'Other Non-Manufacturing',
    513: 'Other Non-Manufacturing', 516: 'Other Non-Manufacturing', 517: 'Other Non-Manufacturing',
    518: 'Other Non-Manufacturing', 519: 'Other Non-Manufacturing', 521: 'Other Non-Manufacturing',
    522: 'Other Non-Manufacturing', 523: 'Other Non-Manufacturing', 524: 'Other Non-Manufacturing',
    525: 'Other Non-Manufacturing', 531: 'Other Non-Manufacturing', 532: 'Other Non-Manufacturing',
    533: 'Other Non-Manufacturing', 541: 'Education, Research and Development',
    551: 'Other Non-Manufacturing', 561: 'Other Non-Manufacturing', 562: 'Other Non-Manufacturing',
    611: 'Education, Research and Development',
    621: 'Other Non-Manufacturing', 622: 'Other Non-Manufacturing', 623: 'Other Non-Manufacturing',
    624: 'Other Non-Manufacturing', 711: 'Other Non-Manufacturing', 712: 'Other Non-Manufacturing',
    713: 'Other Non-Manufacturing', 721: 'Other Non-Manufacturing', 722: 'Other Non-Manufacturing',
    811: 'Other Non-Manufacturing', 812: 'Other Non-Manufacturing', 813: 'Other Non-Manufacturing',
    814: 'Other Non-Manufacturing', 921: 'Other Non-Manufacturing', 922: 'Other Non-Manufacturing',
    923: 'Other Non-Manufacturing', 924: 'Other Non-Manufacturing', 925: 'Other Non-Manufacturing',
    926: 'Other Non-Manufacturing', 927: 'Other Non-Manufacturing', 928: 'Other Non-Manufacturing',
    999: 'Other Non-Manufacturing'
}

mapping_df['IFR Industry'] = mapping_df['NAICS'].map(naics_to_ifr)

ai_robotics_adoption = {
    111: 0.029, 112: 0.029, 113: 0.029, 114: 0.029, 115: 0.029,
    211: 0.028, 212: 0.028, 213: 0.028,
    221: 0.085,
    236: 0.053, 237: 0.053, 238: 0.053,
    311: 6.776, 312: 6.776,
    313: 0.154, 314: 0.154, 315: 0.154, 316: 0.154,
    321: 2.155, 337: 2.155,
    322: 0.273, 323: 0.273,
    324: 13.497, 325: 13.497, 326: 13.497,
    327: 1.409,
    331: 4.406,
    332: 10.599,
    333: 3.994,
    334: 2.701, 335: 2.701,
    336: 47.101,
    339: 1.703,
    611: 0.214,
}

mapping_df['AI/Robotics Adoption (units per thousand workers)'] = mapping_df['NAICS'].map(ai_robotics_adoption).fillna(0.001)

# Filter out "Other Non-Manufacturing"
significant_adoption = mapping_df[mapping_df['IFR Industry'] != 'Other Non-Manufacturing']

def get_qualitative_assessment(value):
    if value <= 1:
        return 'Low'
    elif value <= 10:
        return 'Medium'
    elif value <= 100:
        return 'High'
    else:
        return 'Very High'

# Add Qualitative Assessment to mapping_df
mapping_df['Qualitative Assessment'] = mapping_df.apply(
    lambda row: get_qualitative_assessment(row['AI/Robotics Adoption (units per thousand workers)'])
    if row['IFR Industry'] != 'Other Non-Manufacturing'
    else 'Not Applicable', 
    axis=1
)

# Calculate statistics for each bucket (excluding 'Not Applicable')
stats = significant_adoption.groupby(
    significant_adoption['AI/Robotics Adoption (units per thousand workers)'].apply(get_qualitative_assessment)
)['AI/Robotics Adoption (units per thousand workers)'].agg(['count', 'min', 'max', 'mean', 'median', 'std'])
stats.columns = ['Count', 'Min', 'Max', 'Average', 'Median', 'StdDev']

# Calculate overall statistics (excluding "Other Non-Manufacturing")
overall_stats = significant_adoption['AI/Robotics Adoption (units per thousand workers)'].agg(['std', 'max', 'min', 'mean', 'median'])
overall_stats = pd.DataFrame({
    'Statistic': ['StdDev', 'Max', 'Min', 'Average', 'Median'],
    'Value': overall_stats
})

# Calculate median for each bucket
bucket_medians = significant_adoption.groupby(
    significant_adoption['AI/Robotics Adoption (units per thousand workers)'].apply(get_qualitative_assessment)
)['AI/Robotics Adoption (units per thousand workers)'].median()

# Dictionary of NAICS codes and their qualitative assessments for Other Non-Manufacturing
other_non_manufacturing_assessments = {
    '423': 'Low', '424': 'Low', '425': 'Low', '441': 'Low', '444': 'Low', '445': 'Medium',
    '449': 'Medium', '455': 'Medium', '456': 'Low', '457': 'Medium', '458': 'Medium',
    '459': 'Medium', '481': 'High', '482': 'High', '483': 'High', '484': 'High',
    '485': 'High', '486': 'Low', '487': 'Low', '488': 'Medium', '491': 'Medium',
    '492': 'High', '493': 'High', '512': 'Medium', '513': 'Low', '516': 'Medium',
    '517': 'High', '518': 'Low', '519': 'Low', '521': 'Low', '522': 'Medium',
    '523': 'Medium', '524': 'Medium', '525': 'Medium', '531': 'Medium', '532': 'Low',
    '533': 'Low', '541': 'Medium', '551': 'Low', '561': 'High', '562': 'High',
    '611': 'Low', '621': 'Low', '622': 'Low', '623': 'Low', '624': 'Low', '711': 'Low',
    '712': 'Low', '713': 'Medium', '721': 'High', '722': 'Medium', '811': 'Low',
    '812': 'Low', '813': 'Low', '814': 'Low', '921': 'Medium', '922': 'Low',
    '923': 'Low', '924': 'Low', '925': 'Low', '926': 'Low', '927': 'Low', '928': 'Low',
    '999': 'Low'
}

# Function to assign median values to Other Non-Manufacturing industries
def assign_median_to_other(row):
    if row['IFR Industry'] == 'Other Non-Manufacturing':
        naics = row['NAICS']
        qual_assessment = other_non_manufacturing_assessments.get(str(naics), 'Low')
        return bucket_medians[qual_assessment]
    else:
        return row['AI/Robotics Adoption (units per thousand workers)']

# Assign median values to Other Non-Manufacturing industries
mapping_df['AI/Robotics Adoption (units per thousand workers)'] = mapping_df.apply(assign_median_to_other, axis=1)

# Update Qualitative Assessment for Other Non-Manufacturing industries
mapping_df['Qualitative Assessment'] = mapping_df.apply(
    lambda row: other_non_manufacturing_assessments.get(str(row['NAICS']), 'Low')
    if row['IFR Industry'] == 'Other Non-Manufacturing'
    else row['Qualitative Assessment'],
    axis=1
)

# Update ai_robotics_adoption dictionary
for naics, assessment in other_non_manufacturing_assessments.items():
    ai_robotics_adoption[int(naics)] = bucket_medians[assessment]

# New calculate_county_units function
def calculate_county_units(county_data, mapping_df):
    total_units = 0
    total_emp = 0
    for naics, percentage in county_data.items():
        naics_3digit = int(str(naics)[:3])
        industry_row = mapping_df[mapping_df['NAICS'] == naics_3digit]
        if not industry_row.empty:
            adoption_units = industry_row['AI/Robotics Adoption (units per thousand workers)'].values[0]
            proportion = percentage / 100
            emp = proportion * total_employment[county]
            total_units += adoption_units * emp / 1000
            total_emp += emp
    return total_units, total_emp

# Recalculate county units
county_units = {}
county_employment = {}
for county, data in results.items():
    units, emp = calculate_county_units(data, mapping_df)
    county_units[county] = units
    county_employment[county] = emp

# Update county_units_df
county_units_df = pd.DataFrame.from_dict(county_units, orient='index', columns=['Total AI/Robotics Units'])
county_units_df['Total Employment'] = pd.Series(county_employment)
county_units_df.index.name = 'County'
county_units_df['Total AI/Robotics Units per Thousand Workers'] = county_units_df['Total AI/Robotics Units'] / county_units_df['Total Employment'] * 1000

# Read the county to market mapping CSV
county_market_mapping = pd.read_csv(mapping_csv_file_path)

# Create a dictionary to map counties to markets
county_to_market = {}
for _, row in county_market_mapping.iterrows():
    market = row['market_publish']
    for i in range(1, 4):
        county = row[f'County {i}']
        if pd.notna(county):
            county_to_market[county] = market

# Add market information to county_units_df
county_units_df['Market'] = county_units_df.index.map(lambda x: county_to_market.get(x, 'Other Market'))

# Calculate market-level data
market_units_df = county_units_df.groupby('Market').agg({
    'Total AI/Robotics Units': 'sum',
    'Total Employment': 'sum'
})
market_units_df['Total AI/Robotics Units per Thousand Workers'] = market_units_df['Total AI/Robotics Units'] / market_units_df['Total Employment'] * 1000

# Calculate Z-scores for Market AI Units
market_units_df['Z-score'] = (market_units_df['Total AI/Robotics Units'] - market_units_df['Total AI/Robotics Units'].mean()) / market_units_df['Total AI/Robotics Units'].std()

# Calculate statistics for each bucket (including updated Other Non-Manufacturing)
stats = mapping_df.groupby('Qualitative Assessment')['AI/Robotics Adoption (units per thousand workers)'].agg(['count', 'min', 'max', 'mean', 'median', 'std'])
stats.columns = ['Count', 'Min', 'Max', 'Average', 'Median', 'StdDev']

# Calculate overall statistics
overall_stats = mapping_df['AI/Robotics Adoption (units per thousand workers)'].agg(['std', 'max', 'min', 'mean', 'median'])
overall_stats = pd.DataFrame({
    'Statistic': ['StdDev', 'Max', 'Min', 'Average', 'Median'],
    'Value': overall_stats
})

# Save results to a new Excel file
with pd.ExcelWriter('industry_analysis_results.xlsx') as writer:
    output_df.to_excel(writer, sheet_name='Industry Analysis', index=True)
    mapping_df.to_excel(writer, sheet_name='NAICS Code Mapping', index=False)
    pd.DataFrame(naics_to_ifr.items(), columns=['NAICS', 'IFR Industry']).to_excel(writer, sheet_name='NAICS to IFR Mapping', index=False)
    county_units_df.to_excel(writer, sheet_name='County AI-Robotics Units')
    county_market_mapping.to_excel(writer, sheet_name='County to Market Lookup', index=False)
    market_units_df.to_excel(writer, sheet_name='Market AI Robotics Units')
    stats.to_excel(writer, sheet_name='Qualitative Assessment Stats')
    overall_stats.to_excel(writer, sheet_name='Overall Adoption Stats', index=False)

print("Results saved to 'industry_analysis_results.xlsx' with eight sheets: 'Industry Analysis', 'NAICS Code Mapping', 'NAICS to IFR Mapping', 'County AI-Robotics Units', 'County to Market Lookup', 'Market AI Robotics Units', 'Qualitative Assessment Stats', and 'Overall Adoption Stats'")