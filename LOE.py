import docx
import pandas as pd

# years
# years_list = ['2017','2018','2019','2020','2021','2022','2023','2024']
years_list = ['2017']

# sheet_names = ['2017-18', '2018-19', '2019-20', '2020-21', '2021-22', '2023-24']
sheet_names = ['2017-18']

owner_name = 'DIPAN MEHTA'
owner_pan = 'FFCPM9766K'

def read_word_file(word_file_path: str, ta_payee_mapping, sheet, payee_address):
    doc = docx.Document(word_file_path)
    # for every payee we have to make these documents
    for index, row in payee_address.iterrows():

        for paragraph in doc.paragraphs:
            
            if "April 01, ______" in paragraph.text:
                paragraph.text = paragraph.text.replace("April 01, ______",f"April 01, {sheet.split('-')[0]}")
            elif "April 01, _____" in paragraph.text:
                paragraph.text = paragraph.text.replace("April 01, _____",f"April 01, {sheet.split('-')[0]}")
            elif "To," in paragraph.text:
                paragraph.text="To,"
                paragraph.text = paragraph.text.replace("To,"
                ,f"""To,\n{row['TRADING ADVISOR - CHANGED']},\n{row['ADDRESS']}\n""")
            # elif "Name:" in paragraph.text:
            #     paragraph.text="Name:"
            #     paragraph.text = paragraph.text.replace("Name:",f"Name:{owner_name}")
            # elif "PAN:" in paragraph.text:
            #     paragraph.text="PAN:"
            #     paragraph.text = paragraph.text.replace("PAN:",f"PAN:{owner_pan}")


        filtered_ta_payee_mapping = ta_payee_mapping[ta_payee_mapping['TRADING ADVISOR'] == row['TRADING ADVISOR - CHANGED']]

        for index,each_row in filtered_ta_payee_mapping.iterrows():
            trading_advisor = each_row['TRADING ADVISOR']
            pan_list = each_row["PAN"]
            payee_list = each_row['PAYEE']
            pan_payee_dict = {}
            for i in range(0,len(pan_list)):
                if payee_list[i] not in pan_payee_dict.keys() and payee_list[i] != trading_advisor:
                    pan_payee_dict[payee_list[i]] = pan_list[i]

            # insert these as table now
            # append the table in the end
            if len(pan_payee_dict.keys()) > 0:
                table = doc.add_table(rows=len(pan_payee_dict.keys())+1,cols=3)
                # table.style = 'Table Grid'

                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'PAYEE'
                hdr_cells[1].text = 'PAN'
                hdr_cells[2].text = 'SIGNATURE'
                for key, value in pan_payee_dict.items():
                    modified_row = table.add_row().cells
                    modified_row[0].text = str(key)
                    modified_row[1].text = str(value)
        
        doc.save(f'sample_output_{trading_advisor}.docx')

# get trading advisor and payee grouped by for last page table
for sheet in sheet_names:
    year_folder = sheet.replace('/', '-')
    # names and address - load from EW details of file
    master_data = pd.read_csv('EW Master 1.csv')
    payee_address = master_data[['TRADING ADVISOR - CHANGED','ADDRESS']]
    payee_address = payee_address.drop_duplicates()
    data = pd.read_excel('EW - Details of Professional Fees Paid for last 7 years 1.xlsx', sheet_name=sheet)
    filtered_data = data[['TRADING ADVISOR','PAYEE','PAN']]
    grouped_data_by_advisor = filtered_data.groupby('TRADING ADVISOR')
    ta_payee_mapping = grouped_data_by_advisor.agg({
        "PAN": list,
        "PAYEE": list
    }).reset_index()

    read_word_file('SAMPLE LOE.docx',ta_payee_mapping,sheet,payee_address)