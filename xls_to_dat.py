import streamlit as st
import pandas as pd
import numpy as np
import os
import platform
import datetime
import base64
import tempfile



TEMPLATES = {
    "SAWT_1700": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_1701Q": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_1701": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_1702Q": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_1702": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_2550M": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_2550Q": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_2551M": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "SAWT_2553": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, (int,str), (str,float), str, str, str, str, str, (int,float), (int,float), (int,float)],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "QAP_1601EQ Schedule 1": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "QAP_1601EQ Schedule 2": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float)],
        "required": [True, True, False, True, False, False, False, True, True]
    },
    "QAP_1601FQ Schedule 1": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"tax_rate",	"tax_amount"],
        "types": [datetime.date, str, (str, float), str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "QAP_1601FQ Schedule 2": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "fringeBenefit",	"tax_rate",	"grossUpValue", "tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, float, (int,float), float, float],
        "required": [True, True, False, True, False, False, False, True, False, True, True, True]
    },
    "QAP_1601FQ Schedule 3": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "statusCode", "ATC", "income_payment"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, str, (int,float)],
        "required": [True, True, False, True, False, False, False, False, True, True]
    },
    "MAP_1600 VT": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "MAP_1600 PT": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "1604E Schedule 3": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment",	"ewt_rate",	"tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, True, True, True, True]
    },
    "1604E Schedule 4": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "income_payment"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, (int, float)],
        "required": [True, True, False, True, False, False, False, True, True]
    },
    "1604F Schedule 4": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "statusCode", "ATC", "income_payment", "tax_rate", "tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, str, (int, float), (int,float), float],
        "required": [True, True, False, True, False, False, False, False, True, True, True, True]
    },
    "1604F Schedule 5": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "ATC", "fringeBenefit", "tax_rate", "grossUpValue", "tax_amount"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, float, (int,float), float, float],
        "required": [True, True, False, True, False, False, False, True, False, False, True]
    },
    "1604F Schedule 6": {
        "columns": ["Reporting_Month", "Vendor TIN", "branchCode",	"companyName", "surName", "firstName",	"middleName", "statusCode", "ATC", "income_payment"],
        "types": [datetime.date, str, (str,float), str, str, str, str, str, str, (int,float)],
        "required": [True, True, False, True, False, False, False, False, True, True]
    }
} 


MAIN_CATEGORIES = {
    "SAWT": ["SAWT_1700", "SAWT_1701Q", "SAWT_1701", "SAWT_1702Q", "SAWT_1702", "SAWT_2550M", "SAWT_2550Q", "SAWT_2551M", "SAWT_2553"],
    "QAP_1601EQ": ["QAP_1601EQ Schedule 1", "QAP_1601EQ Schedule 2"],
    "QAP_1601FQ": ["QAP_1601FQ Schedule 1", "QAP_1601FQ Schedule 2", "QAP_1601FQ Schedule 3"],
    "MAP_1600": ["MAP_1600 VT", "MAP_1600 PT"],
    "1604E": ["1604E Schedule 3", "1604E Schedule 4"],
    "1604F": ["1604F Schedule 4", "1604F Schedule 5", "1604F Schedule 6"]
}


def detect_template(df):
    for template_name, template_info in TEMPLATES.items():
        if set(df.columns) == set(template_info["columns"]):
            return template_name
    return None

def validate_template(df, template_key):
    errors = []
    
    template_info = TEMPLATES.get(template_key)
    if not template_info:
        errors.append(f"Template {template_key} not found.")
        return errors

    if not set(template_info["columns"]).issubset(df.columns):
        missing_columns = set(template_info["columns"]) - set(df.columns)
        errors.append(f"Missing columns: {', '.join(missing_columns)}.")

    for col, required in zip(template_info["columns"], template_info["required"]):
        if required:
            missing_values = df[col][(df[col].isna()) | (df[col] == "") | (df[col].isnull())]
            if not missing_values.empty:
                errors.append(f"Required column {col} contains missing values.")
                st.write(f"Debug: Missing values in column {col}: {missing_values}")

    st.write(f"Debug: Total Errors = {len(errors)}")  
    return errors


def get_downloads_folder():
    return tempfile.gettempdir()


def get_unique_filename(base_path, tax_id, branch_code, reporting_month, filename_footer, extension):
    counter = 1
    new_path = os.path.join(base_path, f"{tax_id}{branch_code}{reporting_month}{filename_footer}{extension}")
    
    while os.path.exists(new_path):
        new_path = os.path.join(base_path, f"{tax_id}{branch_code}{reporting_month}{filename_footer}({counter}){extension}")
        counter += 1
    
    return new_path


def generate_footer(df, headers, header5, header7, template_name):
    footer = headers_original.copy()


    if "SAWT" in template_name or "MAP_1600" in template_name:
        footer[0] = footer[0].replace("H", "C")
        footer[1] = footer[1].replace("H", "C")
    else:
        footer[0] = df.iloc[0,0].replace("D", "C")
        footer[1] = footer[1].replace("H", "")
        
    if template_name in ["1604E Schedule 3", "1604E Schedule 4", "1604F Schedule 4", "1604F Schedule 5", "1604F Schedule 6"]:
        last_day_of_month = pd.Timestamp(footer[5]).to_period('M').to_timestamp('M')
        footer[5] = last_day_of_month.strftime('%m/%d/%Y')
        
    del footer[6]  # header7
    del footer[4]  # header5
    
    if ("SAWT" in template_name or "MAP" in template_name or "QAP" in template_name or "1604E" in template_name or "1604F Schedule 6" in template_name) and ("income_payment" in df.columns or "grossUpValue" in df.columns):
        if "income_payment" in df.columns:
            income_payment_sum = df["income_payment"].sum()
            footer.append(f"{income_payment_sum:.2f}")
    
        if "grossUpValue" in df.columns:
            gross_up_value_sum = df["grossUpValue"].sum()
            footer.append(f"{gross_up_value_sum:.2f}")


    if "tax_amount" in df.columns:
        tax_amount_sum = df["tax_amount"].sum()
        footer.append(f"{tax_amount_sum:.2f}")

    return footer



def save_as_dat(df, filename, headers):
    downloads_folder = get_downloads_folder()
    
    base_filename = os.path.basename(filename)
    filepath_without_extension = downloads_folder
    unique_filepath = get_unique_filename(filepath_without_extension, tax_id, branch_code, reporting_month, filename_footer, ".dat")
    
    with open(unique_filepath, 'w') as f:
        cleaned_headers = [str(header) if header else "" for header in headers]
        f.write(",".join(cleaned_headers) + "\n")
        
        for _, row in df.iterrows():
            cleaned_row = [
                ('"' + str(item) + '"') if col == "companyName" else str(item) if pd.notna(item) else "" 
                for col, item in zip(df.columns, row)
            ]
            f.write(",".join(cleaned_row) + "\n")
        
        footer = generate_footer(df, headers, header5, header7, template_name)
        f.write(",".join([str(item) for item in footer]))
    

    return unique_filepath


def get_file_download_link(file_path, file_name=None):
    if file_name is None:
        file_name = os.path.basename(file_path)
    
    with open(file_path, 'rb') as f:
        data = f.read()
    
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Download {file_name}</a>'




MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def match_columns_to_template(data, template_name):
    template_cols = TEMPLATES[template_name]['columns']
    uploaded_cols = list(data.columns)
    
    st.write("Please match the uploaded columns to the required template columns:")
    
    column_mapping = {}
    already_selected = []

    exclude_columns = ["Fixed_Column_1", "Fixed_Column_2", "Reporting_Month"]

    if "Vendor TIN" in template_cols and "Vendor TIN" not in exclude_columns:
        options = ["--Not Present--"] + [col for col in uploaded_cols if col not in already_selected and col not in exclude_columns]
        default_index = options.index("Vendor TIN") if "Vendor TIN" in uploaded_cols else 0
        tin_selected = st.selectbox(f"Match column for Vendor TIN", options, index=default_index)
        column_mapping["Vendor TIN"] = tin_selected
        if tin_selected != "--Not Present--":
            already_selected.append(tin_selected)
    
    if "companyName" in template_cols and "companyName" not in exclude_columns:
        options = ["--Not Present--"] + [col for col in uploaded_cols if col not in already_selected and col not in exclude_columns]
        default_index = options.index("companyName") if "companyName" in uploaded_cols else 0
        company_selected = st.selectbox(f"Match column for companyName", options, index=default_index)
        column_mapping["companyName"] = company_selected
        if company_selected != "--Not Present--":
            already_selected.append(company_selected)
    
    for template_col in template_cols:
        if template_col in ["Reporting_Month", "Vendor TIN", "companyName", "Fixed_Column_1", "Fixed_Column_2"]:  
            continue
        
       
        if company_selected != "--Not Present--" and template_col in ["surName", "firstName", "middleName"]:
            st.write(f"The column {template_col} is disabled because companyName has been selected.")
            column_mapping[template_col] = "--Not Present--"
            continue
        
        options = ["--Not Present--"] + [col for col in uploaded_cols if col not in already_selected and col not in exclude_columns]
        default_index = options.index(template_col) if template_col in uploaded_cols else 0
        selected_col = st.selectbox(f"Match column for {template_col}", options, index=default_index)
        column_mapping[template_col] = selected_col
        if selected_col != "--Not Present--":
            already_selected.append(selected_col)

    if column_mapping.get("branchCode", "") == "--Not Present--":
        data["branchCode"] = "0000"

    return column_mapping



def add_missing_columns(df, template_name):
    template_info = TEMPLATES[template_name]
    for col in template_info["columns"]:
        if col not in df.columns:
            df[col] = np.nan
    return df


st.title("XLS to DAT Converter")
st.write("""
Follow the instructions below to convert your file.
""")


st.subheader("1. Upload Your File")
uploaded_file = st.file_uploader("Choose an xls or xlsx file", type=["xls", "xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    
    sheet_name = st.selectbox("Select the sheet you want to work with:", sheet_names)
    
    header_row = st.number_input("Select the row number where headers start", min_value=0, value=0, step=1)
    start_data_row = st.number_input("Select the row number to start the data from", min_value=header_row, value=header_row, step=1)
    
    headers = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, nrows=0)
    tin_columns = [col for col in headers.columns if "TIN" in col]
    
    dtype_dict = {col: str for col in tin_columns}
    data = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, dtype=dtype_dict).iloc[start_data_row - header_row:] 


    st.write(data)

    
    st.subheader("2. Select Template")
    st.write("Select the main category and specific template that matches your file's format.")

    main_category = st.selectbox("Select the main category:", list(MAIN_CATEGORIES.keys()))
    template_name = st.selectbox("Select the template:", MAIN_CATEGORIES[main_category])
    
    if "Reporting_Month" in TEMPLATES[template_name]["columns"]:
        selected_month = st.selectbox("Select Month:", MONTHS)
        month_idx = MONTHS.index(selected_month) + 1
        data["Reporting_Month"] = pd.Timestamp(f"2023-{month_idx}-01").date()
        
    header6 = pd.Timestamp(f"2023-{month_idx}-01").strftime('%m/%Y') if month_idx else None


    col_mapping = match_columns_to_template(data, template_name)
    for template_col, matched_col in col_mapping.items():
        data = data.rename(columns={matched_col: template_col})

    data = add_missing_columns(data, template_name)  
    
    columns_order = TEMPLATES[template_name]["columns"]
    data = data[columns_order]
    required_cols = [col for col, req in zip(TEMPLATES[template_name]["columns"], TEMPLATES[template_name]["required"]) if req]
    data = data.dropna(subset=required_cols, how='any')
    
    if 'ewt_rate' in data.columns:
        data['ewt_rate'] = data['ewt_rate'] * 100

    if 'tax_rate' in data.columns:
        data['tax_rate'] = data['tax_rate'] * 100
            
    
    st.subheader("3. Data Preview and Validation")
    st.write("Review your data below and check if there are any validation errors.")
 
 
    st.write(data)
    
        
    st.subheader("4. Provide Additional Information")
    st.write("Provide additional required details.")
        
    tax_id = st.text_input("Tax Identification Number: ", key="tax_id_input")    
    branch_code = st.text_input("Branch Code: ", key="branch_code_input")
    rdo_code = st.text_input("RDO Code: ", key="rdo_code_input")
        
    data_with_fixed_columns = data.copy()
        
    if main_category == "SAWT" or main_category == "MAP_1600":
        if main_category == "SAWT":
            data_with_fixed_columns["Fixed_Column_1"] = "D" + "SAWT"
        else: 
            data_with_fixed_columns["Fixed_Column_1"] = "D" + "MAP"
        data_with_fixed_columns["Fixed_Column_2"] = "D" + template_name.split("_")[1]
    else:
        if "Schedule" in template_name:
            schedule_number = template_name.split(" ")[-1]
            data_with_fixed_columns["Fixed_Column_1"] = "D" + schedule_number
            base_template_name_parts = template_name.rsplit(" ", 2)[0].split("_")  
            if len(base_template_name_parts) > 1:
                data_with_fixed_columns["Fixed_Column_2"] = "D" + base_template_name_parts[1].replace("QAP_", "")
            else:
                data_with_fixed_columns["Fixed_Column_2"] = "D" + base_template_name_parts[0]
        else:
            data_with_fixed_columns["Fixed_Column_1"] = "D" + template_name
            data_with_fixed_columns["Fixed_Column_2"] = "D" + template_name
                
        
    data_with_fixed_columns['Fixed_Column_3'] = range(1, len(data) + 1)
    
    fixed_columns = ["Fixed_Column_1", "Fixed_Column_2", "Fixed_Column_3"]
    other_columns = [col for col in data.columns if col not in fixed_columns]
    data_with_fixed_columns = data_with_fixed_columns[fixed_columns + other_columns]
    
    cols = data_with_fixed_columns.columns.tolist()
    
    if "Reporting_Month" in data_with_fixed_columns.columns:
        data_with_fixed_columns["Reporting_Month"] = pd.to_datetime(data_with_fixed_columns["Reporting_Month"])
        data_with_fixed_columns["Reporting_Month"] = data_with_fixed_columns["Reporting_Month"].dt.strftime('%m/%Y')
    
    # For SAWT, MAP, or specific QAP templates
    if main_category == "SAWT" or main_category == "MAP_1600" or template_name in ["QAP_1601EQ Schedule 1", "QAP_1601EQ Schedule 2"]:
        if "Reporting_Month" in cols and "ATC" in cols:
            cols.insert(cols.index("ATC") - 1, cols.pop(cols.index("Reporting_Month")))
        # Ensure Fixed_Column_3 is the third column, if it exists
        if "Fixed_Column_3" in cols:
            cols.insert(2, cols.pop(cols.index("Fixed_Column_3")))
    
        
    # For 1601FQ
    elif template_name in ["QAP_1601FQ Schedule 1", "QAP_1601FQ Schedule 2"]:
        if "Fixed_Column_3" in data_with_fixed_columns and "ATC" in data_with_fixed_columns:
            cols.insert(cols.index("ATC") -1, cols.pop(cols.index("Fixed_Column_3")))
        if "middleName" in data_with_fixed_columns and "Reporting_Month" in data_with_fixed_columns:
            cols.insert(cols.index("middleName")+1, cols.pop(cols.index("Reporting_Month")))
        
    # For 1604E and 1604F
    if template_name in ["1604E Schedule 3", "1604E Schedule 4", "1604F Schedule 4", "1604F Schedule 5", "1604F Schedule 6"]:
        data_with_fixed_columns["Tax_ID_Column"] = tax_id
        data_with_fixed_columns["Branch_Code_Column"] = branch_code
    
        desired_order = ["Tax_ID_Column", "Branch_Code_Column", "Reporting_Month", "Fixed_Column_3", "Vendor TIN", "branchCode"]
    
        for column_name in desired_order:
            if column_name in cols:
                cols.remove(column_name)
        for idx, column_name in enumerate(desired_order, 2):
            cols.insert(idx, column_name)
    
    data_with_fixed_columns = data_with_fixed_columns[cols]
    

    if 'companyName' in data.columns:
        company_name = st.text_input("Company Name: ")
        header5 = '"' + company_name + '"'
    else:
        first_name = st.text_input("Firstname: ")
        middle_name = st.text_input("Middlename: ")
        last_name = st.text_input("Lastname: ")
        header5 = f"{first_name} {middle_name} {last_name}"
            
    if main_category == "SAWT" or main_category == "MAP_1600":
        if main_category == "SAWT":
            header1 = "HSAWT"
            header2 = "H" + template_name.split("_")[1] 
        else:
            header1 = "HMAP"
            header2 = "H" + template_name.split("_")[1].split(" ")[0] + template_name.split(" ")[1]
        
    elif main_category in ["QAP_1601EQ", "QAP_1601FQ"]:
        header1 = "HQAP"
        header2 = "H" + template_name.split("_")[1].split("Schedule")[0]  
    else:
        base_template_name = template_name.split(" ")[0]
        header1_value = data_with_fixed_columns["Fixed_Column_1"].unique()[0]
        header1 = "H" + str(header1_value)
        header2 = "H" + base_template_name
        
    header3 = tax_id
    header4 = branch_code
    header7 = rdo_code
    reporting_month = pd.Timestamp(f"2023-{month_idx}-01").to_period('M').to_timestamp('M').strftime('%m%d%Y') if month_idx else None
        
    header6 = pd.Timestamp(f"2023-{month_idx}-01").strftime('%m/%Y') if month_idx else None
    headers_original = [header1, header2, header3, header4, header5, header6, header7]
    
    footer = generate_footer(data_with_fixed_columns, headers_original, header5, header7, template_name)
    footer_copy = headers_original.copy()
    filename_footer = footer_copy[1].replace("H", "")

    if template_name in ["1604E Schedule 3", "1604E Schedule 4", "1604F Schedule 4", "1604F Schedule 5", "1604F Schedule 6"]:
        last_day_of_month = pd.Timestamp(f"2023-{month_idx}-01").to_period('M').to_timestamp('M')
        header6 = last_day_of_month.strftime('%m/%d/%Y') if month_idx else None
        headers = [header2, header3, header4, header6, "N", "0", header7]
    else:
        headers = headers_original
    

    st.subheader("5. Convert to DAT")
    st.write("After confirming all details, click the button below to convert your file to DAT format.")

    if st.button('Convert to DAT'):
        all_no_data_or_nan_cols = (data_with_fixed_columns.replace("No data", np.nan).isna().all(axis=0))
    
        data_with_fixed_columns.loc[:, all_no_data_or_nan_cols] = ""
    
        dat_file = os.path.splitext(os.path.basename(uploaded_file.name))[0] + ".dat"
        saved_path = save_as_dat(data_with_fixed_columns, dat_file, headers)
        st.success(f"File converted successfully!")
        
        download_link = get_file_download_link(saved_path)
        st.markdown(download_link, unsafe_allow_html=True)

    

