import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from datetime import datetime
from pathlib import Path

# Set the page title and other configurations
st.set_page_config(
    page_title="Lọc trùng K12",  # Page title
    page_icon="https://drive.google.com/uc?export=view&id=1WDsZ7FvDubT9dyXNdcyfGSMGW3iduDtw"
)

# Define filtered sheet names and new column names
filtered_sheet_names = ['DATA 2',
 'Data_SG_TSD1',
 'Lop4_DT',
 'Data gửi TK',
 'DATA 8',
 'DATA 3',
 'Lop7_Kho1',
 'Data_HN_TSD3',
 'Data 4',
 'Data_ThanhTri_Tsd2',
 'Lop_4',
 'Data_SG',
 'DATA 6',
 'DATA 4',
 'DATA 10',
 'Data_Kho 1_TSD3',
 'Data_Tonghop',
 'DATA 9',
 'Lop5_ThanhTri',
 'DATA luyện gọi ',
 'Lớp 4',
 'Data_Kho 3',
 'Lop6_Kho1',
 'Lop6_HN',
 'Data 10',
 'Data_Thanhtri_TSD1',
 'Data 9',
 'Quảng Bình',
 'Lop5_HN',
 'Data 1',
 'Các lớp khác']
new_column_names = ['Mã Kho', 
    'Tên trường', 
    'Tên cha/mẹ 1', 
    'Tên cha/mẹ 2', 
    'Họ tên con', 
    'SĐT', 
    'Năm sinh', 
    'Lớp của con', 
    'Địa chỉ', 
    'STT', 
    'CTV', 
    'Ngày gọi', 
    'Ca', 
    'Level', 
    'Trạng thái cuộc gọi', 
    'Lý do KH từ chối CC2.3', 
    '"Kết quả ngày 1\n(Ngày - giờ gọi - note chi tiết)"', 
    '"Kết quả ngày 2\n(Ngày - giờ gọi - note chi tiết)"']

def process_excel_files(uploaded_files):
    """
    Process uploaded Excel files, filter relevant sheets, 
    and combine the data into a single DataFrame.
    """
    data_frames = []
    total_files = len(uploaded_files)

    # Placeholder for progress message
    progress_message = st.empty()
    
    # Initialize progress message to show 0 files processed initially
    progress_message.info(f"Đang tổng hợp 0 trên tổng số {total_files} file...")
    
    for i, file in enumerate(uploaded_files):
        try:
            # Load the Excel file
            excel_file = pd.ExcelFile(file)
            
            # Loop through each filtered sheet name
            for sheet in filtered_sheet_names:
                if sheet in excel_file.sheet_names:
                    # Read the sheet into a DataFrame, skip the first row
                    df = pd.read_excel(file, sheet_name=sheet, header=1)
                    
                    # Keep only the first 18 columns
                    df = df.iloc[:, :18]
                    
                    # Rename the columns according to the provided mapping
                    df.columns = new_column_names
                    
                    # Optionally add a column to keep track of the source file/sheet
                    df['Source File'] = file.name  # .name attribute to get the file name
                    df['Source Sheet'] = sheet
                    
                    # Append the DataFrame to the list
                    data_frames.append(df)
            
            # Update progress
            progress_message.info(f"Đang tổng hợp {i + 1} trên tổng số {total_files} file...")
        
        except Exception as e:
            st.error(f"Error processing {file.name}: {e}")

        # Final message when processing is done
    progress_message.info(f"Đã hoàn thành tổng hợp file: {total_files} file")
    
    # Merge all DataFrames into a single DataFrame
    if data_frames:
        merge = pd.concat(data_frames, ignore_index=True)
        return merge
    else:
        st.warning("No data found in the specified sheets.")
        return None


def clean_data(merge):
    progress_message = st.empty()
    progress_message.info("Bắt đầu lọc trùng...")
    """
    """
    if merge is not None:
        # Clean 'SĐT'
        merge = merge[~merge['SĐT'].isna()]
        merge.loc[:, 'SĐT'] = merge['SĐT'].astype(str)
        merge = merge.loc[~merge['SĐT'].str.contains(r'[a-zA-ZÀ-ỹ]', regex=True)]
        merge.loc[:, 'SĐT'] = (
            merge['SĐT'].astype(str)
            .str.replace(r'\.0$', '', regex=True)
        )
        merge.loc[:, 'SĐT'] = (
            merge['SĐT'].astype(str)
            .str.replace(r'[^\d/&,-]', '', regex=True)
            .str.replace(r'\.', '', regex=True)
        )
        merge = merge[merge['SĐT'].apply(lambda x: 9 <= len(str(x)) <= 25)]

        def split_sdt(value):
            if any(char in value for char in [',', '-', '&', '/']):
                return pd.Series(value.split(sep=next((char for char in [',', '-', '&', '/'] if char in value))))
            else:
                return pd.Series([value, None])
        
        merge[['SĐT_1', 'SĐT_2']] = merge['SĐT'].apply(split_sdt)
        merge['SĐT'] = merge['SĐT_1'].astype(str).str[-9:]
        merge['SĐT_2'] = merge['SĐT_2'].astype(str).str[-9:]
        merge['SĐT_2'] = merge['SĐT_2'].replace('None', '')

        # Drop intermediate column
        merge = merge.drop(columns=['SĐT_1'])

        # Clean 'Năm sinh'
        merge['Năm sinh'] = merge['Năm sinh'].fillna('').astype(str)
        merge['Năm sinh'] = merge['Năm sinh'].apply(lambda x: '' if re.search(r'[a-zA-Z]', str(x)) else x)
        merge.loc[:, 'Năm sinh'] = (
            merge['Năm sinh'].astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .str.replace(r'\.', '', regex=True)
        )
        merge.loc[merge['Năm sinh'].str.len() != 4, 'Năm sinh'] = ''

        # Clean 'Ngày gọi'
        merge['Ngày gọi'] = pd.to_datetime(merge['Ngày gọi'], errors='coerce')

        # Clean 'Level'
        merge['Level'] = merge['Level'].fillna('').astype(str)
        merge['Level'] = merge['Level'].str.replace(r'\s*CC\s*|\s*cc\s*', '', regex=True, case=False)
        merge.loc[merge['Level'].str.contains(r'[a-zA-Z]', regex=True), 'Level'] = ''
        merge['Level'] = merge['Level'].str.strip()
        merge['Level'] = merge['Level'].replace('', '0')
        merge['Level'] = merge['Level'].astype(float)

        # Clean 'Họ tên con'
        merge['Họ tên con'] = merge['Họ tên con'].str.strip().str.lower()

        # Process 'Năm sinh'
        non_empty_nam_sinh = merge[merge['Năm sinh'] != 0]
        grouped = non_empty_nam_sinh.groupby('SĐT').agg(
            nunique=('Năm sinh', 'nunique'),
            first=('Năm sinh', 'first')
        ).reset_index()
        sdt_unique_nam_sinh = grouped[grouped['nunique'] == 1][['SĐT', 'first']]
        sdt_unique_nam_sinh.rename(columns={'first': 'Năm sinh'}, inplace=True)
        sdt_to_nam_sinh = sdt_unique_nam_sinh.set_index('SĐT')['Năm sinh'].to_dict()
        sdt_unique_nam_sinh_list = sdt_unique_nam_sinh['SĐT'].tolist()
        merge['Năm sinh'] = merge['Năm sinh'].astype(str)
        filled_nam_sinh = merge['SĐT'].map(sdt_to_nam_sinh).astype(str)
        merge['Năm sinh'] = merge['Năm sinh'].where(filled_nam_sinh.isna(), filled_nam_sinh)
        merge['Năm sinh'] = merge['Năm sinh'].replace('', 0)
        merge['Năm sinh'] = pd.to_numeric(merge['Năm sinh'], errors='coerce')
        merge['Năm sinh'] = merge['Năm sinh'].replace(np.nan, 0)

        # Mapping and filling 'Năm sinh'
        result_dict = { (sdt, ho_ten): max_nam_sinh for (sdt, ho_ten), max_nam_sinh in merge.groupby(['SĐT', 'Họ tên con'])['Năm sinh'].max().items() }
        mask1 = ~merge['SĐT'].isin(sdt_unique_nam_sinh_list)
        mapping_series = pd.Series(result_dict)
        merge.loc[mask1, 'Năm sinh'] = merge.loc[mask1].apply(
            lambda row: mapping_series.get((row['SĐT'], row['Họ tên con']), row['Năm sinh']),
            axis=1
        )
        merge['Năm sinh'] = merge['Năm sinh'].astype(str).str.replace(r'\.0$', '', regex=True).replace('0', '')

        # Create 'Key' and separate 'info' and 'call'
        merge['Key'] = merge['SĐT'].astype(str) + '-' + merge['Năm sinh'].astype(str)
        info = merge[['Mã Kho', 'Tên trường', 'Tên cha/mẹ 1', 'Tên cha/mẹ 2', 'Họ tên con', 'Lớp của con', 'Địa chỉ', 'Key']].copy()
        info = info.drop_duplicates(subset='Key', keep='first').reset_index(drop=True)
        reason = merge[['Ngày gọi','Lý do KH từ chối CC2.3','Key']].copy()
        reason = reason.dropna(subset=['Lý do KH từ chối CC2.3'])
        reason = reason.sort_values(by=['Key', 'Ngày gọi'], ascending=[True, False])
        reason = reason.drop_duplicates(subset=['Key'], keep='first')
        reason = reason[['Lý do KH từ chối CC2.3','Key']]
        call = merge[['Ngày gọi', 'Ca', 'Trạng thái cuộc gọi', 'Lý do KH từ chối CC2.3', '"Kết quả ngày 1\n(Ngày - giờ gọi - note chi tiết)"',
                      '"Kết quả ngày 2\n(Ngày - giờ gọi - note chi tiết)"', 'Key']].copy()

        # Keep row with highest 'Level'
        call = call.sort_values(by=['Key', 'Ngày gọi'], ascending=[True, False])
        call = call.drop_duplicates(subset=['Key'], keep='first')
        
        # Merge date with 'call' and 'info'
        date = merge.groupby('Key').agg(
            Ngay_goi_max=('Ngày gọi', 'max'),  # Max call date
            So_lan_chia=('Key', 'size'),        # Count occurrences
            Level_max=('Level', 'max')
        ).reset_index()

        date = date.rename(columns={
            'Ngay_goi_max': 'Ngày gọi',
            'So_lan_chia': 'Số lần chia',
            'Level_max': 'Level'
})

        final = (
            date
.merge(call, how='left', left_on=['Key', 'Ngày gọi'], right_on=['Key', 'Ngày gọi'])
            .merge(info, how='left', on='Key')
.merge(reason, how='left', on='Key')

        )

        # Add 'SĐT' and 'Năm sinh' to final DataFrame
        final['SĐT'] = '0' + final['Key'].str[:9]
        mask = final['Key'].str.len() < 14
        final['Năm sinh'] = np.where(mask, '', final['Key'].str[-4:])
        sdt_to_sdt2_dict = merge.groupby('SĐT')['SĐT_2'].agg(
            lambda x: x[x != ''].iloc[0] if not x[x != ''].empty else ''
        ).to_dict()
        final['SĐT_2'] = final['SĐT'].map(sdt_to_sdt2_dict).fillna('')

        # Drop unnecessary columns
        final = final[['Mã Kho', 'Tên trường', 'Tên cha/mẹ 1', 'Tên cha/mẹ 2', 'Họ tên con',
                       'SĐT', 'SĐT_2', 'Năm sinh', 'Lớp của con', 'Địa chỉ', 'Ngày gọi','Số lần chia',
                       'Level', 'Trạng thái cuộc gọi', 'Lý do KH từ chối CC2.3',
                       '"Kết quả ngày 1\n(Ngày - giờ gọi - note chi tiết)"',
                       '"Kết quả ngày 2\n(Ngày - giờ gọi - note chi tiết)"']]

        progress_message.info("Hoàn tất!")       
        return final

    else:
        st.warning("No data to clean.")
        return None


# Streamlit UI
st.title("Lọc trùng contact K12")

# File uploader
uploaded_files = st.file_uploader("Tải lên file Excel", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    st.write(f"Tổng số file đã upload: {len(uploaded_files)}")

# LỌC TRÙNG button
if st.button("LỌC TRÙNG", key="process_button", use_container_width=True):    
    # Initialize session state messages
    # Process files
    merge = process_excel_files(uploaded_files)

    if merge is not None:
        # Clean data
        final_df = clean_data(merge)

        if final_df is not None:
            # Allow download of final DataFrame
            towrite = io.BytesIO()
            final_df.to_excel(towrite, index=False, header=True)
            towrite.seek(0)

            # Get current date and format it
            current_date = datetime.now().strftime("%d-%m-%Y")
            file_name = f"Lọc trùng contact {current_date}.xlsx"

            # Make the download button red and obvious
            st.markdown("""
                <style>
                .stDownloadButton button {
                    font-size: 20px;
                    padding: 10px 20px;
                    background-color: red;
                    color: white;
                    border-radius: 10px;
                    border: 2px solid white;
                    font-weight: bold;
                }
                .stDownloadButton button:hover {
                    background-color: darkred;
                    color: white;
                    border: 2px solid white;
                }
                </style>
            """, unsafe_allow_html=True)

            st.download_button(
                label="TẢI XUỐNG DỮ LIỆU ĐÃ LỌC TRÙNG",
                data=towrite,
                file_name=file_name,
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
    else:
        st.warning("No data to process.")