import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

file1_path = 'E:/SmartIC/to_excel/SS2MN01_021.xlsx'
file2_path = 'E:/SmartIC/to_excel/SS2MN01_022.xlsx'
image_folder_base1 = 'E:/SmartIC/to_excel/SS2MN01_021/'
image_folder_base2 = 'E:/SmartIC/to_excel/SS2MN01_022/'
output_file_path = 'E:/SmartIC/to_excel/021_022_merged_defect_data.xlsx'

df1 = pd.read_excel(file1_path)
df2 = pd.read_excel(file2_path)

merge_keys = ['INDEX', 'LINE', 'VISIONTYPE', 'DEFECTID', 'DEFECTNAME', 'COLORIMAGE', 'VERIFYIMAGE']

df2_renamed = df2.rename(columns={'COLORIMAGE': 'SECOND_COLORIMAGE', 'VERIFYIMAGE': 'SECOND_VERIFYIMAGE'})

merged_df = pd.merge(df1, df2_renamed, left_on=merge_keys, right_on=['INDEX', 'LINE', 'VISIONTYPE', 'DEFECTID', 'DEFECTNAME', 'SECOND_COLORIMAGE', 'SECOND_VERIFYIMAGE'], how='outer', indicator=True)

merged_df = merged_df[merge_keys + ['SECOND_COLORIMAGE', 'SECOND_VERIFYIMAGE', '_merge']]

merged_df.to_excel(output_file_path, index=False)

wb = load_workbook(output_file_path)
ws = wb.active

def add_image(ws, img_path, cell, scale=1.0):
    if pd.notna(img_path) and os.path.exists(img_path):
        img = Image(img_path)
        img.width = img.width * scale
        img.height = img.height * scale
        ws.add_image(img, cell)
        return img.width, img.height
    return None, None

def set_cell_size(ws, col, row, width, height):
    ws.column_dimensions[col].width = width / 7.5
    ws.row_dimensions[row].height = height * 0.75

image_folders1 = {
    'TOP': os.path.join(image_folder_base1, 'TOP'),
    'BOTTOM': os.path.join(image_folder_base1, 'BOTTOM'),
    'MONO': os.path.join(image_folder_base1, 'MONO')
}

image_folders2 = {
    'TOP': os.path.join(image_folder_base2, 'TOP'),
    'BOTTOM': os.path.join(image_folder_base2, 'BOTTOM'),
    'MONO': os.path.join(image_folder_base2, 'MONO')
}

for idx, row in merged_df.iterrows():
    if row['_merge'] in ['both', 'left_only']:
        if pd.notna(row["COLORIMAGE"]):
            color_image_path1 = os.path.join(image_folders1.get(row["VISIONTYPE"], image_folder_base1), row["COLORIMAGE"])
            col_row = f'F{idx + 2}'
            width, height = add_image(ws, color_image_path1, col_row)
            if width and height:
                set_cell_size(ws, 'F', idx + 2, width, height)

        if pd.notna(row["VERIFYIMAGE"]):
            verify_image_path1 = os.path.join(image_folders1.get(row["VISIONTYPE"], image_folder_base1), row["VERIFYIMAGE"])
            col_row = f'G{idx + 2}'
            width, height = add_image(ws, verify_image_path1, col_row)
            if width and height:
                set_cell_size(ws, 'G', idx + 2, width, height)
    
    if row['_merge'] in ['both', 'right_only']:
        if pd.notna(row["SECOND_COLORIMAGE"]):
            color_image_path2 = os.path.join(image_folders2.get(row["VISIONTYPE"], image_folder_base2), row["SECOND_COLORIMAGE"])
            col_row = f'H{idx + 2}'
            width, height = add_image(ws, color_image_path2, col_row)
            if width and height:
                set_cell_size(ws, 'H', idx + 2, width, height)

        if pd.notna(row["SECOND_VERIFYIMAGE"]):
            verify_image_path2 = os.path.join(image_folders2.get(row["VISIONTYPE"], image_folder_base2), row["SECOND_VERIFYIMAGE"])
            col_row = f'I{idx + 2}'
            width, height = add_image(ws, verify_image_path2, col_row)
            if width and height:
                set_cell_size(ws, 'I', idx + 2, width, height)

wb.save(output_file_path)
