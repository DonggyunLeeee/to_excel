import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

file_paths = [
    'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_10.xlsx',
    'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_12.xlsx',
    'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_14.xlsx'
]

image_folders = [
    {
        'base': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_10/',
        'TOP': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_10/TOP',
        'BOTTOM': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_10/BOTTOM',
        'MONO': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_10/MONO'
    },
    {
        'base': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_12/',
        'TOP': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_12/TOP',
        'BOTTOM': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_12/BOTTOM',
        'MONO': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_12/MONO'
    },
    {
        'base': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_14/',
        'TOP': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_14/TOP',
        'BOTTOM': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_14/BOTTOM',
        'MONO': 'C:/Repo/SmartIC/to_excel/M08P2429B-112-0-4_14/MONO'
    }
]

output_file_path = 'C:/Repo/SmartIC/to_excel/merged_defect_data.xlsx'

def read_excel_files(file_paths):
    dfs = []
    for file_path in file_paths:
        dfs.append(pd.read_excel(file_path))
    return dfs

def merge_dataframes(dfs, merge_keys):
    df_merged = dfs[0][merge_keys + ['COLORIMAGE', 'VERIFYIMAGE']]
    for i, df in enumerate(dfs[1:], start=2):
        df_renamed = df[merge_keys + ['COLORIMAGE', 'VERIFYIMAGE']].rename(
            columns={'COLORIMAGE': f'COLORIMAGE_{i}', 'VERIFYIMAGE': f'VERIFYIMAGE_{i}'})
        df_merged = pd.merge(df_merged, df_renamed, on=merge_keys, how='outer')
    return df_merged

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

def insert_images(ws, merged_df, image_folders, merge_keys):
    for idx, row in merged_df.iterrows():
        for i in range(1, len(image_folders) + 1):
            color_image_col = f'COLORIMAGE_{i}' if i > 1 else 'COLORIMAGE'
            verify_image_col = f'VERIFYIMAGE_{i}' if i > 1 else 'VERIFYIMAGE'
            
            if pd.notna(row[color_image_col]):
                color_image_path = os.path.join(image_folders[i-1].get(row["VISIONTYPE"], image_folders[i-1]['base']), row[color_image_col])
                col_row = f'{chr(70 + 2*(i-1))}{idx + 2}'
                width, height = add_image(ws, color_image_path, col_row)
                if width and height:
                    set_cell_size(ws, chr(70 + 2*(i-1)), idx + 2, width, height)

            if pd.notna(row[verify_image_col]):
                verify_image_path = os.path.join(image_folders[i-1].get(row["VISIONTYPE"], image_folders[i-1]['base']), row[verify_image_col])
                col_row = f'{chr(71 + 2*(i-1))}{idx + 2}'
                width, height = add_image(ws, verify_image_path, col_row)
                if width and height:
                    set_cell_size(ws, chr(71 + 2*(i-1)), idx + 2, width, height)

def main(file_paths, image_folders, output_file_path, merge_keys=['INDEX', 'LINE', 'VISIONTYPE', 'DEFECTID', 'DEFECTNAME']):
    dfs = read_excel_files(file_paths)
    merged_df = merge_dataframes(dfs, merge_keys)
    merged_df.to_excel(output_file_path, index=False)

    wb = load_workbook(output_file_path)
    ws = wb.active

    insert_images(ws, merged_df, image_folders, merge_keys)

    wb.save(output_file_path)

main(file_paths, image_folders, output_file_path)
