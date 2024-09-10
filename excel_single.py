import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
from tkinter import Tk, filedialog, messagebox

def process_file(data_file, file_path, image_folder_base):
    df = pd.read_csv(data_file)
    filtered_df = df[df['DELEGATE'] == 'V'].reset_index(drop=True)
    filtered_df.to_excel(file_path, index=False)

    wb = load_workbook(file_path)
    ws = wb.active

    def add_image(ws, img_path, cell, scale=1.0):
        if os.path.exists(img_path):
            img = Image(img_path)
            img.width = img.width * scale
            img.height = img.height * scale
            ws.add_image(img, cell)
            return img.width, img.height
        return None, None

    def set_cell_size(ws, col, row, width, height):
        ws.column_dimensions[col].width = width / 7.5
        ws.row_dimensions[row].height = height * 0.75

    image_folders = {
        'TOP': os.path.join(image_folder_base, 'TOP'),
        'BOTTOM': os.path.join(image_folder_base, 'BOTTOM'),
        'MONO': os.path.join(image_folder_base, 'MONO')
    }

    for idx, row in filtered_df.iterrows():
        color_image_path = os.path.join(image_folders.get(row["VISIONTYPE"], image_folder_base), row["COLORIMAGE"])
        verify_image_path = os.path.join(image_folders.get(row["VISIONTYPE"], image_folder_base), row["VERIFYIMAGE"]) if pd.notna(row["VERIFYIMAGE"]) else None

        col_row = f'O{idx + 2}'
        width, height = add_image(ws, color_image_path, col_row)
        if width and height:
            set_cell_size(ws, 'O', idx + 2, width, height)

        if verify_image_path:
            col_row = f'P{idx + 2}'
            width, height = add_image(ws, verify_image_path, col_row)
            if width and height:
                set_cell_size(ws, 'P', idx + 2, width, height)

    wb.save(file_path)

def main():
    root = Tk()
    root.withdraw()

    data_file = filedialog.askopenfilename(title="Select the data file (TXT)", filetypes=[("Text files", "*.txt")])
    if not data_file:
        messagebox.showerror("Error", "No data file selected!")
        return

    file_path = filedialog.asksaveasfilename(title="Save the Excel file as", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showerror("Error", "No save path selected!")
        return

    image_folder_base = filedialog.askdirectory(title="Select the base image folder")
    if not image_folder_base:
        messagebox.showerror("Error", "No image folder selected!")
        return

    process_file(data_file, file_path, image_folder_base)
    messagebox.showinfo("Success", f"Excel file created successfully at {file_path}")

if __name__ == "__main__":
    main()
