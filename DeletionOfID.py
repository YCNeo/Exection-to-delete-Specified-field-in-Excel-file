import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

def delete_record_and_image():
    id_to_delete = entry.get()

    if not id_to_delete:
        messagebox.showwarning("Input Error", "Please input 身分證字號.")
        return

    # Path to the Excel file and image folder
    excel_file_path = "mock_data.xlsx"
    image_folder_path = "."

    try:
        # Load the Excel file
        df = pd.read_excel(excel_file_path)

        # Check if ID exists
        if id_to_delete not in df["身分證字號"].values:
            messagebox.showinfo("Not Found", f"在 {excel_file_path} 並未找到身分證字號: {id_to_delete}")
            return
        
        # check if file exists
        image_file_path = os.path.join(image_folder_path, f"{id_to_delete}.png")
        if not os.path.exists(image_folder_path):
            messagebox.showinfo("Not Found", f"身分證字號: {id_to_delete} 相關圖片不存在")
            return

        # Find and drop the row with the given "身分證字號"
        df = df[df["身分證字號"] != id_to_delete]

        # Save the updated DataFrame back to the Excel file
        df.to_excel(excel_file_path, index=False)

        # Delete the corresponding image file
        os.remove(image_file_path)
        messagebox.showinfo("Success", f"身分證字號: {id_to_delete} 相關資料已刪除")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main application window
root = tk.Tk()
root.title("Delete Record and Image")

# Create and place the GUI components
label = tk.Label(root, text="輸入身分證字號:")
label.pack(pady=10)

entry = tk.Entry(root, width=30)
entry.pack(pady=10)

delete_button = tk.Button(root, text="Delete", command=delete_record_and_image)
delete_button.pack(pady=20)

# Start the GUI event loop
root.mainloop()
