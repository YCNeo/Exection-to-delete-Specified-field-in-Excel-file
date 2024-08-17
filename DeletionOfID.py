from tkinter import Tk, ttk, Label, Button, messagebox
from pandas import read_excel
from os import path, listdir

def delete_record_and_image():
    selected_file = combo.get()
    id_to_delete = path.splitext(selected_file)[0]  # Remove the file extension

    if not id_to_delete:
        messagebox.showwarning("Input Error", "請輸入身分證字號.")
        return

    # Path to the Excel file and image folder
    excel_file_path = "證號清冊.xlsx"
    image_folder_path = "."

    try:
        # Load the Excel file
        df = read_excel(excel_file_path)

        # Check if ID exists
        if id_to_delete not in df["身分證字號"].values:
            messagebox.showinfo("Not Found", f"在 {excel_file_path} 中並未找到身分證字號: {id_to_delete}")
            return
        
        # check if file exists
        image_file_path = path.join(image_folder_path, f"{selected_file}")
        if not path.exists(image_folder_path):
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
root = Tk()
root.title("刪除資料")

# Set the root window size (width and height)
root.geometry("270x150")

# Create and place the GUI components
label = Label(root, text="選擇圖片:")
label.pack(pady=10)

# List all files in the current directory
image_folder_path = "."
all_files = listdir(image_folder_path)

# Filter the list to include only .png and .jpg image files
image_files = [file for file in all_files if file.endswith((".png", ".jpg", ".svg", ".pdf", ".gif"))]

# Create the drop-down menu (combobox)
combo = ttk.Combobox(root, values=image_files, width=25)
combo.pack(pady=10)

delete_button = Button(root, text="刪除", command=delete_record_and_image)
delete_button.pack(pady=20)

# Start the GUI event loop
root.mainloop()
