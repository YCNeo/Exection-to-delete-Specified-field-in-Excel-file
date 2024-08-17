from tkinter import Tk, ttk, Label, Button, messagebox as msgBox
from pandas import read_excel
from os import path, listdir, remove

def deleteRecordAndImage():
    selectedFile = combo.get()
    targeID = path.splitext(selectedFile)[0]  # Remove the file extension

    if not targeID:
        msgBox.showwarning("Input Error", "請輸入身分證字號.")
        return

    # Path to the Excel file and image folder
    excelPath = "證號清冊.xlsx"
    folderPath = "."

    try:
        # Load the Excel file
        df = read_excel(excelPath)

        # Check if ID exists
        if targeID not in df["身分證字號"].values:
            msgBox.showinfo("Not Found", f"在 {excelPath} 中並未找到身分證字號: {targeID}")
            return
        
        # check if file exists
        image_file_path = path.join(folderPath, f"{selectedFile}")
        if not path.exists(folderPath):
            msgBox.showinfo("Not Found", f"身分證字號: {targeID} 相關圖片不存在")
            return

        # Find and drop the row with the given "身分證字號"
        df = df[df["身分證字號"] != targeID]

        # Save the updated DataFrame back to the Excel file
        df.to_excel(excelPath, index=False)

        # Delete the corresponding image file
        remove(image_file_path)
        msgBox.showinfo("Success", f"身分證字號: {targeID} 相關資料已刪除")
    
    except Exception as e:
        msgBox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main application window
root = Tk()
root.title("刪除資料")

# Set the root window size (width and height)
root.geometry("270x150")

# Create and place the GUI components
label = Label(root, text="選擇圖片:")
label.pack(pady=10)

# List all files in the current directory
folderPath = "."
allFiles = listdir(folderPath)

# Filter the list to include only .png and .jpg image files
imageFiles = [file for file in allFiles if file.endswith((".png", ".jpg", ".svg", ".pdf", ".gif"))]

# Create the drop-down menu (combobox)
combo = ttk.Combobox(root, values=imageFiles, width=25)
combo.pack(pady=10)

deleteButton = Button(root, text="刪除", command=deleteRecordAndImage)
deleteButton.pack(pady=20)

# Start the GUI event loop
root.mainloop()
