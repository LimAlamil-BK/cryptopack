import os
from os import listdir
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pyzipper
import zipfile


# Function to read the list of filenames from the credentials file
# credentials file contains the corresponding password of each file (filename)
def read_filename(credentials):
    try:
        # Try reading as CSV
        data = pd.read_csv(credentials.name)
        filenames = data['Filename'].astype(str).tolist()
        return filenames
    except:
        # If CSV fails, try reading as Excel
        data = pd.read_excel(credentials.name, dtype='object')
        filenames = data['Filename'].astype(str).tolist()
        return filenames


# Function to read the list of passwords from the credentials file
def read_password(credentials):
    try:
        # Try reading as CSV
        data = pd.read_csv(credentials.name)
        passwords = data['Password'].astype(str).tolist()
        return passwords
    except:
        # If CSV fails, try reading as Excel
        data = pd.read_excel(credentials.name, dtype='object')
        passwords = data['Password'].astype(str).tolist()

        return passwords


# Function to compress and encrypt matching files with passwords
def compress_file(path, sheets, filenames, passwords):

    counter = 0 # Count successful compressions


    try:
        # For every spreadsheet that needs to be compressed and encrypted...
        for sheet in sheets:

            # Strip extension from sheet name to get the base filename
            filename = ""
            for ext in ('.xls', '.xlsx', '.csv', '.txt'):
                if sheet.endswith(ext):
                    filename = sheet.removesuffix(ext)
                    break
            

            # Get the corresponding password of the spreadsheet based on its filename
            try:
                index = filenames.index(filename)
            except ValueError:
                continue    # Skip if filename not found

            password = passwords[index]

            # Get the directory of the spreadsheet that needs to be compressed and encrypted
            sheet_file_path = os.path.join(path, sheet)


            # Define subdirectory to store encrypted files
            output_folder = os.path.join(path, "encrypted")
            # Create the folder if it doesn't exist
            os.makedirs(output_folder, exist_ok=True)
            # Full path to the output zip file
            zip_file_path = os.path.join(output_folder, f"{filename}.zip")


            # If a password is provided (i.e., not NaN), compress and encrypt the file
            if ((filename in filenames) and str(password).lower() != 'nan'):
                try:
                    # Compress and encrypt the spreadsheet
                    with pyzipper.AESZipFile(zip_file_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                        zf.setpassword(password.encode())               # password must be bytes
                        zf.setencryption(pyzipper.WZ_AES, nbits=256)    # AES-256
                        zf.write(sheet_file_path, arcname=sheet)        # arcname ensures correct name inside zip
                    counter += 1    # Successfully processed

                except Exception as e:
                    print(f"Error compressing/encrypting {sheet}: {e}")
                    continue
            # If no password is provided, compress only the file
            else:
                try:
                    # Compress the spreadsheet
                    with zipfile.ZipFile(zip_file_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                        zf.write(sheet_file_path, arcname=sheet)        # arcname ensures correct name inside zip
                    counter += 1    # Successfully processed

                except Exception as e:
                    print(f"Error compressing {sheet}: {e}")
                    continue
        
        # Show success message after processing
        if counter > 0:
            tk.messagebox.showinfo(title = "Success", message = "Files compressed and/or encrypted successfully. " + str(counter) + " files compressed and encrypted.")
        else:
            tk.messagebox.showinfo(title = "Success", message = "No files compressed and/or encrypted.")

    except Exception as error:
        # Show warning message if something goes wrong
        message = "Error: " + str(type(error).__name__)
        message = "Error: " + str(error)

        tk.messagebox.showwarning(title = "Error", message = message)


# Main function to run the program
def main():
    root = tk.Tk()
    root.withdraw()

    # Ask user to select the folder containing sheets
    path = filedialog.askdirectory(title="Select folder of files to compress and encrypt")

    # Ask user to select the credentials file
    credentials = filedialog.askopenfile(
        filetypes = [('All files', '*'),('.csv', '*.csv'), ('.xlsx', '*.xlsx')],
        title="Select credentials file"
    )

    # Collect valid files with supported extensions
    sheets = [filename for filename in listdir(path) if filename.endswith((".xls", ".xlsx", ".csv", ".txt"))]

    # Extract filename-password mappings from the credentials file
    filenames = read_filename(credentials)
    passwords = read_password(credentials)

    # Begin compression process
    compress_file(path, sheets, filenames, passwords)

# Run the script
main()