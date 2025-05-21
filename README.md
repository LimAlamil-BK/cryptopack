# File Compressor and Encryptor

## 📝 General
Developers: Robien Jinon, Limuelle Alamil <br />
Description: This is a Python program that compress and encrypt spreadsheet/s using the `pyzipper` module and Python's built-in `zipfile` module. It asks for two inputs: 1) <ins>directory of files</ins> to be compressed and encrypted and 2) <ins>spreadsheet</ins> containing the data (filenames and passwords) which will be used to encrypt the files.

[Download the program](https://github.com/LimAlamil-BK/cryptopack/blob/main/dist/zipper.exe)

## ❔ How to use (for end users)
1. Download the repository (https://github.com/LimAlamil-BK/cryptopack). If you can't access the repository, ask for permission to be added as collaborator.<br /><br />
![image](https://github.com/user-attachments/assets/da6a88c5-e69b-4160-9edc-39352a90debb)

2. Extract the compressed repository and find the file `zipper.exe` inside the dist folder.<br /><br />
![image](https://github.com/user-attachments/assets/7500d1d4-d1d9-469f-8dba-ac0f14df415c)

3. Before starting, make sure that you have
   - all the <ins> files (`csv`, `xlsx`, `xls`, `txt`) to be compressed and encrypted in one directory </ins> <br /><br />
   ![image](https://github.com/user-attachments/assets/2793e2c9-25e6-43c8-ba18-ad00eab6ead0)
   - the <ins> credentials (`csv`, `xlsx`, `xls`) file </ins> that contains the list of filenames and passwords ready; the filename refers to the file to be compressed and encrypted and the password refers to its encryption key  <br /><br />
   ![image](https://github.com/user-attachments/assets/cb7d0104-da8f-4dd8-8918-d9ff4f552845)
   ![image](https://github.com/user-attachments/assets/d1180908-c706-4c6b-88d0-728a56262198)

   ⚠️ Make sure that the list of filenames in the credentials file is coherent with the filenames of the actual files to be encrypted in the directory. For example, MISDATA.xls will not be compressed and encrypted if the filename in the credentials file is MSDATA.<br />
   ⚠️ Make sure that the filename and password columns in the credentials file have headers that are named "Filename" and "Password" exactly (see image above) to avoid errors.<br />
   ⚠️ To ensure that every file is both compressed and encrypted, make sure that every filename in the credentials file has a corresponding password. If a filename has no corresponding password, it will only be compressed. As such, if a file's filename is not in the credentials file, it will not be processed at all.<br />
   ⚠️ Make sure that the files to be compressed and encrypted are all in the same directory. The credentials file can be stored anywhere. Likewise, the `zipper.exe` file can be stored anywhere and it will still work.<br />

4. Double click the `zipper.exe` file to start the program.<br />
   💭 An empty command prompt window might appear upon starting the `zipper.exe` file. It appears because the executable is, by default, a console application. Just ignore that and wait for the file dialog to appear. It will probably take a few seconds. Take a sip your coffee while waiting. <br /><br />

5. A file dialog asking to select a directory will appear. Select the directory where the files to be compressed and encrypted is located.<br /><br />
   ![image](https://github.com/user-attachments/assets/55879702-40af-4499-9ffc-3fce7315ec89)
   
6. Another file dialog asking to select the credentials file (`csv`, `xlsx`, `xls`) will appear. This file is where the list of filenames and passwords are stored.<br /><br />
   ![image](https://github.com/user-attachments/assets/b7c29761-d8f1-45ca-beef-7d2dfe3a80db)

7. If the compression and encryption is successful, a toast like the one below will appear. Click OK.<br /><br />
   ![image](https://github.com/user-attachments/assets/113ae0d4-8f7e-4105-bf92-11a0dbecb29e)

8. You can try to open the files to see if the compressiona and encryption actually worked. The files will be located in a folder named `encrypted` inside the directory of the raw files you selected earlier.
   ![image](https://github.com/user-attachments/assets/cf2dabdc-ca2b-4b57-a313-6283a2972014) <br /><br />

## ❓ How to use (for developers)

### Pre-requisite knowledge/skills:<br />
[![My Skills](https://skillicons.dev/icons?i=python,git)](https://skillicons.dev)

### Instructions
  
1. Clone the repository (https://github.com/LimAlamil-BK/cryptopack) to your local device. If you can't access, ask for permission to be added as collaborator. <br /> <br />

2. The repository contains several files:
   - `dist/zipper.exe` - This is the source code as an executable file. This is accessed by the end-users.
   - `README.md` - This contains the documentation (how to use the program).
   - `zipper.py` - This is the source code. This is accessed by the developers only. Changes in the program are made here and once done, this is converted into an executable file to be used by end-users.<br /><br />
   
3. Install the following if you haven't done so in your local development environment:
   **Python** - This is the programming language used to write the source code. [How to install Python](https://www.digitalocean.com/community/tutorials/install-python-windows-10) [Download Python here](https://www.python.org/downloads/) <br />
   **Git** - This is used for version control and collaboration. [How to install Git](https://phoenixnap.com/kb/how-to-install-git-windows) <br />
   
   After installing Python, install the following Python packages as well. Install them by running `pip install <package-name>` in the terminal (e.g., `pip install pandas`).
   - `pandas` - Python's library for data analysis and manipulation. This will be used to read the spreadsheets.
   - `pyzipper` - A Python library for creating and reading password-protected ZIP files.
   - `pyinstaller` - Used to make an executable file of the source code. <br /><br />

4. To start making changes, open the repository and create a new branch by running `git branch <branch_name>` in the terminal (e.g., `git branch lim-bug-fixes`). This way, you are not directly making modification in the working source code. <br /><br />
   💭 You are most likely going to work on the source code (`zipper.py`) only. Once done with your changes, convert the source code again into an executable file. To do this, run `pyinstaller --onefile zipper.py` in the terminal. If successful, the folders `dist` and `build`, and file `zipper.spec` will appear. The executable file will be inside the `dist` folder.

5. Once done with your changes, commit your changes and push your branch to the repository. If your branch has no merge conflicts with the `dev` branch (development branch), proceed to merge. <br />
   ⚠️ Don't merge with the `main` branch unless you are absolutely sure that your branch works properly and error-free. Always merge with the `dev` branch first. If `dev` branch is non-existent, create it first by branching out from the `main` branch.
   💭 You don't have to push everything to the repository. Just the `dist/` and `zipper.py` will do.
