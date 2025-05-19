# File Compressor and Encryptor

## 📝 General
Developers: Robien Jinon, Limuelle Alamil <br />
Description: This is a Python program that compress and encrypt spreadsheet/s using the `pyzipper` module and Python's built-in `zipfile` module. It asks for two inputs: 1) directory of files to be compressed and encrypted and 2) spreadsheet containing the data (filenames and passwords) which will be used to encrypt the files.

[Download the program](https://github.com/LimAlamil-BK/cryptopack/blob/main/dist/zipper.exe)

## ❔ How to use (for end users)
1. Download the repository (https://github.com/LimAlamil-BK/cryptopack). If you can't access the repository, ask for permission to be added as collaborator.<br /><br />
![image](https://github.com/user-attachments/assets/da6a88c5-e69b-4160-9edc-39352a90debb)

2. Extract the compressed repository and find the file `zipper.exe` inside the dist folder.<br /><br />
![image](https://github.com/user-attachments/assets/7500d1d4-d1d9-469f-8dba-ac0f14df415c)

3. Before starting, make sure that you have
   - all the <ins> files (`csv`, `xlsx`, `xls`, `txt`) to be compressed and encrypted in one directory </ins> <br />
   ![image](https://github.com/user-attachments/assets/2793e2c9-25e6-43c8-ba18-ad00eab6ead0)
   - the <ins> credentials (`csv`, `xlsx`, `xls`) file </ins> that contains the list of filenames and passwords ready; the filename refers to the file to be compressed and encrypted and the password refers to its encryption key  <br />
   ![image](https://github.com/user-attachments/assets/cb7d0104-da8f-4dd8-8918-d9ff4f552845)
   ![image](https://github.com/user-attachments/assets/d1180908-c706-4c6b-88d0-728a56262198)

   ⚠️ NOTE: Make sure that the list of filenames in the credentials file is coherent with the filenames of the actual files to be encrypted in the directory. <br />
   ⚠️ NOTE: Make sure that the filename and password columns in the credentials file have headers that are named "Filename" and "Password" exactly (see image above) to avoid errors. <br />
   ⚠️ NOTE: Make sure that the files to be compressed and encrypted are all in the same directory. The credentials file can be stored anywhere. Likewise, the `zipper.exe` file can be stored anywhere and it will still work.<br />

4. Double click the `zipper.exe` file to start the program.<br />
5. A file dialog asking to select a directory will appear. Select the directory where the files to be compressed and encrypted is located.<br /><br />

