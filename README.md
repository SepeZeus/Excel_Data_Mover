# Excel_Data_Mover(EDM)

**EDM** is a tool to move specific column data within excel files.

### Features

- Taking all of the data from (a) specified column(s)
- Duplicate column values
- Zip different sets column of data together
- Write/append to specific column(s) within the master file
- Create a backup of the master file each successful run of the program
- Column(s) get/write input takes in letters instead of numbers

### Requirements
* Python 3.7.4
* Openpyxl

## Usage

Upon running the program you will be prompted to drag and drop a file into the program, so that the program knows specifically where to get the file from and what file it is supposed to read. You can specify the file location by typing it out, however it is recommended that you simply drag and drop the file.

When prompted to specify the column(s) from which to get data from and the column(s) to write to, the values need to be letters more specifically they need correspond to one of the column letters within an excel file (Excel column names range from A, B, C onwards).

Only even number of pairs can be zipped together, example. A, B, C, D -> AB, CD.  E, F, G -> Error
