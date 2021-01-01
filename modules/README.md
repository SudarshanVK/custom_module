# Synopsis

* This module writes a list of dictionaries into an excel spreadsheet.
* It uses the dictionary keys as the header for each column.

# Parameters

| Parameter | Choises/Default | Comments |
| - | - | - |
| path<br />string/required |   | The path to the file that needs to be modified. |
| workbook<br />string/required<br /> |   | The name of the excel spreadsheet that needs to be modified. |
| worksheet<br />string/required<br />|   | The name of the worksheet that needs to be modified.|
| data<br />list/required<br />|   |The actual facts that need to be written in the file.|
| create<br />bool| default no  | If specified, the file will be created if it does not exist.|

# Examples

Sample module execution:
```bash
- name: Write facts to spreadsheet
  xls_write:
    path: ./result
    workbook: workbook.xlsx
    worksheet: worksheet
    data: "{{ data_list }}"
    create: yes
```

Data example:
```bash
data_list:
- header1: value1
  header2:
   - test1
   - test2
  header3: value3

- header1: another_value1
  header2:
   - key1: value1
   - key2: value2
  header3: another_value

```

# Sample output
![alt text](images/Sample_xls_write.png)
