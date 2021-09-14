# Synopsis

* This module reads a multi-tabbed excel spread sheet into an iterable list of dictionaries

# Parameters

The module accepts just one parameter as input
*src: path to the source excel spreadsheet.

# Examples

Sample module execution:
```bash
- name: Read facts from spreadsheet
  xls_facts:
    src: source.xlsx
  register: excel_data
```

Output example:
```bash
    "ansible_facts": {
        "spreadsheet_Sheet1": [
            {
                "Hostname": "Switch-1",
                "Mgmt_ip": "10.0.0.1"
            },
            {
                "Hostname": "Switch-2",
                "Mgmt_ip": "10.0.0.2"
            },
            {
                "Hostname": "Switch-3",
                "Mgmt_ip": "10.0.0.3"
            }
        ],
        "spreadsheet_Sheet2": [
            {
                "Description": "To Spine-1",
                "Interface": "Ethernet1/1",
                "Interface_IP": "192.168.100.1/30"
            },
            {
                "Description": "To Spine-2",
                "Interface": "Ethernet1/2",
                "Interface_IP": "192.168.100.5/30"
            }
        ]
    },
```
