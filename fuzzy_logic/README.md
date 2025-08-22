

## Description
This script uses fuzzy logic to help users find the correct researcher even when they make typos, omit words, or only enter a partial name or email address.

If the input is incomplete (e.g., only a first or last name), the script can still find potential matches.

If multiple matches are found, the script will prompt the user to provide additional information to narrow down the result.


## Installation
Install the required libraries, use the following commands:
<pre>
 ```bash 
 pip install rapidfuzz 
 pip install pandas 
 pip install openpyxl 
 ```
</pre>

## Import necessary libraries
```python
import pandas as pd
from rapidfuzz import fuzz, process
```


## Excel File Requirements
The Excel file should contain the following columns:

- `First_Name`
- `Last_Name`
- `Employee_Full_Name`
- `Email_Address`
- `Job_Name`
- `Department`
- `School`

## How to Use
 1. Place the Excel file in the same directory as the script.
 2. Run the script
 3. Enter a name or email when prompted.
 4. If multiple matches are found, enter more information to narrow down the result.

 ## Features
- Fuzzy matching for names and emails
- Handles typos and partial input
- Interactive refinement for multiple matches
- Excel-based data input for easy updates
