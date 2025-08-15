# pandabook

## Data Export:

* Export a collection of pandas dataframes to an Excel workbook with advanced styling options.
* You can export many dataframes into the same file. You'll get an index page with links to each tab.
* Handles arbitrary data types, such as lists, dictionaries, uuid, and others.
* Handles multi-level indexes
* Can do conditional formatting based on values.
* Allows for granular control over each column, such as widths and styles.


    ![Index table](media/img/index_table.png)

    ![Mixed values](media/img/mixed_values.png)

    ![Multi-level index](media/img/multi_level_index.png)



## Data Import:

* You can also load a collection of named tables from an Excel file back to dataframes.

    ![Data import](media/img/data_import.png)



## Setup instructions

### 1. Clone the repository

```bash
git clone <repository-url>
```

### 2. Install dependencies

This project uses [`uv`](https://github.com/astral-sh/uv) for dependency management. Make sure `uv` is installed on your system.

To install dependencies:

```bash
uv venv
uv pip install .
```
To activate the virtual environment:
```bash
.venv\Scripts\activate
```
After this you can run the Jupyter notebook to see worked examples.

## More info:

[Refer to example.ipynb for worked examples.](example.ipynb)
