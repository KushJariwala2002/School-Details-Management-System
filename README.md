# School Details Management System

A Python GUI application built using **Tkinter**, designed to manage school details such as city, school name, principal name, address, contact number, and email. The data is fetched from and saved into an **Excel file**. Users can interact with the records using various options, including fetch, update/insert, and delete.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Technologies Used](#technologies-used)
- [Screenshots](#screenshots)
- [Logging](#logging)
- [Contribution](#contribution)
- [License](#license)

---

## Features

- **Dynamic Excel File Selection**: Users can select an Excel file during runtime to manage school records.
- **Sheet Selection**: Ability to choose the desired worksheet from the selected Excel file.
- **Fetch School Details**: Retrieve and display school details using the **Sr.No** as a reference.
- **Insert or Update School Details**: Modify existing school information or insert new records.
- **Delete School Details**: Remove entries based on **Sr.No**.
- **User-Friendly GUI**: Simple and intuitive interface using **Tkinter** with visual feedback for actions.
- **Logo Integration**: Ability to customize the app with a logo.
- **Logging**: Integrated logging to track actions, errors, and application events.
- **Input Validation**: Ensures that required fields are correctly filled, and critical inputs are properly validated.
- **Responsive Design**: Color-coded success and error feedback in input fields for user actions.

---

## Installation

### Prerequisites

- **Python 3.x** installed on your system.
- Required Python packages listed in the `requirements.txt`.

### Step-by-Step Instructions

1. Clone the repository:
    ```bash
    git clone https://github.com/your-repo-url/School-Details-Management.git
    ```
2. Navigate into the project directory:
    ```bash
    cd School-Details-Management
    ```
3. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```
4. Run the application:
    ```bash
    python excel.py
    ```

### Running the Application

1. Upon launching the app, you will be prompted to select an **Excel file**.
2. Choose the sheet within the file to manage school details.
3. Use the **Fetch, Update, Delete**, or **Insert** options to interact with the data.

---

## Usage

- **Select Excel File**: Upon launching, the user is prompted to select an Excel file containing school details.
- **Choose Worksheet**: From the dropdown, select the worksheet in the Excel file where the school details are stored.
- **Fetch School Details**: Enter the **Sr.No** and click "Fetch Details" to retrieve information from the Excel sheet.
- **Update or Insert School Details**: Modify the displayed fields and click "Update/Insert School" to save the changes to the Excel file.
- **Delete School Details**: Enter the **Sr.No** of the record to be deleted and click "Delete School" to remove it.
- **Clear Fields**: Click "Clear Fields" to reset all input fields.

---

## Technologies Used

- **Python 3.x**
- **Tkinter**: For the graphical user interface.
- **Pillow (PIL)**: For image handling and logo integration.
- **OpenPyXL**: For reading from and writing to Excel files.
- **OS**: For file path handling.
- **Logging**: To record user actions and errors.

---

## Screenshots

Here are some screenshots showcasing the functionality:

### Main Interface:

![Main Interface](./images/main-interface.png)


---

## Logging

This application integrates a logging feature to record:
- User actions such as fetch, insert, update, and delete operations.
- Errors and exceptions for debugging and user feedback.
- Logging messages are saved into a `logfile.txt` in the project directory.

---

## Contribution

Feel free to contribute by opening issues or submitting pull requests. Please ensure your code follows best practices and includes relevant test cases.

---

## License

This project is licensed under the MIT License.
