# School Management System

This is a simple GUI-based school management system that allows you to manage students and teachers. The application is built using Python and Tkinter for the GUI, and it uses OpenPyXL to handle Excel files for data storage.

## Features

- Add student details (ID, first name, last name, age, grade)
- Add teacher details (ID, first name, last name, subject)
- Display all students and teachers
- Find a specific student or teacher by ID
- Save all data to an Excel file

## Requirements

- Python 3.x
- tkinter
- openpyxl

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/school-management-system.git
    cd school-management-system
    ```

2. Install the required Python packages:
    ```bash
    pip install openpyxl
    ```

3. Run the application:
    ```bash
    python app.py
    ```

## Usage

### Adding a Student

1. Enter the student details (ID, first name, last name, age, grade) in the respective fields.
2. Click the "Add Student" button.
3. The student details will be saved to the Excel file and displayed in a message box.

### Adding a Teacher

1. Enter the teacher details (ID, first name, last name, subject) in the respective fields.
2. Click the "Add Teacher" button.
3. The teacher details will be saved to the Excel file and displayed in a message box.

### Displaying All Students

1. Click the "Display Students" button.
2. A message box will appear showing all the student details.

### Displaying All Teachers

1. Click the "Display Teachers" button.
2. A message box will appear showing all the teacher details.

### Finding a Student by ID

1. Enter the student ID in the respective field.
2. Click the "Find Student" button.
3. A message box will appear showing the student details if found, or an error message if not found.

### Finding a Teacher by ID

1. Enter the teacher ID in the respective field.
2. Click the "Find Teacher" button.
3. A message box will appear showing the teacher details if found, or an error message if not found.

### Quitting the Application

1. Click the "Quit" button.
2. The application will close.

## Data Storage

The application uses OpenPyXL to save the data to an Excel file. The file contains two sheets: one for students and one for teachers.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
