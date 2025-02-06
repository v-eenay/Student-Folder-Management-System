# Student Folder Management System

[![Python](https://img.shields.io/badge/python-3.x-blue.svg)](https://www.python.org/)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/v-eenay/Student-Folder-Management-System/graphs/commit-activity)

## 🎯 Overview

The **Student Folder Management System** is a robust Python-based automation tool that streamlines the organization of student data into structured folders and files. This system efficiently processes Excel sheets containing student information and group assignments, creating a well-organized hierarchical folder structure based on sections and groups, while generating personalized Excel files for each student.

## ✨ Key Features

- 📁 **Automated Folder Creation**: Creates section and group folders dynamically based on input data
- 📊 **Smart File Generation**: Generates individual Excel files for each student with pre-filled information
- 🔍 **Intelligent Name Matching**: Implements case-insensitive matching for accurate student grouping
- 📝 **Detailed Logging**: Maintains comprehensive logs for process tracking and troubleshooting

## 🛠️ Requirements

- Python 3.x
- Required Libraries:
  - `openpyxl` - Excel file handling
  - `os` - Operating system operations
  - `shutil` - File operations

### Installation

bash
pip install openpyxl


## 🚀 Getting Started

1. **Clone the Repository**

   ```bash
   git clone https://github.com/v-eenay/Student-Folder-Management-System.git
   cd Student-Folder-Management-System
   ```

2. **Data Preparation**

   Prepare an Excel file (`student_list.xlsx`) with two sheets:
   - `StudentList`: Student details (University ID, Name, Section)
   - `VivaGroups`: Group assignments for students

3. **Execute the Script**

   ```bash
   python manage_student_folders.py
   ```

4. **View Results**

   The system will generate a `folders` directory containing organized section and group folders with individual student files.

## 📂 Directory Structure


folders/<br>
├── L1C1/<br>
│   ├── L1C1G1/<br>
│   │   ├── 23087654 JOHN SMITH.xlsx<br>
│   │   └── 23087699 EMMA WILSON.xlsx<br>
│   └── L1C1G2/<br>
│       ├── 23087823 MICHAEL BROWN.xlsx<br>
│       └── ...<br>
├── L1C2/<br>
│   ├── L1C2G1/<br>
│   │   └── ...<br>
│   └── ...<br>
└── ...<br>

## 🤝 Contributing

We welcome contributions! Follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request


## 📞 Contact

For support or inquiries:

- **Developer**: Vinay Koirala
- **Email**: koiralavinay@gmail.com
- **GitHub**: [@v-eenay](https://github.com/v-eenay)

---

<div align="center">
  <sub>Built with ❤️ by Vinay Koirala</sub>
  <br>
  © 2023 Vinay Koirala. All rights reserved.
</div>