# 📊 Student Data Management System (Desktop Application)

A fully functional desktop application built using Python and Tkinter that allows users to manage student records efficiently with a clean GUI and Excel integration.

---

## 🚀 Features

- ➕ Add student records (Name, Age, Class)
- ❌ Delete selected entries
- 🧹 Clear entire table with confirmation
- 📤 Export data to Excel (.xlsx)
- 💾 Save updates to existing Excel file
- ⚠️ Duplicate entry detection system
- ✅ Input validation (name, age, class)
- 🔒 Exit protection (prevents accidental data loss)

---

## 🖥️ User Interface

- Built with Tkinter and ttk Treeview
- Structured table display for student data
- Simple and user-friendly layout
- Real-time updates on actions

---

## 🧠 How It Works

### Data Entry
- User inputs Name, Age, and Class
- Validations ensure:
  - Name contains only letters
  - Age & Class are numeric and valid
- Converts class into format like:
  - 1 → 1st  
  - 2 → 2nd  
  - 3 → 3rd  
  - others → th  

---

### Data Management
- Data is stored inside a Treeview table
- Duplicate detection prompts user before inserting

---

### Excel Integration
- Uses `openpyxl` to:
  - Create Excel files
  - Write structured student data
  - Update existing files without duplication

---

### Save System
- Tracks whether file is saved or not
- Prevents accidental exit without saving
- Prompts user before closing app

---

## 📁 File Output

- Exports data in `.xlsx` format
- Columns:
  - Name
  - Age
  - Class

---

## ⚙️ Technologies Used

- Python 🐍  
- Tkinter (GUI)  
- ttk (Treeview Table)  
- openpyxl (Excel handling)

---

## 💡 Use Cases

- School student record management  
- Small-scale data entry systems  
- Beginner-friendly database alternative  
- Desktop data handling tools  

---

## 📌 Note

This is a standalone desktop application.  
No external database is required — all data is handled locally and can be exported to Excel.

---

## 👨‍💻 Developer

Ved Vatsal  
Python Developer | Automation & Web Scraping Specialist
