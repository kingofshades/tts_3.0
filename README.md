# 📅 UMT Timetable Scheduler

Welcome to the **UMT Timetable Scheduler**! This project helps you merge usage data and Excel room schedules to generate timetables for the University of Management and Technology (UMT). 🎓

## 🚀 Features:

- 📊 Load and save usage data in JSON format
- 📋 Parse Excel files for room and student capacity data
- 🗓️ Generate timetables based on room availability and constraints
- 📈 Export timetables and usage data to Excel
- 🖥️ Interactive Streamlit web interface

## 🛠️ Installation:

First, clone the repository:
```sh
git clone https://github.com/kingofshades/UMT_TIMETABLE_SCHEDULER.git
cd umt-timetable-scheduler
```

Create a virtual environment and activate it:
```sh
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
OR you can use anaconda as well
```

Install the required libraries:
```sh
pip install -r requirements.txt
```

## 📦 Required Libraries:
- ortools
- streamlit
- pandas
- openpyxl

## ▶️ Running the Application:
To start the Streamlit application, run:
```sh
streamlit run app.py
```

## 📂 Project Structure:

```sh
data/
    data_io.py
    usage_data.json
scheduling/
    solver.py
    utils.py
app.py
```

## 📜 Usage
Upload Excel: Upload an Excel file containing program courses/ roadmaps, rooms and student capacity data.
Generate Timetable: Click the button to generate the timetable based on the provided data.
Export Data: Export the generated timetables and usage data to Excel.

## 📝 Notes
Ensure that the Excel file follows the required format for room and student capacity data.
The application will display errors if there are any issues with the data or constraints.

## 🎉 Enjoy Scheduling!
Feel free to contribute to this project by submitting issues or pull requests. Happy scheduling! 😊
