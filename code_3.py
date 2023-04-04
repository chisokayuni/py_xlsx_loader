import pandas as pd
import mysql.connector
import tkinter as tk
from tkinter import filedialog

# Define database connection parameters
db_config = {
    'host': 'localhost',
    'database': 'stud_portal_v1',
    'user': 'root',
    'password': ''
}

# Connect to MySQL database
cnx = mysql.connector.connect(**db_config)
cursor = cnx.cursor()

# Define function to handle file upload
def upload_file():
    # Create Tkinter file dialog and get selected file path
    file_path = filedialog.askopenfilename()
    
    if file_path:
        try:
            # Load data from Excel file into pandas DataFrame
            df = pd.read_excel(file_path, sheet_name='students')
            
            # Loop through rows in DataFrame and insert into MySQL database
            for i, row in df.iterrows():
                # Execute the query
                query = "INSERT INTO students (stud_id, first_name, last_name, prog_id, cohort) VALUES (%s, %s, %s, %s, %s)"
                cursor.execute(query, tuple(row))
            # Commit the changes
            cnx.commit()
            
            # Show success message
            status_label.config(text="Upload completed successfully!", fg="green")
        except Exception as e:
            # Show error message
            status_label.config(text="Upload failed: " + str(e), fg="red")
    
    # Close the cursor and connection
    cursor.close()
    cnx.close()

# Create Tkinter GUI window
root = tk.Tk()
root.title("Excel Upload Tool")

# Create heading label
heading_label = tk.Label(root, text="Upload Excel File to MySQL Database", font=("Arial", 16))
heading_label.pack(pady=20)

# Create upload button
upload_button = tk.Button(root, text="Upload Excel File", font=("Arial", 14), command=upload_file, bg="#0077be", fg="white", padx=20, pady=10)
upload_button.pack()

# Create status label
status_label = tk.Label(root, text="", font=("Arial", 12))
status_label.pack(pady=20)

# Set window size and center window on screen
window_width = 500
window_height = 250
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cord = int((screen_width / 2) - (window_width / 2))
y_cord = int((screen_height / 2) - (window_height / 2))
root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cord, y_cord))

# Start the Tkinter event loop
root.mainloop()
