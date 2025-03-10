import openpyxl
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import messagebox

# Initialize the main application window
ctk.set_appearance_mode("Dark")  # Set to Dark mode for white text
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

# Function to open the payroll details window
def open_payroll_window():
    # Close the main window
    app.destroy()

    # Create a new window for payroll details
    payroll_window = ctk.CTk()
    payroll_window.title("Employee Payroll Details")
    payroll_window.geometry("960x740")  # Set window size to 960x740

    # Initialize a scrollable frame for multiple employees
    scrollable_frame = ctk.CTkScrollableFrame(payroll_window, fg_color="transparent")
    scrollable_frame.pack(fill="both", expand=True, padx=20, pady=20)

    try:
        # Load the workbook and select the active sheet
        book = load_workbook('data.xlsx')
        sheet = book.active

        # Define the fields and their corresponding cell addresses
        fields = [
            ("Employee Name", "A"),
            ("Employee ID", "B"),
            ("Department", "C"),
            ("Designation", "D"),
            ("Date of Joining", "E"),
            ("Date of Birth", "F"),
            ("UAN", "G"),
            ("PF No", "H"),
            ("ESI No", "I"),
            ("Basic Salary", "J"),
            ("Conveyance", "K"),
            ("Special Allowance", "L"),
            ("PF Deduction", "M"),
            ("ESI Deduction", "N"),
            ("PT Deduction", "O"),
        ]

        # Iterate through each row (employee) in the Excel sheet
        for row in range(2, sheet.max_row + 1):  # Start from row 2 (skip header)
            # Read the values from the Excel sheet
            data = {}
            for field, col in fields:
                cell = f"{col}{row}"
                data[field] = sheet[cell].value

            # Calculate total earnings, deductions, and net pay
            total_earnings = data["Basic Salary"] + data["Conveyance"] + data["Special Allowance"]
            total_deductions = data["PF Deduction"] + data["ESI Deduction"] + data["PT Deduction"]
            net_pay = total_earnings - total_deductions

            # Create a frame for each employee
            employee_frame = ctk.CTkFrame(scrollable_frame, fg_color="#2E2E2E")  # Dark grey
            employee_frame.pack(fill="x", padx=10, pady=10)

            # Display the results in three columns
            # Column 1: Employee Details
            details = [
                ("Employee Name", data["Employee Name"]),
                ("Employee ID", data["Employee ID"]),
                ("Department", data["Department"]),
                ("Designation", data["Designation"]),
                ("Date of Joining", data["Date of Joining"]),
                ("Date of Birth", data["Date of Birth"]),
                ("UAN", data["UAN"]),
                ("PF No", data["PF No"]),
                ("ESI No", data["ESI No"]),
            ]

            details_frame = ctk.CTkFrame(employee_frame, fg_color="transparent")
            details_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

            ctk.CTkLabel(
                details_frame, 
                text="Details", 
                font=("Helvetica", 16, "bold"),  # Bold font for headers
                text_color="white"  # White color for headers
            ).pack(pady=10)

            for field, value in details:
                row_frame = ctk.CTkFrame(details_frame, fg_color="transparent")
                row_frame.pack(fill="x", padx=10, pady=5)
                ctk.CTkLabel(
                    row_frame, 
                    text=field, 
                    font=("Helvetica", 14),  # Regular font for fields
                    text_color="white"  # White color for fields
                ).pack(side="left", padx=10)
                ctk.CTkLabel(
                    row_frame, 
                    text=value, 
                    font=("Helvetica", 14),  # Regular font for values
                    text_color="white"  # White color for values
                ).pack(side="right", padx=10)

            # Column 2: Earnings
            earnings = [
                ("Basic Salary", f"₹{data['Basic Salary']}"),
                ("Conveyance", f"₹{data['Conveyance']}"),
                ("Special Allowance", f"₹{data['Special Allowance']}"),
            ]

            earnings_frame = ctk.CTkFrame(employee_frame, fg_color="transparent")
            earnings_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

            ctk.CTkLabel(
                earnings_frame, 
                text="Earnings", 
                font=("Helvetica", 16, "bold"),  # Bold font for headers
                text_color="white"  # White color for headers
            ).pack(pady=10)

            for field, value in earnings:
                row_frame = ctk.CTkFrame(earnings_frame, fg_color="transparent")
                row_frame.pack(fill="x", padx=10, pady=5)
                ctk.CTkLabel(
                    row_frame, 
                    text=field, 
                    font=("Helvetica", 14),  # Regular font for fields
                    text_color="white"  # White color for fields
                ).pack(side="left", padx=10)
                ctk.CTkLabel(
                    row_frame, 
                    text=value, 
                    font=("Helvetica", 14),  # Regular font for values
                    text_color="white"  # White color for values
                ).pack(side="right", padx=10)

            # Column 3: Deductions
            deductions = [
                ("PF Deduction", f"₹{data['PF Deduction']}"),
                ("ESI Deduction", f"₹{data['ESI Deduction']}"),
                ("PT Deduction", f"₹{data['PT Deduction']}"),
            ]

            deductions_frame = ctk.CTkFrame(employee_frame, fg_color="transparent")
            deductions_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

            ctk.CTkLabel(
                deductions_frame, 
                text="Deductions", 
                font=("Helvetica", 16, "bold"),  # Bold font for headers
                text_color="white"  # White color for headers
            ).pack(pady=10)

            for field, value in deductions:
                row_frame = ctk.CTkFrame(deductions_frame, fg_color="transparent")
                row_frame.pack(fill="x", padx=10, pady=5)
                ctk.CTkLabel(
                    row_frame, 
                    text=field, 
                    font=("Helvetica", 14),  # Regular font for fields
                    text_color="white"  # White color for fields
                ).pack(side="left", padx=10)
                ctk.CTkLabel(
                    row_frame, 
                    text=value, 
                    font=("Helvetica", 14),  # Regular font for values
                    text_color="white"  # White color for values
                ).pack(side="right", padx=10)

            # Summary Section
            summary_frame = ctk.CTkFrame(employee_frame, fg_color="transparent")
            summary_frame.pack(fill="x", padx=20, pady=10)

            ctk.CTkLabel(
                summary_frame, 
                text="Summary", 
                font=("Helvetica", 18, "bold"),  # Larger and bold font for summary header
                text_color="white"  # White color for summary header
            ).pack(pady=10)

            summary_data = [
                ("Total Earnings", f"₹{total_earnings}"),
                ("Total Deductions", f"₹{total_deductions}"),
                ("Net Pay", f"₹{net_pay}"),
            ]

            for field, value in summary_data:
                row_frame = ctk.CTkFrame(summary_frame, fg_color="transparent")
                row_frame.pack(fill="x", padx=20, pady=5)  # Adjusted padding for better spacing
                ctk.CTkLabel(
                    row_frame, 
                    text=field, 
                    font=("Helvetica", 16, "bold"),  # Bold font for summary fields
                    text_color="white"  # White color for summary fields
                ).pack(side="left", padx=10)
                ctk.CTkLabel(
                    row_frame, 
                    text=value, 
                    font=("Helvetica", 16, "bold"),  # Bold font for summary values
                    text_color="white"  # White color for summary values
                ).pack(side="right", padx=10)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

    # Quit Button for the payroll window
    quit_button = ctk.CTkButton(
        payroll_window, 
        text="Quit", 
        command=payroll_window.destroy,  # Close the payroll window
        font=("Helvetica", 16),  # Larger font for button
        fg_color="#E74C3C",  # Red color for button
        hover_color="#943126",  # Darker red on hover
        text_color="white"  # White text for button
    )
    quit_button.pack(pady=20)

    # Run the payroll window
    payroll_window.mainloop()

# Main Application Window
app = ctk.CTk()
app.title("Employee Payroll Calculator")
app.geometry("400x200")  # Smaller size for the main window

# Title Label
title_label = ctk.CTkLabel(
    app, 
    text="Employee Payroll Calculator", 
    font=("Helvetica", 24, "bold"),  # Larger and bold font for title
    text_color="white"  # White color for title
)
title_label.pack(pady=20)

# Calculate Payroll Button
calculate_button = ctk.CTkButton(
    app, 
    text="Calculate Payroll", 
    command=open_payroll_window,  # Open the payroll details window
    font=("Helvetica", 16),  # Larger font for button
    fg_color="#2E86C1",  # Blue color for button
    hover_color="#1B4F72",  # Darker blue on hover
    text_color="white"  # White text for button
)
calculate_button.pack(pady=10)

# Quit Button for the main window
quit_button = ctk.CTkButton(
    app, 
    text="Quit", 
    command=app.quit,  # Close the application
    font=("Helvetica", 16),  # Larger font for button
    fg_color="#E74C3C",  # Red color for button
    hover_color="#943126",  # Darker red on hover
    text_color="white"  # White text for button
)
quit_button.pack(pady=10)

# Run the main application
app.mainloop()