"""
Project done by Zack and Shiva.
It was interesting project to complete.
"""
# importing necessary python modules
import openpyxl
import csv

# Open the Excel file
wb = openpyxl.load_workbook("Lab4Data.xlsx")

# Select the desired worksheet
ws = wb.active

# Create a CSV file for writing
# Using group name instead of khatris2
csv_filename = "Group5.csv"

# Manually hard coding the categories
category = ["Child Labour Total", " ", "Child Labour garcons", " ", "Child Labour files", " ", "Child marriage by 15",
            " ", "Child marriage by 18",
            " ", "Birth registration total", " ", "Female genital cutting in women", " ",
            "Female genital cutting in girls ", " ",
            "Female genital support", " ", "Wife Beating Justification in male", " ",
            "Wife Beating Justification in female", " ",
            "Violent Discipline Total", " ", "Violent Discipline Male", " ", "Violent Discipline Female"]
# Opening CSV file to write
with open(csv_filename, 'w', newline='') as csv_file:
    csv_writer = csv.writer(csv_file)

    # Helps to write the header for our csv file
    csv_writer.writerow(["Name", "CategoryName", "CategoryTotal"])
    # Helps to loop through the desired section where useful data is present
    for row in ws["B15":"AE211"]:
        country_name = row[0].value
        num = 0
        # Main Logic of the program where values of category is accessed
        for index, cell in enumerate(row):
            if index > 2:
                # Accessing only the numeric values
                if isinstance(cell.value, (int, float)):
                    csv_writer.writerow([country_name, category[num], cell.value])
                num = num + 1

# Counting the number of lines rows in created CSV file
with open(csv_filename, 'r') as csv_file:
    count = sum(1 for line in csv_file)
    print(f"Number of lines in the CSV file excluding the header: {count - 1}")

# Closing the Excel work book
wb.close()
