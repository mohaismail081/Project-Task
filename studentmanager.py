import pandas as pd
import os

# --- Configuration ---
FILE_NAME = 'students.xlsx'
SHEET_NAME = 'Roster'
COLUMNS = ['roll_no', 'name', 'marks']
# ---------------------

class StudentManager:
    """
    Manages student data, reading from and writing to an Excel file 
    using pandas. Encapsulates all CRUD and reporting logic.
    """
    def __init__(self, file_name, sheet_name):
        """Initializes the manager, loads data, and sets configurations."""
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.students_df = self._load_data()

    # --- Utility/Private Methods ---

    def _load_data(self):
        """Loads student data from the Excel file."""
        if os.path.exists(self.file_name):
            try:
                df = pd.read_excel(self.file_name, sheet_name=self.sheet_name)
                df['roll_no'] = df['roll_no'].fillna(-1).astype(int)
                return df
            except Exception as e:
                print(f"⚠️ Error reading file {self.file_name}: {e}. Starting with an empty roster.")

        # Create a new empty DataFrame if loading failed or file doesn't exist
        return pd.DataFrame(columns=COLUMNS)

    def _save_data(self):
        """Saves the current DataFrame back to the Excel file."""
        try:
            with pd.ExcelWriter(self.file_name, engine='xlsxwriter') as writer:
                self.students_df.to_excel(writer, sheet_name=self.sheet_name, index=False)
            print(f"Data saved successfully to {self.file_name}.")
        except Exception as e:
            print(f"❌ Error saving data to Excel: {e}")

    def _get_integer_input(self, prompt, min_val=None, max_val=None):
        """Handles integer input validation."""
        while True:
            try:
                value = int(input(prompt))
                if min_val is not None and value < min_val:
                    print(f"❌ Value must be at least {min_val}.")
                    continue
                if max_val is not None and value > max_val:
                    print(f"❌ Value cannot exceed {max_val}.")
                    continue
                return value
            except ValueError:
                print("❌ Invalid input. Please enter a whole number.")

    def _apply_class_label(self, df):
        """Adds a 'class' column to a DataFrame copy based on marks."""
        if df.empty:
            return df
            
        bins = [0, 35, 50, 60, 75, 101] 
        labels = ['Fail', 'Third Class', 'Second Class', 'First Class', 'Distinction']
        
        df['class'] = pd.cut(df['marks'], bins=bins, labels=labels, right=False)
        df['class'] = df['class'].astype(str).replace('nan', 'Fail')
        
        return df

    # --- Public CRUD & Reporting Methods ---

    def add_student(self):
        """Adds a new student record to the DataFrame with validation."""
        print("\n--- Add New Student ---")
        name = input("Enter student name: ").strip()
        if not name:
            print("❌ Student name cannot be empty.")
            return

        while True:
            roll_no = self._get_integer_input("Enter roll number: ")
            if roll_no in self.students_df['roll_no'].values:
                print(f"❌ Error: Roll number {roll_no} already exists.")
            else:
                break

        marks = self._get_integer_input("Enter marks (0-100): ", min_val=0, max_val=100)

        new_row = pd.DataFrame([{'roll_no': roll_no, 'name': name, 'marks': marks}])
        self.students_df = pd.concat([self.students_df, new_row], ignore_index=True)
        
        self._save_data()
        print("✅ Student added successfully.")

    def view_students(self):
        """Prints all student records from the DataFrame."""
        if self.students_df.empty:
            print("The student roster is currently empty.")
            return
        
        display_df = self._apply_class_label(self.students_df.copy()) 
        display_df = display_df.sort_values(by='roll_no').reset_index(drop=True)
        
        print("\n--- Student Roster ---")
        print(display_df.to_string(index=False))
        print("----------------------")


    def search_student(self):
        """Searches for a student by roll number."""
        roll_no = self._get_integer_input("Enter roll number to search: ")
        
        result = self.students_df[self.students_df['roll_no'] == roll_no]
        
        if not result.empty:
            result_with_class = self._apply_class_label(result.copy())
            print(f"\n✅ Student Found (Roll No: {roll_no}):")
            print(result_with_class.to_string(index=False, header=True))
        else:
            print(f"Student with roll number {roll_no} not found.")

    def update_student(self):
        """Updates the name or marks for an existing student."""
        roll_no = self._get_integer_input("Enter roll number of student to update: ")
        
        index_to_update = self.students_df[self.students_df['roll_no'] == roll_no].index
        
        if index_to_update.empty:
            print(f"❌ Student with roll number {roll_no} not found.")
            return

        print("\nCurrent Record:")
        print(self.students_df.loc[index_to_update].to_string(index=False))

        print("\nWhat do you want to update?")
        print("1. Name\n2. Marks")
        choice = input("Enter your choice (1 or 2): ")

        if choice == '1':
            new_name = input("Enter new name: ").strip()
            if new_name:
                self.students_df.loc[index_to_update, 'name'] = new_name
                print("✅ Name updated.")
            else:
                print("❌ Name update cancelled (Name cannot be empty).")
                return
        elif choice == '2':
            new_marks = self._get_integer_input("Enter new marks (0-100): ", min_val=0, max_val=100)
            self.students_df.loc[index_to_update, 'marks'] = new_marks
            print("✅ Marks updated.")
        else:
            print("❌ Invalid choice. Update cancelled.")
            return

        self._save_data()

    def delete_student(self):
        """Deletes a student record by roll number."""
        roll_no_to_delete = self._get_integer_input("Enter roll number to delete: ")
        
        if roll_no_to_delete not in self.students_df['roll_no'].values:
            print(f"Student with roll number {roll_no_to_delete} not found.")
            return

        self.students_df = self.students_df[self.students_df['roll_no'] != roll_no_to_delete]
        
        self._save_data()
        print(f"✅ Student with roll number {roll_no_to_delete} deleted.")

    def generate_report(self):
        """Calculates, classifies, and displays basic statistics on student marks."""
        if self.students_df.empty:
            print("Cannot generate report: The student roster is empty.")
            return

        # Use the utility to classify the data for the report
        report_df = self._apply_class_label(self.students_df.copy()) 
        
        marks = report_df['marks']
        
        print("\n--- Student Performance Report ---")
        
        # General Statistics
        print(f"Total Students: {len(report_df)}")
        print(f"Average Marks: {marks.mean():.2f}")
        print(f"Highest Marks: {marks.max()}")
        print(f"Lowest Marks: {marks.min()}")

        # Classification Summary
        print("\n--- Class Summary ---")
        class_counts = report_df['class'].value_counts().sort_index()
        print(class_counts.to_string())

        # Top Performer(s)
        if not marks.empty:
            highest_mark = marks.max()
            top_students = report_df[report_df['marks'] == highest_mark]
            print("\nTop Performer(s):")
            print(top_students[['roll_no', 'name', 'marks', 'class']].to_string(index=False))
        
        print("----------------------------------")

# ------------------- Main Menu Logic (Stays Outside the Class) -------------------

def menu(manager):
    """Main application loop, interacting with the StudentManager object."""
    print("Welcome to the OOP-Enhanced Student Management System!")
    while True:
        print("\n--- Menu ---")
        print("1. Add Student (Create)")
        print("2. View All (Read)")
        print("3. Search")
        print("4. Update Student")
        print("5. Delete Student")
        print("6. Generate Report")
        print("7. Exit")
        choice = input("Enter your choice: ")
        
        if choice == '1':
            manager.add_student()
        elif choice == '2':
            manager.view_students()
        elif choice == '3':
            manager.search_student()
        elif choice == '4':
            manager.update_student()
        elif choice == '5':
            manager.delete_student()
        elif choice == '6':
            manager.generate_report()
        elif choice == '7':
            print("Exiting application. Goodbye! 👋")
            break
        else:
            print("Invalid choice. Please select a number from the menu.")

# --- Execution ---
if __name__ == "__main__":
    # Create the manager object once. This calls the __init__ and loads the data.
    student_manager = StudentManager(FILE_NAME, SHEET_NAME)
    
    # Start the application menu, passing the object
    menu(student_manager)