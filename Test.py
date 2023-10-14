import openpyxl
import tkinter as tk
from tkinter import messagebox
import time

class TestGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Test")
        self.root.geometry("500x400")

        self.page = 0  # Tracks the current page (0 for name and roll, 1 for questions)

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.main_frame)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.pages = [self.create_name_roll_page, self.create_question_page]
        self.pages[self.page]()

    def create_name_roll_page(self):
        self.questions_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.questions_frame, anchor="nw")

        self.name_label = tk.Label(self.questions_frame, text="Name:")
        self.name_label.pack()
        self.name_entry = tk.Entry(self.questions_frame)
        self.name_entry.pack()

        self.roll_label = tk.Label(self.questions_frame, text="Roll No.:")
        self.roll_label.pack()
        self.roll_entry = tk.Entry(self.questions_frame)
        self.roll_entry.pack()

        self.next_button = tk.Button(self.root, text="Next", command=self.next_page)
        self.next_button.pack()

    def create_question_page(self):
        self.questions_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.questions_frame, anchor="nw")

        self.questions = {
            'Q1': {
                'Question': 'What is the capital of France?',
                'Options': ['A. Paris', 'B. London', 'C. Rome', 'D. Berlin'],
                'Answer': 'A'
            },
            'Q2': {
                'Question': 'Which planet is known as the Red Planet?',
                'Options': ['A. Mars', 'B. Venus', 'C. Jupiter', 'D. Saturn'],
                'Answer': 'A'
            },
            'Q3': {
                'Question': 'In which year the civil code was applied?',
                'Options': ['A. 1085', 'B. 1805', 'C. 1850', 'D. 1580']
            }
            # Add more questions here
        }

        self.responses = {}

        self.create_questions()

        self.timer_label = tk.Label(self.root, text="Time Left: 30:00")
        self.timer_label.pack()
        self.start_timer(60)  # 30 minutes timer

        self.submit_button = tk.Button(self.root, text="Save and Exit", command=self.save_responses)
        self.submit_button.pack()

    def create_questions(self):
        for question_id, question_data in self.questions.items():
            question_label = tk.Label(self.questions_frame, text=question_data['Question'])
            question_label.pack()

            response_var = tk.StringVar()
            response_var.set("")
            self.responses[question_id] = response_var

            for option in question_data['Options']:
                option_radio = tk.Radiobutton(self.questions_frame, text=option, variable=response_var, value=option[0])
                option_radio.pack(anchor=tk.W)

    def start_timer(self, seconds):
        self.end_time = time.time() + seconds
        self.update_timer()

    def update_timer(self):
        remaining_time = self.end_time - time.time()
        if remaining_time <= 0:
            self.save_responses()
            self.root.destroy()
        else:
            minutes = int(remaining_time // 60)
            seconds = int(remaining_time % 60)
            self.timer_label.config(text=f"Time Left: {minutes:02d}:{seconds:02d}")
            self.root.after(1000, self.update_timer)

    def next_page(self):
        if self.page == 0:
            name = self.name_entry.get().strip()
            roll_no = self.roll_entry.get().strip()

            if not name or not roll_no:
                messagebox.showerror("Error", "Please enter your name and roll number.")
                return

            self.page += 1
            self.pages[self.page]()
        elif self.page == 1:
            self.save_responses()
            self.root.destroy()

    def save_responses(self):
        name = self.name_entry.get().strip()
        roll_no = self.roll_entry.get().strip()

        if not name or not roll_no:
            return  # Don't save if name or roll number is missing

        # Create a new Excel workbook
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Set the column headers
        sheet['A1'] = 'Name'
        sheet['B1'] = 'Roll No.'
        sheet['C1'] = 'Question'
        sheet['D1'] = 'Option A'
        sheet['E1'] = 'Option B'
        sheet['F1'] = 'Option C'
        sheet['G1'] = 'Option D'
        sheet['H1'] = 'Response'

        # Save the responses in the Excel file
        row = 2
        for question_id, question_data in self.questions.items():
            response = self.responses[question_id].get()

            sheet[f'A{row}'] = name
            sheet[f'B{row}'] = roll_no
            sheet[f'C{row}'] = question_id
            sheet[f'D{row}'] = question_data['Options'][0]
            sheet[f'E{row}'] = question_data['Options'][1]
            sheet[f'F{row}'] = question_data['Options'][2]
            sheet[f'G{row}'] = question_data['Options'][3]
            sheet[f'H{row}'] = response

            row += 1

        # Save the workbook
        workbook.save('student_responses.xlsx')
        messagebox.showinfo("Success", "Test completed. Your responses have been saved in student_responses.xlsx.")
        self.root.destroy()

# Create the Tkinter root window
root = tk.Tk()

# Create an instance of the TestGUI class
test_gui = TestGUI(root)

# Run the Tkinter event loop
root.mainloop()
