import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import re
import os
from difflib import Differ
import pandas as pd
from tkinter.scrolledtext import ScrolledText
import xlsxwriter

class JavaAssessmentGrader:
    def __init__(self, root):
        self.root = root
        self.root.title("Java Practical Assessment Grader")
        self.root.geometry("1400x1000")  # Increased height
        
        # Variables
        self.marking_scheme_path = tk.StringVar()
        self.student_submission_path = tk.StringVar()
        self.student_name = tk.StringVar()
        self.total_marks = tk.DoubleVar(value=0.0)
        self.achieved_marks = tk.DoubleVar(value=0.0)
        self.current_selection = {"scheme": "", "submission": ""}
        self.current_selection_pos = {"scheme": {"start": "", "end": ""}, "submission": {"start": "", "end": ""}}
        
        # Clipboard storage
        self.clipboard_content = ""

        # Create UI with adjusted proportions
        self.create_widgets()

    def copy_selection(self):
        """Copy selected student code."""
        try:
            sel_start = self.student_submission_text.index(tk.SEL_FIRST)
            sel_end = self.student_submission_text.index(tk.SEL_LAST)
            self.clipboard_content = self.student_submission_text.get(sel_start, sel_end)
            self.root.clipboard_clear()
            self.root.clipboard_append(self.clipboard_content)
        except tk.TclError:
            pass  # No selection made

    def paste_into_entry(self):
        """Paste copied reference code and prompt user for awarded marks."""
        selected = self.results_tree.selection()
        if not selected or not self.clipboard_content:
            return

        item = selected[0]
        values = self.results_tree.item(item, 'values')

        if values[5] == "not_found":  # Ensure we're pasting into a "Not Found" entry
            allocated_marks = float(values[1])  # Convert allocated marks to float

            # Ask user for awarded marks
            awarded_marks = simpledialog.askfloat(
                "Assign Marks",
                f"Enter awarded marks (Max {allocated_marks}):",
                minvalue=0.0,
                maxvalue=allocated_marks
            )

            if awarded_marks is None:  # User canceled input
                return

            # Update the grading table entry
            self.results_tree.item(item, values=(
                values[0],  # Assessment Criteria
                allocated_marks,  # Allocated Marks
                awarded_marks,  # User-entered Awarded Marks
                "Reference code manually added",
                self.clipboard_content,  # The pasted reference code
                "found"  # Update status
            ))

            # Remove "Not Found" highlighting
            self.results_tree.item(item, tags=())  # Clear row highlight
            self.marking_scheme_text.tag_remove('not_found', '1.0', tk.END)

            # Clear clipboard after pasting
            self.clipboard_content = ""

            # Notify user of successful update
            messagebox.showinfo("Success", "Reference code has been pasted. Marks updated.")

    def create_widgets(self):
        # Main container using grid for better control
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # ========== File Selection Section ==========
        file_frame = ttk.LabelFrame(main_container, text="File Selection", padding="10")
        file_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        # Marking scheme selection
        ttk.Label(file_frame, text="Marking Scheme:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(file_frame, textvariable=self.marking_scheme_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_marking_scheme, width=10).grid(row=0, column=2, padx=5)
        
        # Student submission selection
        ttk.Label(file_frame, text="Student Submission:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Entry(file_frame, textvariable=self.student_submission_path, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_student_submission, width=10).grid(row=1, column=2, padx=5)
        
        # Student name and load button
        ttk.Label(file_frame, text="Student Name:").grid(row=2, column=0, sticky="w", padx=5)
        ttk.Entry(file_frame, textvariable=self.student_name, width=20).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Button(file_frame, text="Load Files", command=self.load_files, width=15).grid(row=2, column=2, padx=5)
        
        # ========== Code Comparison Section ==========
        code_frame = ttk.Frame(main_container)
        code_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        code_frame.grid_rowconfigure(0, weight=1)
        code_frame.grid_columnconfigure(0, weight=1)
        code_frame.grid_columnconfigure(1, weight=1)
        
        # Left pane - Marking scheme
        scheme_frame = ttk.LabelFrame(code_frame, text="Marking Scheme", padding="5")
        scheme_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        
        scheme_container = ttk.Frame(scheme_frame)
        scheme_container.pack(fill=tk.BOTH, expand=True)
        
        scheme_vscroll = ttk.Scrollbar(scheme_container, orient=tk.VERTICAL)
        scheme_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        scheme_hscroll = ttk.Scrollbar(scheme_container, orient=tk.HORIZONTAL)
        scheme_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.marking_scheme_text = tk.Text(
            scheme_container, wrap=tk.NONE, font=('Courier', 10),
            yscrollcommand=scheme_vscroll.set,
            xscrollcommand=scheme_hscroll.set,
            width=60, height=25)
        self.marking_scheme_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scheme_vscroll.config(command=self.marking_scheme_text.yview)
        scheme_hscroll.config(command=self.marking_scheme_text.xview)
        self.marking_scheme_text.bind("<<Selection>>", lambda e: self.on_text_select(e, "scheme"))
        
        # Right pane - Student submission
        submission_frame = ttk.LabelFrame(code_frame, text="Student Submission", padding="5")
        submission_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        
        submission_container = ttk.Frame(submission_frame)
        submission_container.pack(fill=tk.BOTH, expand=True)
        
        submission_vscroll = ttk.Scrollbar(submission_container, orient=tk.VERTICAL)
        submission_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        submission_hscroll = ttk.Scrollbar(submission_container, orient=tk.HORIZONTAL)
        submission_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.student_submission_text = tk.Text(
            submission_container, wrap=tk.NONE, font=('Courier', 10),
            yscrollcommand=submission_vscroll.set,
            xscrollcommand=submission_hscroll.set,
            width=60, height=25)
        self.student_submission_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        submission_vscroll.config(command=self.student_submission_text.yview)
        submission_hscroll.config(command=self.student_submission_text.xview)
        self.student_submission_text.bind("<<Selection>>", lambda e: self.on_text_select(e, "submission"))
        
        # ========== Action Buttons ==========
        action_frame = ttk.Frame(main_container)
        action_frame.grid(row=2, column=0, sticky="ew", pady=5)
        
        # Left side buttons
        left_btn_frame = ttk.Frame(action_frame)
        left_btn_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(left_btn_frame, text="Compare", command=self.compare_selections, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(left_btn_frame, text="Grade", command=self.grade_selection, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(left_btn_frame, text="Clear", command=self.clear_selections, width=12).pack(side=tk.LEFT, padx=2)
        
        # Copy/Paste buttons
        copy_paste_frame = ttk.Frame(action_frame)
        copy_paste_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.copy_btn = ttk.Button(copy_paste_frame, text="Copy Student Code", command=self.copy_selection, width=16)
        self.copy_btn.pack(side=tk.LEFT, padx=2)
        self.paste_btn = ttk.Button(copy_paste_frame, text="Paste Reference Code", command=self.paste_into_entry, width=16)
        self.paste_btn.pack(side=tk.LEFT, padx=2)
        
        # Right side buttons
        right_btn_frame = ttk.Frame(action_frame)
        right_btn_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        ttk.Button(right_btn_frame, text="Calculate", command=self.calculate_marks, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(right_btn_frame, text="Save TXT", command=self.save_results_txt, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(right_btn_frame, text="Save Excel", command=self.save_results_excel, width=12).pack(side=tk.LEFT, padx=2)
        
        # Marks display
        marks_frame = ttk.Frame(main_container)
        marks_frame.grid(row=3, column=0, sticky="ew", pady=5)
        
        ttk.Label(marks_frame, text="Total Marks:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Label(marks_frame, textvariable=self.total_marks, font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        ttk.Label(marks_frame, text="Achieved Marks:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Label(marks_frame, textvariable=self.achieved_marks, font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        
        # ========== Results Table ==========
        results_frame = ttk.LabelFrame(main_container, text="Grading Results", padding="5")
        results_frame.grid(row=4, column=0, sticky="nsew", pady=5)
        
        # Treeview with scrollbars
        tree_container = ttk.Frame(results_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        tree_vscroll = ttk.Scrollbar(tree_container)
        tree_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree_hscroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        tree_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Configure results tree
        self.results_tree = ttk.Treeview(
            tree_container, 
            columns=('criteria', 'allocated', 'awarded', 'comments', 'reference', 'status'), 
            show='headings', 
            height=12,
            yscrollcommand=tree_vscroll.set,
            xscrollcommand=tree_hscroll.set
        )
        
        # Configure columns
        self.results_tree.heading('criteria', text='Assessment Criteria')
        self.results_tree.column('criteria', width=400, stretch=tk.YES)
        
        self.results_tree.heading('allocated', text='Allocated Marks')
        self.results_tree.column('allocated', width=100, anchor=tk.CENTER)
        
        self.results_tree.heading('awarded', text='Awarded Marks')
        self.results_tree.column('awarded', width=100, anchor=tk.CENTER)
        
        self.results_tree.heading('comments', text='Comments')
        self.results_tree.column('comments', width=300, stretch=tk.YES)
        
        self.results_tree.heading('reference', text='Reference Code')
        self.results_tree.column('reference', width=200, stretch=tk.YES)
        
        # Hidden status column
        self.results_tree.heading('status', text='Status')
        self.results_tree.column('status', width=0, stretch=tk.NO)
        
        self.results_tree.pack(fill=tk.BOTH, expand=True)
        tree_vscroll.config(command=self.results_tree.yview)
        tree_hscroll.config(command=self.results_tree.xview)
        
        # Assign marks button
        assign_frame = ttk.Frame(results_frame)
        assign_frame.pack(fill=tk.X, pady=5)
        self.assign_btn = ttk.Button(assign_frame, text="Assign Marks to Selected", 
                                command=self.assign_marks_to_selected, state=tk.DISABLED)
        self.assign_btn.pack(side=tk.LEFT, padx=5)
        
        # Configure tags and bindings
        self.configure_tags_and_bindings()

    def configure_tags_and_bindings(self):
        """Configure text tags and event bindings for the GUI"""
        # Right-click menu for treeview
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="Edit Awarded Marks", command=self.edit_awarded_marks)
        self.tree_menu.add_command(label="Edit Comments", command=self.edit_comments)
        self.tree_menu.add_command(label="View Reference Code", command=self.view_reference_code)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Delete Row", command=self.delete_selected_row)
        self.results_tree.bind("<Button-3>", self.show_tree_menu)
        
        # Double-click editing
        self.results_tree.bind("<Double-1>", self.on_tree_double_click)
        self.results_tree.bind("<Button-1>", self.on_result_click)
        
        # Configure text colors
        self.marking_scheme_text.tag_configure('mark', background='yellow')
        self.marking_scheme_text.tag_configure('selected', background='lightblue')
        self.marking_scheme_text.tag_configure('not_found', background='orange')
        
        self.student_submission_text.tag_configure('selected', background='lightblue')
        self.student_submission_text.tag_configure('match', background='lightgreen')
        self.student_submission_text.tag_configure('mismatch', background='pink')
        self.student_submission_text.tag_configure('missing', background='orange')
        self.student_submission_text.tag_configure('graded', background='#a0e0a0')
        self.student_submission_text.tag_configure('search', background='yellow')
        
        self.results_tree.tag_configure('not_found', background='#ffdddd')  # Light red
            
    def on_text_select(self, event, source):
        """Handle text selection in either marking scheme or student submission"""
        try:
            widget = self.marking_scheme_text if source == "scheme" else self.student_submission_text
            
            # Get the selected text and positions
            sel_start = widget.index(tk.SEL_FIRST)
            sel_end = widget.index(tk.SEL_LAST)
            selected_text = widget.get(sel_start, sel_end)
            
            # Store the selection
            self.current_selection[source] = selected_text
            self.current_selection_pos[source] = {"start": sel_start, "end": sel_end}
            
            # Highlight the selection
            widget.tag_remove('selected', '1.0', tk.END)
            widget.tag_add('selected', sel_start, sel_end)
            
        except tk.TclError:
            # No text currently selected
            self.current_selection[source] = ""
            self.current_selection_pos[source] = {"start": "", "end": ""}
    
    def clear_selections(self):
        """Clear all current selections"""
        self.marking_scheme_text.tag_remove('selected', '1.0', tk.END)
        self.student_submission_text.tag_remove('selected', '1.0', tk.END)
        self.current_selection = {"scheme": "", "submission": ""}
        self.current_selection_pos = {"scheme": {"start": "", "end": ""}, "submission": {"start": "", "end": ""}}
    
    def compare_selections(self):
        """Compare the selected code from both panels"""
        scheme_sel = self.current_selection["scheme"]
        submission_sel = self.current_selection["submission"]
        
        if not scheme_sel and not submission_sel:
            messagebox.showwarning("No Selections", "Please select code from both panels to compare")
            return
        
        # Create comparison dialog
        compare_dialog = tk.Toplevel(self.root)
        compare_dialog.title("Code Comparison")
        compare_dialog.geometry("800x600")
        
        # Frame for comparison
        compare_frame = ttk.Frame(compare_dialog, padding="10")
        compare_frame.pack(fill=tk.BOTH, expand=True)
        
        # Marking scheme selection
        ttk.Label(compare_frame, text="Marking Scheme Reference:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        scheme_text = tk.Text(compare_frame, wrap=tk.WORD, height=10, font=('Courier', 10))
        scheme_text.pack(fill=tk.X, pady=5)
        scheme_text.insert(tk.END, scheme_sel)
        scheme_text.config(state=tk.DISABLED)
        
        # Student submission selection
        ttk.Label(compare_frame, text="Student Submission:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        submission_text = tk.Text(compare_frame, wrap=tk.WORD, height=10, font=('Courier', 10))
        submission_text.pack(fill=tk.X, pady=5)
        submission_text.insert(tk.END, submission_sel)
        submission_text.config(state=tk.DISABLED)
        
        # Comparison notes
        ttk.Label(compare_frame, text="Comparison Notes:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        notes_entry = tk.Text(compare_frame, wrap=tk.WORD, height=5, font=('Arial', 10))
        notes_entry.pack(fill=tk.X, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(compare_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Close", command=compare_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Center the dialog
        compare_dialog.transient(self.root)
        compare_dialog.grab_set()
        self.root.wait_window(compare_dialog)
    
    def grade_selection(self):
        """Grade the currently selected student code with optional reference to marking scheme"""
        submission_sel = self.current_selection["submission"]
        scheme_sel = self.current_selection["scheme"]
        
        if not submission_sel:
            messagebox.showwarning("No Selection", "Please select code from the student submission first")
            return
        
        # Create grading dialog
        grade_dialog = tk.Toplevel(self.root)
        grade_dialog.title("Grade Code Selection")
        grade_dialog.geometry("600x400")
        
        # Make dialog modal
        grade_dialog.grab_set()
        
        # Frame for grading inputs
        grade_frame = ttk.Frame(grade_dialog, padding="10")
        grade_frame.pack(fill=tk.BOTH, expand=True)
        
        # Selected code display
        ttk.Label(grade_frame, text="Selected Student Code:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        code_text = tk.Text(grade_frame, wrap=tk.WORD, height=5, font=('Courier', 10))
        code_text.pack(fill=tk.X, pady=5)
        code_text.insert(tk.END, submission_sel)
        code_text.config(state=tk.DISABLED)
        
        # Reference code (if any)
        if scheme_sel:
            ttk.Label(grade_frame, text="Reference Code from Marking Scheme:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
            ref_text = tk.Text(grade_frame, wrap=tk.WORD, height=5, font=('Courier', 10))
            ref_text.pack(fill=tk.X, pady=5)
            ref_text.insert(tk.END, scheme_sel)
            ref_text.config(state=tk.DISABLED)
        
        # Grading inputs
        input_frame = ttk.Frame(grade_frame)
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(input_frame, text="Allocated Marks:").grid(row=0, column=0, padx=5, sticky=tk.W)
        allocated_spin = ttk.Spinbox(input_frame, from_=0, to=100, increment=0.5)
        allocated_spin.grid(row=0, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="Awarded Marks:").grid(row=1, column=0, padx=5, sticky=tk.W)
        awarded_spin = ttk.Spinbox(input_frame, from_=0, to=100, increment=0.5)
        awarded_spin.grid(row=1, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="Comments:").grid(row=2, column=0, padx=5, sticky=tk.W)
        comments_entry = ttk.Entry(input_frame, width=50)
        comments_entry.grid(row=2, column=1, padx=5, sticky=tk.W)
        
        # Buttons frame
        button_frame = ttk.Frame(grade_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def submit_grade():
            try:
                allocated = float(allocated_spin.get())
                awarded = float(awarded_spin.get())
                comments = comments_entry.get()
                
                if awarded > allocated:
                    messagebox.showerror("Invalid Marks", "Awarded marks cannot exceed allocated marks")
                    return
                
                # Add to results tree
                self.results_tree.insert('', tk.END, values=(
                    f"Manual: {submission_sel[:50]}..." if len(submission_sel) > 50 else f"Manual: {submission_sel}",
                    allocated,
                    awarded,
                    comments,
                    scheme_sel[:50] + "..." if scheme_sel and len(scheme_sel) > 50 else scheme_sel
                ))

                ## Remove the `not_found` highlight in the marking scheme when grading occurs
                sel_start = self.current_selection_pos["scheme"]["start"]
                sel_end = self.current_selection_pos["scheme"]["end"]
                self.marking_scheme_text.tag_remove('not_found', sel_start, sel_end)

                # Mark the code as graded and remove not-found highlighting if present
                sel_start = self.current_selection_pos["submission"]["start"]
                sel_end = self.current_selection_pos["submission"]["end"]
                self.student_submission_text.tag_add('graded', sel_start, sel_end)
                self.student_submission_text.tag_remove('not_found', sel_start, sel_end)
                
                # Update only achieved marks (not total marks)
                current_achieved = self.achieved_marks.get()
                self.achieved_marks.set(current_achieved + awarded)
                
                # Clear selections
                self.clear_selections()
                
                # Close dialog
                grade_dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {str(e)}")
        
        # Submit button - now properly bound to submit_grade function
        submit_btn = ttk.Button(button_frame, text="Submit Grade", command=submit_grade)
        submit_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=grade_dialog.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Make sure dialog is properly centered
        grade_dialog.transient(self.root)
        grade_dialog.wait_window()
    
    def view_reference_code(self):
        """View the full reference code for selected item"""
        selected = self.results_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.results_tree.item(item, 'values')
        reference = values[4] if len(values) > 4 else ""
        
        if not reference:
            messagebox.showinfo("No Reference", "No reference code available for this item")
            return
        
        # Create view dialog
        view_dialog = tk.Toplevel(self.root)
        view_dialog.title("Reference Code")
        view_dialog.geometry("600x400")
        
        # Frame for reference code
        ref_frame = ttk.Frame(view_dialog, padding="10")
        ref_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(ref_frame, text="Reference Code from Marking Scheme:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        ref_text = tk.Text(ref_frame, wrap=tk.WORD, font=('Courier', 10))
        ref_text.pack(fill=tk.BOTH, expand=True, pady=5)
        ref_text.insert(tk.END, reference)
        ref_text.config(state=tk.DISABLED)
        
        # Close button
        ttk.Button(ref_frame, text="Close", command=view_dialog.destroy).pack(anchor=tk.E)
        
        # Center the dialog
        view_dialog.transient(self.root)
        view_dialog.grab_set()
        self.root.wait_window(view_dialog)
    
    def show_tree_menu(self, event):
        """Show right-click menu for treeview items"""
        item = self.results_tree.identify_row(event.y)
        if item:
            self.results_tree.selection_set(item)
            self.tree_menu.post(event.x_root, event.y_root)
    
    def edit_awarded_marks(self):
        selected = self.results_tree.selection()
        if not selected:
            return

        item = selected[0]
        current_values = self.results_tree.item(item, 'values')
        allocated = float(current_values[1])

        awarded = simpledialog.askfloat(
            "Edit Awarded Marks",
            f"Enter awarded marks (allocated: {allocated}):",
            parent=self.root,
            minvalue=0.0,
            maxvalue=allocated,
            initialvalue=float(current_values[2])
        )

        if awarded is not None:
            # Update the treeview with all 6 values
            self.results_tree.item(item, values=(
                current_values[0],
                current_values[1],
                awarded,
                current_values[3],
                current_values[4] if len(current_values) > 4 else "",
                current_values[5] if len(current_values) > 5 else "found"
            ))

        # Ensure achieved marks are recalculated
        self.update_achieved_marks()

    def edit_comments(self):
        """Edit comments for selected criteria"""
        selected = self.results_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        current_values = self.results_tree.item(item, 'values')
        
        comment = simpledialog.askstring(
            "Edit Comments",
            "Enter new comments:",
            parent=self.root,
            initialvalue=current_values[3]
        )
        
        if comment is not None:
            # Update the treeview with all 6 values
            self.results_tree.item(item, values=(
                current_values[0],
                current_values[1],
                current_values[2],
                comment,
                current_values[4] if len(current_values) > 4 else "",
                current_values[5] if len(current_values) > 5 else "found"
            ))
    
    def update_achieved_marks(self):
        """Recalculate achieved marks from treeview"""
        total = 0.0
        for item in self.results_tree.get_children():
            awarded = float(self.results_tree.item(item, 'values')[2])
            total += awarded
        self.achieved_marks.set(total)

    def delete_selected_row(self):
        """Delete the selected row from the treeview"""
        selected = self.results_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.results_tree.item(item, 'values')
        
        # Ask for confirmation
        if messagebox.askyesno("Confirm Delete", "Delete this grading entry?"):
            # If this was a manually graded item, we might need to remove highlighting
            if values[0].startswith("Manual:"):
                # Find the reference in the student submission and remove 'graded' tag
                student_code = values[0][7:].strip()  # Remove "Manual: " prefix
                self.remove_graded_highlight(student_code)
            
            self.results_tree.delete(item)
            self.update_achieved_marks()
    
    def remove_graded_highlight(self, code_snippet):
        """Remove graded highlighting from student submission text"""
        if not code_snippet:
            return
            
        # Get all text
        full_text = self.student_submission_text.get("1.0", tk.END)
        
        # Find the code snippet in the text
        start_pos = full_text.find(code_snippet)
        if start_pos == -1:
            return
            
        # Convert to line.column format
        start_line = full_text.count('\n', 0, start_pos) + 1
        line_start = full_text.rfind('\n', 0, start_pos) + 1
        start_col = start_pos - line_start
        
        end_pos = start_pos + len(code_snippet)
        end_line = full_text.count('\n', 0, end_pos) + 1
        line_end = full_text.rfind('\n', 0, end_pos) + 1
        end_col = end_pos - line_end
        
        # Remove the 'graded' tag
        self.student_submission_text.tag_remove('graded', 
                                              f"{start_line}.{start_col}", 
                                              f"{end_line}.{end_col}")
    
    def on_tree_double_click(self, event):
        """Handle double-click events on treeview for in-line editing"""
        region = self.results_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.results_tree.identify_column(event.x)
            item = self.results_tree.identify_row(event.y)
            
            if item and column in ("#2", "#3", "#4"):  # Only allow editing allocated, awarded, comments
                self.edit_tree_cell(item, column)
    
    def edit_tree_cell(self, item, column):
        """Edit a specific cell in the treeview"""
        # Get current values
        values = list(self.results_tree.item(item, 'values'))
        col_index = int(column[1:]) - 1  # Convert #2 to 1, etc.
        current_value = values[col_index]
        
        # Get column bounding box
        x, y, width, height = self.results_tree.bbox(item, column)
        
        # Create entry widget for editing
        entry = ttk.Entry(self.results_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, str(current_value))
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit(event=None):
            """Save the edited value"""
            new_value = entry.get()
            
            # Validate input for numeric columns
            if column in ("#2", "#3"):  # Allocated or Awarded marks
                try:
                    new_value = float(new_value)
                    if column == "#3":  # Awarded marks
                        allocated = float(values[1])  # Get allocated marks from the row
                        if new_value > allocated:
                            messagebox.showerror("Error", "Awarded marks cannot exceed allocated marks")
                            entry.destroy()
                            return
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid number")
                    entry.destroy()
                    return
            
            # Update the treeview
            values[col_index] = new_value
            self.results_tree.item(item, values=values)
            
            # Update achieved marks if awarded marks were changed
            if column == "#3":
                self.update_achieved_marks()
            
            entry.destroy()
        
        def cancel_edit(event=None):
            """Cancel editing"""
            entry.destroy()
        
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)
    
    def browse_marking_scheme(self):
        filepath = filedialog.askopenfilename(
            title="Select Marking Scheme Java File",
            filetypes=(("Java files", "*.java"), ("All files", "*.*"))
        )
        if filepath:
            self.marking_scheme_path.set(filepath)
    
    def browse_student_submission(self):
        filepath = filedialog.askopenfilename(
            title="Select Student Submission Java File",
            filetypes=(("Java files", "*.java"), ("All files", "*.*"))
        )
        if filepath:
            self.student_submission_path.set(filepath)
            # Extract student name from filename as a default
            filename = os.path.basename(filepath)
            self.student_name.set(os.path.splitext(filename)[0])
    
    def load_files(self):
        if not self.marking_scheme_path.get() or not self.student_submission_path.get():
            messagebox.showerror("Error", "Please select both marking scheme and student submission files")
            return
        
        try:
            # Load marking scheme
            with open(self.marking_scheme_path.get(), 'r') as f:
                marking_scheme_content = f.read()
                self.marking_scheme_text.delete(1.0, tk.END)
                self.marking_scheme_text.insert(tk.END, marking_scheme_content)
                self.highlight_marks_in_scheme()
            
            # Load student submission
            with open(self.student_submission_path.get(), 'r') as f:
                student_content = f.read()
                self.student_submission_text.delete(1.0, tk.END)
                self.student_submission_text.insert(tk.END, student_content)
            
            # Parse marking scheme to get total marks
            self.parse_marking_scheme()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load files: {str(e)}")
    
    def highlight_marks_in_scheme(self):
        """Highlight mark allocations in the marking scheme"""
        text = self.marking_scheme_text.get(1.0, tk.END)
        self.marking_scheme_text.tag_remove('mark', 1.0, tk.END)
        
        # Find all marks in comments (format: // 1.0 or /* 1.0 */)
        pattern = r'(//|/\*)\s*(\d+\.?\d*)\s*(?:\*/)?'
        for match in re.finditer(pattern, text):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            self.marking_scheme_text.tag_add('mark', start, end)
    
    def parse_marking_scheme(self):
        """Parse the marking scheme to extract criteria and allocated marks"""
        text = self.marking_scheme_text.get(1.0, tk.END)
        
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Find all marks in comments and their context
        pattern = r'(.*?)(//|/\*)\s*(\d+\.?\d*)\s*(?:\*/)?'
        total = 0.0
        for match in re.finditer(pattern, text):
            criteria = match.group(1).strip()
            mark = float(match.group(3))
            total += mark
            
            # Add to treeview
            self.results_tree.insert('', tk.END, values=(criteria, mark, 0.0, ""))
        
        self.total_marks.set(total)
    
    def normalize_whitespace(self, code):
        """Normalize whitespace in code for comparison"""
        # Remove extra spaces, keep single spaces
        code = re.sub(r'\s+', ' ', code)
        # Remove spaces around special characters
        code = re.sub(r'\s*([{}();,=+\-*/])\s*', r'\1', code)
        return code.strip()
    
    def calculate_marks(self):
        """Compare student submission with marking scheme and calculate marks"""
        if not self.marking_scheme_path.get() or not self.student_submission_path.get():
            messagebox.showerror("Error", "Please load both files first")
            return
        
        scheme_text = self.marking_scheme_text.get(1.0, tk.END)
        student_text = self.student_submission_text.get(1.0, tk.END)
        
        # Clear previous highlighting
        self.student_submission_text.tag_remove('match', 1.0, tk.END)
        self.student_submission_text.tag_remove('mismatch', 1.0, tk.END)
        self.student_submission_text.tag_remove('missing', 1.0, tk.END)
        self.marking_scheme_text.tag_remove('not_found', 1.0, tk.END)
        
        achieved = 0.0
        
        # First sum up any manually awarded marks
        for item in self.results_tree.get_children():
            values = self.results_tree.item(item, 'values')
            if values[0].startswith("Manual:"):  # This is a manually graded item
                achieved += float(values[2])
        
        # Then add marks for automatically matched criteria
        for item in self.results_tree.get_children():
            values = self.results_tree.item(item, 'values')
            criteria = values[0]
            
            # Skip manually added items
            if criteria.startswith("Manual:"):
                continue
                
            allocated = float(values[1])
            
            # Normalize both the criteria and student code for comparison
            norm_criteria = self.normalize_whitespace(criteria)
            
            # Search for the normalized criteria in student code
            student_code = self.normalize_whitespace(student_text)
            
            if norm_criteria and norm_criteria in student_code:
                awarded = allocated
                achieved += awarded
                # Update with all 6 values (including empty reference and status)
                self.results_tree.item(item, values=(
                    criteria,
                    allocated,
                    awarded,
                    "Found in submission",
                    "",  # Empty reference
                    "found"  # Status
                ), tags=())
                
                # Highlight matching lines in student submission
                self.highlight_matching_lines(criteria)
            else:
                # Update with all 6 values
                self.results_tree.item(item, values=(
                    criteria,
                    allocated,
                    0.0,
                    "Not found in submission",
                    "",  # Empty reference
                    "not_found"  # Status
                ), tags=('not_found',))
                # Highlight not-found code in marking scheme
                self.highlight_not_found_code(criteria)
        
        self.achieved_marks.set(achieved)
        self.update_table_highlights()
        
    def update_table_highlights(self):
        """Update highlighting for not found items"""
        for item in self.results_tree.get_children():
            values = self.results_tree.item(item, 'values')
            if len(values) > 5 and values[5] == "not_found":
                self.results_tree.item(item, tags=('not_found',))
            else:
                self.results_tree.item(item, tags=())
    
    def on_result_click(self, event):
        """Handle clicks on results table"""
        item = self.results_tree.identify_row(event.y)
        if item:
            values = self.results_tree.item(item, 'values')
            if len(values) > 5 and values[5] == "not_found":
                # Enable assign button for not found items
                self.assign_btn.config(state=tk.NORMAL)
                self.current_not_found_item = item
                # Highlight corresponding code in submission
                self.highlight_criteria_in_submission(values[0])
            else:
                self.assign_btn.config(state=tk.DISABLED)
                self.current_not_found_item = None
    
    def highlight_criteria_in_submission(self, criteria):
        """Highlight where criteria might exist in submission"""
        self.student_submission_text.tag_remove('search', '1.0', tk.END)
        
        norm_criteria = self.normalize_whitespace(criteria)
        student_text = self.student_submission_text.get('1.0', tk.END)
        
        # Find approximate matches
        start_idx = '1.0'
        while True:
            start_idx = self.student_submission_text.search(
                norm_criteria, 
                start_idx, 
                stopindex=tk.END,
                nocase=True
            )
            if not start_idx:
                break
            end_idx = f"{start_idx}+{len(norm_criteria)}c"
            self.student_submission_text.tag_add('search', start_idx, end_idx)
            start_idx = end_idx
        
        self.student_submission_text.tag_config('search', background='yellow')
    
    def assign_marks_to_selected(self):
        """Assign marks to currently selected not-found item"""
        if not hasattr(self, 'current_not_found_item') or not self.current_not_found_item:
            return
        
        item = self.current_not_found_item
        values = self.results_tree.item(item, 'values')
        allocated = float(values[1])
        
        # Ask for awarded marks
        awarded = simpledialog.askfloat(
            "Assign Marks",
            f"Enter awarded marks (max {allocated}):",
            parent=self.root,
            minvalue=0.0,
            maxvalue=allocated
        )
        
        if awarded is not None:
            # Get selected text from submission if any
            try:
                sel_start = self.student_submission_text.index(tk.SEL_FIRST)
                sel_end = self.student_submission_text.index(tk.SEL_LAST)
                selected_code = self.student_submission_text.get(sel_start, sel_end)
            except tk.TclError:
                selected_code = "Manually assigned"
            
            # Update the treeview
            self.results_tree.item(item, values=(
                values[0],
                values[1],
                awarded,
                f"Manually assigned: {selected_code[:50]}..." if len(selected_code) > 50 else f"Manually assigned: {selected_code}",
                values[4],
                "found"  # Update status
            ), tags=())
            
            # Clear highlights
            self.student_submission_text.tag_remove('search', '1.0', tk.END)
            
            # Update marks
            self.update_achieved_marks()
            self.assign_btn.config(state=tk.DISABLED)

    def highlight_not_found_code(self, criteria):
        """Highlight code in marking scheme that wasn't found in student submission"""
        scheme_text = self.marking_scheme_text.get(1.0, tk.END)
        
        # Find the line containing the criteria
        for line_num, line in enumerate(scheme_text.splitlines(), 1):
            if criteria.strip() in line:
                start = f"{line_num}.0"
                end = f"{line_num}.end"
                self.marking_scheme_text.tag_add('not_found', start, end)
                break 

    def highlight_matching_lines(self, criteria):
        """Highlight lines in student submission that match the criteria"""
        student_text = self.student_submission_text.get(1.0, tk.END)
        
        # Normalize the criteria for comparison
        norm_criteria = self.normalize_whitespace(criteria)
        
        # Search through each line of student code
        for line_num, line in enumerate(student_text.splitlines(), 1):
            norm_line = self.normalize_whitespace(line)
            if norm_criteria and norm_criteria in norm_line:
                start = f"{line_num}.0"
                end = f"{line_num}.end"
                self.student_submission_text.tag_add('match', start, end)
    
    def save_results_txt(self):
        """Save grading results to a text file"""
        if not self.student_name.get():
            messagebox.showerror("Error", "Please enter student name")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="Save Grading Results (TXT)",
            defaultextension=".txt",
            initialfile=f"{self.student_name.get()}_grading_results.txt",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'w') as f:
                f.write(f"Student: {self.student_name.get()}\n")
                f.write(f"Submission: {os.path.basename(self.student_submission_path.get())}\n")
                f.write(f"Total Marks: {self.total_marks.get()}\n")
                f.write(f"Achieved Marks: {self.achieved_marks.get()}\n\n")
                f.write("Detailed Breakdown:\n")
                f.write("-" * 80 + "\n")
                
                for item in self.results_tree.get_children():
                    criteria, allocated, awarded, comments = self.results_tree.item(item, 'values')
                    f.write(f"Criteria: {criteria}\n")
                    f.write(f"Allocated: {allocated}\tAwarded: {awarded}\n")
                    f.write(f"Comments: {comments}\n")
                    f.write("-" * 80 + "\n")
            
            messagebox.showinfo("Success", f"Results saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save results: {str(e)}")
    
    def save_results_excel(self):
        """Save grading results to an Excel file with proper handling of manual grades"""
        if not self.student_name.get():
            messagebox.showerror("Error", "Please enter student name")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Grading Results (Excel)",
            defaultextension=".xlsx",
            initialfile=f"{self.student_name.get()}_grading_results.xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )

        if not filepath:
            return

        try:
            # Extract data for detailed breakdown
            detailed_data = {
                'Criteria': [],
                'Allocated Marks': [],
                'Awarded Marks': [],
                'Comments': [],
                'Student Submission': []
            }

            total_allocated = 0.0
            total_awarded = 0.0

            # First pass: collect all manual grading entries and map them to their criteria
            manual_grades = {}  # Dictionary to store manually graded items
            for item in self.results_tree.get_children():
                values = self.results_tree.item(item, 'values')
                criteria = values[0]

                if criteria.startswith("Manual:"):
                    # Extract the reference criteria from either the reference field or the manual comment
                    reference_criteria = values[4] if len(values) > 4 and values[4] else criteria[7:].strip()
                    if reference_criteria:  # Only process if we have a reference
                        manual_grades[reference_criteria] = {
                            'allocated': float(values[1]),
                            'awarded': float(values[2]),
                            'comments': values[3],
                            'student_submission': criteria[7:].strip()  # The actual student code
                        }

            # Second pass: process all items
            for item in self.results_tree.get_children():
                values = self.results_tree.item(item, 'values')
                criteria = values[0]
                
                # Skip manual grading entries as we've already processed them
                if criteria.startswith("Manual:"):
                    continue
                    
                allocated = float(values[1])
                awarded = float(values[2])
                comments = values[3]
                student_submission = "-"

                # Check if this criteria has a manual grade
                if criteria in manual_grades:
                    # Use the manually graded values
                    allocated = manual_grades[criteria]['allocated']
                    awarded = manual_grades[criteria]['awarded']
                    comments = manual_grades[criteria]['comments']
                    student_submission = manual_grades[criteria]['student_submission']
                elif awarded > 0:
                    # This was automatically matched
                    student_submission = criteria

                total_allocated += allocated
                total_awarded += awarded

                detailed_data['Criteria'].append(criteria)
                detailed_data['Allocated Marks'].append(allocated)
                detailed_data['Awarded Marks'].append(awarded)
                detailed_data['Comments'].append(comments)
                detailed_data['Student Submission'].append(student_submission)

            # Add any manual grades that didn't match existing criteria (shouldn't happen but just in case)
            for ref_criteria, grade_info in manual_grades.items():
                if ref_criteria not in detailed_data['Criteria']:
                    detailed_data['Criteria'].append(ref_criteria)
                    detailed_data['Allocated Marks'].append(grade_info['allocated'])
                    detailed_data['Awarded Marks'].append(grade_info['awarded'])
                    detailed_data['Comments'].append(grade_info['comments'])
                    detailed_data['Student Submission'].append(grade_info['student_submission'])
                    total_allocated += grade_info['allocated']
                    total_awarded += grade_info['awarded']

            # Create DataFrame
            df_detailed = pd.DataFrame(detailed_data)

            # Append summary row
            summary_row = pd.DataFrame([{
                'Criteria': 'TOTAL',
                'Allocated Marks': total_allocated,
                'Awarded Marks': total_awarded,
                'Comments': '',
                'Student Submission': ''
            }])

            df_detailed = pd.concat([df_detailed, summary_row], ignore_index=True)

            # Save to Excel with xlsxwriter for formatting
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                df_detailed.to_excel(writer, sheet_name='Grading Results', index=False)
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Grading Results']
                
                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Highlight manually graded rows
                manual_format = workbook.add_format({'bg_color': '#FFF2CC'})
                
                # Apply the header format
                for col_num, value in enumerate(df_detailed.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    
                # Apply formatting and auto-adjust column widths
                for i, col in enumerate(df_detailed.columns):
                    max_len = max((
                        df_detailed[col].astype(str).map(len).max(),
                        len(col)
                    )) + 2  # Add a little extra space
                    worksheet.set_column(i, i, max_len)
                    
                    # Highlight manually graded rows
                    if col == 'Student Submission':
                        for row_num in range(1, len(df_detailed)+1):
                            if df_detailed.at[row_num-1, col] != "-" and df_detailed.at[row_num-1, col] != df_detailed.at[row_num-1, 'Criteria']:
                                worksheet.set_row(row_num, None, manual_format)

            messagebox.showinfo("Success", f"Results saved to {filepath}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file: {str(e)}")

    def update_achieved_marks(self):
        """Recalculate achieved marks from treeview"""
        total = 0.0
        for item in self.results_tree.get_children():
            values = self.results_tree.item(item, 'values')
            try:
                awarded = float(values[2]) if values[2] else 0.0
                total += awarded
            except (ValueError, IndexError):
                continue
        self.achieved_marks.set(total)

if __name__ == "__main__":
    root = tk.Tk()
    app = JavaAssessmentGrader(root)
    root.mainloop()