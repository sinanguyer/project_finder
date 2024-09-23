import math
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import os
import re
import shutil
from docx import Document
import fitz  # PyMuPDF
import pandas as pd  # For Excel handling
from transformers import pipeline
from sentence_transformers import SentenceTransformer, util

# Load Hugging Face models
summarizer = pipeline("summarization")
classifier = pipeline("zero-shot-classification")
qa_model = pipeline("question-answering")
sentence_model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')

# Load a spell checker model
spell_checker = pipeline("fill-mask", model="bert-base-uncased")

# Function to split text into chunks of a manageable size
def split_text_into_chunks(text, chunk_size=1024):
    words = text.split()
    chunks = [' '.join(words[i:i+chunk_size]) for i in range(0, len(words), chunk_size)]
    return chunks

class DocumentMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Filtering, Merging, Summarization, and Classification Application")

        # Select Base Folder
        self.input_label = tk.Label(root, text="Select Base Folder:")
        self.input_label.grid(row=0, column=0, padx=10, pady=10)
        self.input_button = tk.Button(root, text="Browse", command=self.browse_base_folder)
        self.input_button.grid(row=0, column=1, padx=10, pady=10)
        self.base_folder = tk.StringVar()
        self.base_folder_entry = tk.Entry(root, textvariable=self.base_folder, width=40)
        self.base_folder_entry.grid(row=0, column=2, padx=10, pady=10)
        
        # First filter: File name filter (PDF, DOCX, Excel)
        self.filter_label = tk.Label(root, text="Enter File Name Filter Terms (comma-separated):")
        self.filter_label.grid(row=1, column=0, padx=10, pady=10)
        self.filter_entry = tk.Entry(root)
        self.filter_entry.grid(row=1, column=2, padx=10, pady=10)
        self.filter_button = tk.Button(root, text="Start First Filter (By Name)", command=self.apply_first_filter)
        self.filter_button.grid(row=1, column=3, padx=10, pady=10)

        # Second filter: Content search filter
        self.content_filter_label = tk.Label(root, text="Enter Content Filter Terms (comma-separated):")
        self.content_filter_label.grid(row=2, column=0, padx=10, pady=10)
        self.content_filter_entry = tk.Entry(root)
        self.content_filter_entry.grid(row=2, column=2, padx=10, pady=10)
        self.content_filter_button = tk.Button(root, text="Start Second Filter (By Content)", command=self.apply_second_filter)
        self.content_filter_button.grid(row=2, column=3, padx=10, pady=10)

        # Third filter: Narrow search filter
        self.narrow_filter_label = tk.Label(root, text="Enter Narrowing Filter Terms (comma-separated):")
        self.narrow_filter_label.grid(row=3, column=0, padx=10, pady=10)
        self.narrow_filter_entry = tk.Entry(root)
        self.narrow_filter_entry.grid(row=3, column=2, padx=10, pady=10)
        self.narrow_filter_button = tk.Button(root, text="Start Third Filter (Narrowing)", command=self.apply_third_filter)
        self.narrow_filter_button.grid(row=3, column=3, padx=10, pady=10)

        # Button to remove PDFs with zero count
        self.remove_zero_button = tk.Button(root, text="Remove PDFs with Zero Counts", command=self.remove_zero_count_pdfs)
        self.remove_zero_button.grid(row=4, column=1, padx=10, pady=10)

        # Display for filtered files
        self.filtered_files_text = scrolledtext.ScrolledText(root, width=60, height=10)
        self.filtered_files_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

        # Select Destination Folder
        self.save_label = tk.Label(root, text="Select Destination Folder:")
        self.save_label.grid(row=6, column=0, padx=10, pady=10)
        self.save_button = tk.Button(root, text="Browse", command=self.browse_save_folder)
        self.save_button.grid(row=6, column=1, padx=10, pady=10)
        self.save_folder = tk.StringVar()
        self.save_folder_entry = tk.Entry(root, textvariable=self.save_folder, width=40)
        self.save_folder_entry.grid(row=6, column=2, padx=10, pady=10)

        # Copy Filtered Files Button
        self.copy_button = tk.Button(root, text="Copy Filtered Files", command=self.copy_filtered_files)
        self.copy_button.grid(row=7, column=1, padx=10, pady=10)

        # PDF Chatting Space
        self.chat_label = tk.Label(root, text="PDF Chatting Area:")
        self.chat_label.grid(row=8, column=0, padx=10, pady=10)
        self.chat_box = scrolledtext.ScrolledText(root, width=60, height=10)
        self.chat_box.grid(row=9, column=0, columnspan=3, padx=10, pady=10)

        # Ask Question Button
        self.ask_button = tk.Button(root, text="Ask Question", command=self.ask_question)
        self.ask_button.grid(row=10, column=1, padx=10, pady=10)

        # Variables to track filtered files and term counters
        self.current_filtered_files = {}
        self.term_counters = {}
        self.filtered_texts = {}

    def browse_base_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.base_folder.set(folder_selected)

    def browse_save_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.save_folder.set(folder_selected)

    # First filter: File name filtering
    def apply_first_filter(self):
        base_folder = self.base_folder.get()
        filter_terms = self.filter_entry.get()

        if not base_folder or not filter_terms:
            messagebox.showerror("Error", "Please select a base folder and enter file name filter terms.")
            return

        filter_terms_list = [term.strip() for term in filter_terms.split(",") if term.strip()]

        self.current_filtered_files = self.filter_by_filename(base_folder, filter_terms_list)

        if not self.current_filtered_files:
            messagebox.showinfo("No Files Found", "No files matched the file name filter criteria.")
            return

        self.display_filtered_files()

    # Second filter: Content filtering
    def apply_second_filter(self):
        content_filter_terms = self.content_filter_entry.get()

        if not self.current_filtered_files or not content_filter_terms:
            messagebox.showerror("Error", "Please apply the first filter and enter content filter terms.")
            return

        content_filter_list = [term.strip() for term in content_filter_terms.split(",") if term.strip()]
        self.current_filtered_files = self.filter_by_content(self.current_filtered_files, content_filter_list)

        self.display_filtered_files()

    # Third filter: Narrowing filter
    def apply_third_filter(self):
        narrow_filter_terms = self.narrow_filter_entry.get()

        if not self.current_filtered_files or not narrow_filter_terms:
            messagebox.showerror("Error", "Please apply the second filter and enter narrowing filter terms.")
            return

        narrow_filter_list = [term.strip() for term in narrow_filter_terms.split(",") if term.strip()]
        self.current_filtered_files = self.filter_by_content(self.current_filtered_files, narrow_filter_list)

        self.display_filtered_files()

    def filter_by_filename(self, folder, filter_terms_list):
        matched_files = {}
        for root, dirs, files in os.walk(folder):
            for file in files:
                if any(term.lower() in file.lower() for term in filter_terms_list) and file.endswith(('.pdf', '.docx', '.xlsx')):
                    matched_files[os.path.join(root, file)] = {term: file.lower().count(term.lower()) for term in filter_terms_list}
        return matched_files

    def filter_by_content(self, file_list, filter_terms_list):
        filtered_files = {}
        for file, counts in file_list.items():
            new_counts = self.file_contains_terms(file, filter_terms_list)
            if any(new_counts.values()):
                filtered_files[file] = new_counts
        return filtered_files

    def file_contains_terms(self, file_path, filter_terms_list):
        counts = {term: 0 for term in filter_terms_list}
        try:
            if file_path.endswith('.docx'):
                text = self.read_docx(file_path)
            elif file_path.endswith('.pdf'):
                text = self.read_pdf(file_path)
            elif file_path.endswith('.xlsx'):
                text = self.read_excel(file_path)
            else:
                return counts

            for term in filter_terms_list:
                counts[term] = len(re.findall(re.escape(term), text, re.IGNORECASE))

        except Exception as e:
            print(f"Error reading file {file_path}: {e}")

        return counts

    def read_docx(self, file_path):
        try:
            doc = Document(file_path)
            return "\n".join([para.text for para in doc.paragraphs])
        except Exception as e:
            print(f"Error reading DOCX file {file_path}: {e}")
            return ""

    def read_pdf(self, file_path):
        try:
            text = ""
            doc = fitz.open(file_path)
            for page in doc:
                text += page.get_text()
            return text
        except Exception as e:
            print(f"Error reading PDF file {file_path}: {e}")
            return ""

    def read_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            return df.to_string()
        except Exception as e:
            print(f"Error reading Excel file {file_path}: {e}")
            return ""

    def remove_zero_count_pdfs(self):
        # Remove files with any zero-count term
        self.current_filtered_files = {file: counts for file, counts in self.current_filtered_files.items() if all(count > 0 for count in counts.values())}
        self.display_filtered_files()

    def display_filtered_files(self):
        self.filtered_files_text.delete('1.0', tk.END)
        for file, counts in self.current_filtered_files.items():
            total_count = sum(counts.values())
            self.filtered_files_text.insert(tk.END, f"{file}: {counts}, Total: {total_count}\n")

    def copy_filtered_files(self):
        save_folder = self.save_folder.get()
        if not save_folder:
            messagebox.showerror("Error", "Please select a destination folder.")
            return

        for file in self.current_filtered_files:
            try:
                shutil.copy(file, save_folder)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy {file}: {str(e)}")

        messagebox.showinfo("Success", "Filtered files have been copied successfully.")

    # PDF Chat Functionality
    def extract_texts_for_chat(self):
        self.filtered_texts = {}
        for file in self.current_filtered_files:
            if file.endswith('.pdf'):
                text = self.read_pdf(file)
                if text:
                    self.filtered_texts[file] = text
            elif file.endswith('.docx'):
                text = self.read_docx(file)
                if text:
                    self.filtered_texts[file] = text

    def ask_question(self):
        question = simpledialog.askstring("Ask Question", "Enter your question about the filtered PDFs:")
        if not question:
            return

        # Extract texts for filtered files
        self.extract_texts_for_chat()

        # Perform question answering
        for file, text in self.filtered_texts.items():
            chunks = split_text_into_chunks(text)
            for chunk in chunks:
                answer = qa_model(question=question, context=chunk)
                if answer and answer['score'] > 0.1:
                    self.chat_box.insert(tk.END, f"Answer from {file}: {answer['answer']}\n\n")
                else:
                    self.chat_box.insert(tk.END, f"No relevant answer found in {file}.\n\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentMergerApp(root)
    root.mainloop()
