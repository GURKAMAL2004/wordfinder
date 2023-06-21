import os
import glob
import docx
import regex as re
import tkinter as tk
from tkinter import messagebox, filedialog
import pyperclip
import requests
from bs4 import BeautifulSoup

# Function to retrieve MS Word files
def get_word_files(folder_path):
    files = glob.glob(os.path.join(folder_path, '*.docx'))
    return files

# Function to extract text from MS Word files
def extract_text_from_word(file_path):
    try:
        doc = docx.Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs]
        text = '\n'.join(paragraphs)
        return text
    except Exception as e:
        messagebox.showerror("Error", f"Error occurred while extracting text from {file_path}: {str(e)}")

# Function to extract text from a given URL
def extract_text_from_url(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        text = soup.get_text()
        return text
    except Exception as e:
        messagebox.showerror("Error", f"Error occurred while extracting text from URL: {str(e)}")

# Function to search for a word or phrase in the provided text
def search(text, word):
    results = re.findall(word, text, flags=re.IGNORECASE)
    return results

# Function to display the search results in a scrollable list
def display_results(results):
    if not results:
        messagebox.showinfo("No Results", "No matches found.")
    else:
        result_list.delete(0, tk.END)
        for result in results:
            result_list.insert(tk.END, result)

# Function to copy the selected result to the clipboard
def copy_to_clipboard():
    selected_result = result_list.get(tk.ACTIVE)
    if selected_result:
        pyperclip.copy(selected_result)
        messagebox.showinfo("Copied", "Selected result copied to clipboard.")

# Function to export the search results to a text file
def export_results():
    selected_result = result_list.get(tk.ACTIVE)
    if selected_result:
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'w') as file:
                    file.write(selected_result)
                messagebox.showinfo("Exported", "Search result exported successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Error occurred while exporting the search result: {str(e)}")

# Function to perform the search operation
def perform_search():
    search_word = search_entry.get()
    folder_path = folder_path_entry.get()
    url = url_entry.get()

    if not search_word:
        messagebox.showwarning("Empty Search", "Please enter a search word or phrase.")
        return

    if not folder_path and not url:
        messagebox.showwarning("No Input", "Please provide a folder path or URL.")
        return

    results = []
    if folder_path:
        word_files = get_word_files(folder_path)
        for file in word_files:
            text = extract_text_from_word(file)
            if text:
                results.extend(search(text, search_word))

    if url:
        text = extract_text_from_url(url)
        if text:
            results.extend(search(text, search_word))

    display_results(results)

# GUI setup
window = tk.Tk()
window.title("MS Word File Search Engine")
window.geometry("500x400")

search_label = tk.Label(window, text="Enter word or phrase to search:")
search_label.pack()

search_entry = tk.Entry(window)
search_entry.pack()

folder_path_label = tk.Label(window, text="Enter folder path (optional):")
folder_path_label.pack()

folder_path_entry = tk.Entry(window)
folder_path_entry.pack()

url_label = tk.Label(window, text="Enter URL (optional):")
url_label.pack()

url_entry = tk.Entry(window)
url_entry.pack()

search_button = tk.Button(window, text="Search", command=perform_search)
search_button.pack()

result_list = tk.Listbox(window, selectmode=tk.SINGLE)
result_list.pack(fill=tk.BOTH, expand=True)

copy_button = tk.Button(window, text="Copy Selected Result", command=copy_to_clipboard)
copy_button.pack()

export_button = tk.Button(window, text="Export Selected Result", command=export_results)
export_button.pack()

window.mainloop()
