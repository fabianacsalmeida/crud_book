import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl.utils import get_column_letter

class Book:    
    def __init__(self, title, author, year):
        self.title = title
        self.author = author
        self.year = year

class BookRepository:    
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = openpyxl.load_workbook(file_path)
        self.sheet = self.workbook.active

    def create(self, book):        
        row = self.sheet.max_row + 1
        self.sheet.cell(row=row, column=1, value=book.title)
        self.sheet.cell(row=row, column=2, value=book.author)
        self.sheet.cell(row=row, column=3, value=book.year)
        self.workbook.save(self.file_path)

    def read(self):       
        books = []
        for row in self.sheet.iter_rows(values_only=True):
            books.append(Book(*row))
        return books

    def update(self, book):
        for row in range(1, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == book.title:
                self.sheet.cell(row=row, column=2, value=book.author)
                self.sheet.cell(row=row, column=3, value=book.year)
                self.workbook.save(self.file_path)
                return
        raise ValueError("Book not found")

def delete(self, title):
    for row in range(1, self.sheet.max_row + 1):
        if self.sheet.cell(row=row, column=1).value == title:
            self.sheet.delete_rows(row)
            self.workbook.save(self.file_path)
            return
    raise ValueError("Book not found")

class BookController:    
    def __init__(self, repository):
        self.repository = repository

    def create_book(self, title, author, year):
        book = Book(title, author, year)
        self.repository.create(book)

    def read_books(self):        
         return self.repository.read()

    def update_book(self, title, author, year):
        book = Book(title, author, year)
        self.repository.update(book)

    def delete_book(self, title):        
         self.repository.delete(title)

class BookView:    
    def __init__(self, master, controller):
        self.master = master
        self.controller = controller

        self.title_label = tk.Label(master, text="Title")
        self.title_label.grid(row=0, column=0)

        self.title_entry = tk.Entry(master)
        self.title_entry.grid(row=0, column=1)

        self.author_label = tk.Label(master, text="Author")
        self.author_label.grid(row=1, column=0)

        self.author_entry = tk.Entry(master)
        self.author_entry.grid(row=1, column=1)

        self.year_label = tk.Label(master, text="Year")
        self.year_label.grid(row=2, column=0)

        self.year_entry = tk.Entry(master)
        self.year_entry.grid(row=2, column=1)

        self.create_button = tk.Button(master, text="Create", command=self.create_book)
        self.create_button.grid(row=3, column=0)

        self.read_button = tk.Button(master, text="Read", command=self.read_books)
        self.read_button.grid(row=3, column=1)

        self.update_button = tk.Button(master, text="Update", command=self.update_book)
        self.update_button.grid(row=4, column=0)

        self.delete_button = tk.Button(master, text="Delete", command=self.delete_book)
        self.delete_button.grid(row=4, column=1)

        self.book_listbox = tk.Listbox(master)
        self.book_listbox.grid(row=5, column=0, columnspan=2)

    def create_book(self):        
        title = self.title_entry.get()
        author = self.author_entry.get()
        year = self.year_entry.get()
        self.controller.create_book(title, author, year)
        self.book_listbox.insert(tk.END, f"{title} by {author} ({year})")

    def read_books(self):        
        books = self.controller.read_books()
        self.book_listbox.delete(0, tk.END)
        for book in books:
            self.book_listbox.insert(tk.END, f"{book.title} by {book.author} ({book.year})")

    def update_book(self):        
        title = self.title_entry.get()
        author = self.author_entry.get()
        year = self.year_entry.get()
        self.controller.update_book(title, author, year)
        self.read_books()

    def delete_book(self):        
        title = self.title_entry.get()
        self.controller.delete_book(title)
        self.read_books()


if __name__ == "__main__":
    file_path = "books.xlsx"
    root = tk.Tk()
    root.title("Gerenciador de Livros")
    root.geometry("300x400")
    repository = BookRepository(file_path)
    controller = BookController(repository)
    BookView(root, controller)
    root.mainloop()

