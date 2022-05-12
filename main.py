import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
import pandas as pd
from openpyxl import load_workbook
import re
import os
import shutil
import requests
from requests_html import HTMLSession
from tkinter.filedialog import askopenfilename, asksaveasfilename


class Application(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        master.title('Book Genre Search')
        master.geometry("700x300")

        self.filePath = None
        self.data = None
        self. totalRows = 0
        self.defaultSearch = "https://www.googleapis.com/books/v1/volumes?q="
        self.bookSearch = "https://www.google.com/books/edition/"

        self.genreMap = {'adventure': 'Adventure', 'children': 'Children', 'juvenile': 'Children',
                         'picture book': 'Children', 'comedy': 'Comedy', 'humor': 'Comedy', 'humour': 'Comedy',
                         'satire': 'Comedy', 'crime': 'Crime', 'noir': 'Crime', 'detective': 'Crime', 'drama': 'Drama',
                         'erotic': 'Erotic', 'fantasy': 'Fantasy', 'fairy Tale': 'Fairy Tale', 'historical': 'Historical',
                         'horror': 'Horror', 'mystery': 'Mystery', 'poetry': 'Poetry', 'christian': 'Religious Fiction',
                         'religio': 'Religious Fiction', 'romance': 'Romance', 'sci-fi': 'Sci-Fi', 'science fiction': 'Sci-Fi',
                         'scifi': 'Sci-Fi', 'space': 'Sci-Fi', 'supernatural': 'Supernatural', 'paranormal': 'Supernatural',
                         'ghost': 'Supernatural', 'suspense': 'Thriller', 'thriller': 'Thriller', 'western': 'Western',
                         'young adult': 'Young Adult', 'new adult': 'Young Adult', 'spy': 'Spy', 'parady': 'Parady',
                         'goth': 'Gothic', 'dystopia': 'Dystopian'}

        self.fileName = tk.StringVar()
        self.fileName.set("Select File")

        self.create_widgets()

    # Initalize GUI
    def create_widgets(self):

        # GUI Text
        self.part_label = tk.Label(
            self.master, text="Please choose an excel file to edit", font=('bold', 14), pady=20)
        self.part_label.grid(row=0, column=0, sticky=tk.W)

        # Buttons
        self.add_btn = tk.Button(
            self.master, textvariable=self.fileName, width=12, command=self.add_item)
        self.add_btn.grid(row=2, column=0, pady=20)

        self.remove_btn = tk.Button(
            self.master, text="Run Program", width=12, command=self.run_file)
        self.remove_btn.grid(row=2, column=1)

        self.prog_bar = ttk.Progressbar(
            orient=HORIZONTAL, length=100, mode='determinate')
        self.prog_bar.grid(row=3, column=0)

    # Add file to update
    def add_item(self):
        Tk().withdraw()
        self.filePath = askopenfilename()

        if not self.filePath.endswith('.csv') and not self.filePath.endswith('.xlsx'):
            tk.messagebox.showerror(
                title=None, message="Please upload a CSV or XLSX file")
            self.filePath = None
        elif self.filePath.endswith('.csv'):
            self.fileName.set(self.filePath)
            self.data = pd.read_csv(self.filePath)
        else:
            self.fileName.set(self.filePath)
            self.data = pd.read_excel(self.filePath, sheet_name=0)

    # main script: goes through uploaded file and updates the genre
    def run_file(self):
        Tk().withdraw()
        self.toggle_buttons()
        # detect which columns to use; if cant detect return early
        authorCol, bookCol, genreCol = self.find_author_book_genre()
        if authorCol == None or bookCol == None or genreCol == None:
            tk.messagebox.showinfo(
                message='Please choose a file with headers of "genre", "title", and "author"!')
            return

        # iterate through the dataframe; each call scrapes the web for genre
        self.totalRows = len(self.data.index)
        for index, row in self.data.iterrows():
            if pd.isnull(row[genreCol]) and not pd.isnull(row[bookCol]):
                try:
                    # adjust books name
                    name = self.update_name(row[bookCol])
                    author = self.update_author(row[authorCol])
                    curSearch = self.defaultSearch + name + author

                    # call google api
                    genreSearch = self.get_web_link(curSearch)

                    # Scrape googles book page for genre
                    self.data.iat[index, genreCol] = self.get_genre(
                        genreSearch)

                except Exception as e:
                    self.data.iat[index, genreCol] = 'Unknown'

            self.update_progres_bar()

        self.data = self.data.loc[:, ~self.data.columns.str.contains(
            '^Unnamed')]  # drop empty columns

        self.save_quit()  # save the new csv file and quit the app

    # removes any user formatting
    def update_name(self, temp):
        name = re.sub("[\(\[].*?[\)\]]", "", temp)
        if name[0] == ' ':
            name = name[1:]
        return name

    # updates the authors name to work in the search
    def update_author(self, temp):
        if ',' in temp:
            temp = temp.split(',')[0]
        return ' ' + temp

    # get the site that contains the genres
    def get_web_link(self, originalLink):
        response = requests.get(originalLink)
        result = response.json()
        book = result['items'][0]['volumeInfo']['title'].replace(
            ' ', '_')  # get exact title
        id = result['items'][0]['id']  # gets id

        return self.bookSearch + book + '/' + id

    # gets the genres
    def get_genre(self, link):
        session = HTMLSession()
        r = session.get(link)  # input html value
        articles = r.html.find('.Z1hOCe')

        # look for genres
        for article in articles:
            if "Genre" in article.text:
                genres = self.standardize_genres(article.text.lower())
                return genres

        # genres not present, attempting to use subjects instead
        for article in articles:
            if "Subject" in article.text:
                genres = self.standardize_genres(article.text.lower())
                return genres

    # Saves csv file and exits the application
    def save_quit(self):
        self.toggle_buttons()
        if self.filePath[-1] == 'v':
            saveLocation = asksaveasfilename(
                defaultextension='.csv', filetypes=[('CSV File', '.csv')])
            self.data.to_csv(saveLocation)
        else:
            saveLocation = asksaveasfilename(
                defaultextension='.xlsx', filetypes=[('Excel File', '.xlsx')])

            source = open(self.filePath, 'rb')
            destination = open(saveLocation, 'wb')
            shutil.copyfileobj(source, destination)

            with pd.ExcelWriter(saveLocation, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                self.data.to_excel(writer, sheet_name='Fiction', index=False)

        # popup to show the file is saved
        tk.messagebox.showinfo(message='File was saved!')
        app.quit()

    # inform user of scrape progress
    def update_progres_bar(self):
        self.prog_bar['value'] += 100 / self.totalRows
        app.update()

    # converts the string from the books website into a consistent string
    def standardize_genres(self, rawValues):
        genres = set()
        for key, value in self.genreMap.items():
            if key in rawValues:
                genres.add(value)
        if genres:
            genres = ', '.join(genres)
        else:
            genres = 'Other'
        return genres

    def find_author_book_genre(self):
        genre, author, book = None, None, None

        for i in range(len(self.data.columns)):
            temp = self.data.columns[i].lower()
            if 'genre' in temp:
                genre = i
            if 'author' in temp:
                author = i
            if 'title' in temp:
                book = i
        return author, book, genre

    def toggle_buttons(self):
        if self.remove_btn['state'] == 'normal':
            self.remove_btn['state'] = 'disabled'
            self. add_btn['state'] = 'disabled'
        else:
            self.remove_btn['state'] = 'normal'
            self. add_btn['state'] = 'normal'

root = tk.Tk()
app = Application(master=root)
app.mainloop()


# TODO: make sure user cant press button while the script is active
# TODO: adjust where the file name shows, currently too long for the box, maybe line below
