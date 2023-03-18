import tkinter as tk
from tkinter import ttk, Frame, Label, PhotoImage
import smartsheet
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter.ttk import Treeview

class MainWindow:
    def __init__(self, master, api_key):
        self.master = master
        self.master.title("AEM Dashboard")
        master.configure(background='white')
        self.api_key = api_key


        # Creating frames
        self.main_frame = MainFrame(master)
        self.usa_frame = USAFrame(master)
        self.costa_rica_frame = CostaRicaFrame(master, self, self.api_key)
        self.eu_frame = EUFrame(master)

        style = ttk.Style()
        style.configure("TButton", relief="flat", background="darkgrey", padding = 10, width = 20, height = 5, borderwidth = 0, font = ('Helvetica', 12, 'bold'), shape = 'rectangle')

        # Creating buttons
        self.button1 = ttk.Button(master, text='Global Operations', style='TButton', command=self.show_main_frame)
        self.button1.place(relx=0.9, rely=0.3, anchor="center")
        self.button2 = ttk.Button(master, text='USA', style='TButton', command=self.show_usa)
        self.button2.place(relx=0.9, rely=0.4, anchor="center")
        self.button3 = ttk.Button(master, text='Costa Rica', style='TButton', command=self.show_costa_rica)
        self.button3.place(relx=0.9, rely=0.5, anchor="center")
        self.button4 = ttk.Button(master, text='EU', style='TButton', command=self.show_eu)
        self.button4.place(relx=0.9, rely=0.6, anchor="center")

        self.main_frame.grid(row=0, column=1, sticky='nsew')

    def ecn_status(self, master):
        sheet_id = 4541475754141572
        ss = smartsheet.Smartsheet(self.api_key)
        sheet = ss.Sheets.get_sheet(sheet_id)
        rows = sheet.rows
        columns = sheet.columns

        # Creating a list of columns for the Treeview
        columns_list = [col.id for col in columns]

        # Creating a list of data for the Treeview
        data_list = [[col.title for col in columns]]
        for row in rows:
            data_list.append([cell.value for cell in row.cells])

        # Creating a frame to hold the treeview and scrollbars
        tree_frame = tk.Frame(master)
        tree_frame.grid(row=1, column=0, sticky='nsew')

        # Creating a Treeview
        tree = ttk.Treeview(tree_frame, columns=columns_list, show='headings')

        # Adding columns to the Treeview
        for col in columns:
            tree.heading(col.id, text=col.title, anchor='w')

        # Adding data to the Treeview
        for data in data_list:
            tree.insert('', 'end', values=data)

        # Setting the appearance of the Treeview
        style = ttk.Style()
        style.configure('Treeview', rowheight=60, background='white', foreground='black', font=('Calibri', 12))
        style.configure('Treeview.Heading', background='white', foreground='black', relief='solid', font=('Calibri', 14, 'bold'))
        style.configure('Treeview.Column.0', background='forestgreen')  # Target the first column by name

        for col in columns_list:
            style.configure(f'Treeview.{col}', borderwidth=1, relief='solid')

        # Changing background color of the Treeview on the first column
        style.map('Treeview', background=[('selected', 'forestgreen')])

        # Adding scrollbars to the Treeview
        xscrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        xscrollbar.grid(row=1, column=0, sticky='we')
        yscrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        yscrollbar.grid(row=0, column=1, sticky='ns')
        tree.configure(xscrollcommand=xscrollbar.set, yscrollcommand=yscrollbar.set)

        # Grid the Treeview
        tree.grid(row=0, column=0, sticky='nsew')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)


    def show_main_frame(self):
        self.main_frame.grid(row=0, column=1, sticky='nsew')
        self.usa_frame.grid_forget()
        self.costa_rica_frame.grid_forget()
        self.eu_frame.grid_forget()

        # Make the original buttons visible again
        self.button1.place(relx=0.9, rely=0.3, anchor="center")
        self.button2.place(relx=0.9, rely=0.4, anchor="center")
        self.button3.place(relx=0.9, rely=0.5, anchor="center")
        self.button4.place(relx=0.9, rely=0.6, anchor="center")


    def show_usa(self):
        self.usa_frame.grid(row=0, column=1, sticky='nsew')
        self.main_frame.grid_forget()
        self.costa_rica_frame.grid_forget()
        self.eu_frame.grid_forget()

    def show_costa_rica(self):
        self.costa_rica_frame.grid(row=0, column=1, sticky='nsew')
        self.main_frame.grid_forget()
        self.usa_frame.grid_forget()
        self.eu_frame.grid_forget()
        
        #Hide the original buttons
        self.button1.place_forget()
        self.button2.place_forget()
        self.button3.place_forget()
        self.button4.place_forget()


    def show_eu(self):
        self.eu_frame.grid(row=0, column=1, sticky='nsew')
        self.main_frame.grid_forget()
        self.usa_frame.grid_forget()
        self.costa_rica_frame.grid_forget()

class MainFrame(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.configure(bg='white')
        label = Label(self, text='Global Operations', bg='white', font=('Segoe UI', 20, 'bold'))
        label.grid()

class USAFrame(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.configure(bg='white')
        label = Label(self, text='USA', bg='white', font=('Segoe UI', 20, 'bold'))
        label.grid()


class CostaRicaFrame(Frame):
    def __init__(self, master, app, api_key):
        super().__init__(master)
        self.app = app
        self.api_key = api_key
        self.configure(bg='white')
        label = Label(self, text='Costa Rica', bg='white', font=('Segoe UI', 20, 'bold'))
        label.grid(row=0, column=0, pady=10)

        #Add ECN Status function inside the class
        self.ecn_status = lambda: app.ecn_status(self)

        style = ttk.Style()
        style.configure("Graph.TButton", relief="groove", background="green", borderwidth=10)
        style.configure("CR.TButton", relief="flat", background="darkgrey", padding=10, width=20, height=5, borderwidth=0, font=('Helvetica', 12, 'bold'), shape='rectangle')


        style = ttk.Style()
        style.configure("CR.TButton", relief="flat", background="darkgrey", padding=10, width=20, height=5, borderwidth=0, font=('Helvetica', 12, 'bold'), shape='rectangle')

        self.button1 = ttk.Button(self, text='ECN Status', style='CR.TButton', command= self.ecn_status)
        self.button1.grid(row=1, column=0, pady=10)
        self.button2 = ttk.Button(self, text='CRAT', style='CR.TButton')
        self.button2.grid(row=2, column=0, pady=10)
        self.button3 = ttk.Button(self, text='CRLVE', style='CR.TButton')
        self.button3.grid(row=3, column=0, pady=10)
        self.button4 = ttk.Button(self, text='Plan', style='CR.TButton')
        self.button4.grid(row=4, column=0, pady=10)
        self.button5 = ttk.Button(self, text='Training', style='CR.TButton')
        self.button5.grid(row=5, column=0, pady=10)
        self.button6 = ttk.Button(self, text='Return', style='CR.TButton', command=self.app.show_main_frame)
        self.button6.grid(row=6, column=0, pady=10)


        # Call the create_graph function to generate the graph
        img = self.create_graph()

        # CHANGED: Create a Button widget with the image instead of a Label widget
        self.graph_button = ttk.Button(self, text='', style='Graph.TButton', image=img, compound='center', command=self.display_more_data)
        self.graph_button.image = img
        # CHANGED: Use grid method to place the button with rowspan to keep the size
        self.graph_button.grid(row=1, column=1, rowspan=6, pady=10, padx=10, sticky='nsew')

        # ADDED: Configure the column weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)



    def create_graph(self):
        # Adding a graph
        sheet_id = 1998891726989188
        ss = smartsheet.Smartsheet(self.api_key)
        sheet = ss.Sheets.get_sheet(sheet_id)
        rows = sheet.rows
        columns = sheet.columns
        df1 = pd.DataFrame(columns=[col.title for col in columns])
        for row in rows:
            df1.loc[row.id] = [col.value for col in row.cells]
        df1.drop(df1.columns[0], axis=1, inplace=True)
        df1.reset_index(drop=True, inplace=True)
        df1 = df1.loc[41:42]
        df1.drop(df1.columns[0:7], axis=1, inplace=True)
        fig, ax = plt.subplots()
        # Size of the graph
        fig.set_size_inches(10, 5)
        df1.iloc[1].plot(kind='line', color='lightcoral', linewidth=2, ax=ax)
        df1.iloc[0].plot(kind='bar', color='forestgreen', linewidth=1, ax=ax)
        ax.set_xlabel('Work Week')
        ax.set_ylabel('Count')

        # Set the y-axis ticks to integers
        y_ticks = np.arange(0, max(df1.iloc[0]) + 1, 1)
        ax.set_yticks(y_ticks)

        # Setting x-axis ticks to integers
        x_ticks = np.arange(0, len(df1.columns), 1)
        ax.set_xticks(x_ticks)

        ax.legend(['Total HC Available', 'Total HC Needed'], loc='upper right')
        # Adding red labels to the bars
        for p in ax.patches:
            ax.annotate(str(int(p.get_height())), (p.get_x() * 1.005, p.get_height() * 1.005), rotation=270)

        # Adding labels to the line graph
        for i, v in enumerate(df1.iloc[1]):
            ax.text(i, v, str(int(v)), color='firebrick', fontweight='bold', rotation=270)

        # Changing background color
        ax.set_facecolor('white')
        # Changing grid color
        ax.grid(color='forestgreen', linestyle='-', linewidth=0.25, alpha=0.5)

        # Changing grid color
        ax.grid(color='forestgreen', linestyle='-', linewidth=0.25, alpha=0.5)

        # Adding highlighted line surrounding the graph
        ax.spines['top'].set_color('forestgreen')
        ax.spines['bottom'].set_color('forestgreen')
        ax.spines['left'].set_color('forestgreen')
        ax.spines['right'].set_color('forestgreen')
        ax.tick_params(axis='x', colors='forestgreen')
        ax.tick_params(axis='y', colors='forestgreen')
        ax.yaxis.label.set_color('forestgreen')
        ax.xaxis.label.set_color('forestgreen')
        ax.title.set_color('forestgreen')
        plt.xticks(rotation=90)

        # Save the figure to a file and load it as a PhotoImage
        fig.savefig("graph.png")
        img = PhotoImage(file="graph.png")

        return img
    
    def display_more_data(self):
        def plot_quarter_data(ax, data, title):
            data.iloc[1].plot(kind='line', color='lightcoral', linewidth=2, ax=ax)
            data.iloc[0].plot(kind='bar', color='forestgreen', linewidth=1, ax=ax)
            ax.set_title(title, y=1.15)
            ax.set_xlabel('Work Week')
            ax.set_ylabel('Count')
            ax.legend(['Total HC Available', 'Total HC Needed'], loc='upper right')
            for p in ax.patches:
                ax.annotate(str(int(p.get_height())), (p.get_x() * 1.005, p.get_height() * 1.005), rotation=270)
            for i, v in enumerate(data.iloc[1]):
                ax.text(i, v, str(int(v)), color='firebrick', fontweight='bold', rotation=270)

        # Data fetching and processing
        sheet_id = 1998891726989188
        ss = smartsheet.Smartsheet(self.api_key)
        sheet = ss.Sheets.get_sheet(sheet_id)
        rows = sheet.rows
        columns = sheet.columns
        df1 = pd.DataFrame(columns=[col.title for col in columns])
        for row in rows:
            df1.loc[row.id] = [col.value for col in row.cells]
        df1.drop(df1.columns[0], axis=1, inplace=True)
        df1.reset_index(drop=True, inplace=True)
        df1.drop(df1.columns[0:7], axis=1, inplace=True)
        df1.reset_index(drop=True, inplace=True)
        Q1 = df1.iloc[:, :13]
        Q2 = df1.iloc[:, 13:26]
        Q3 = df1.iloc[:, 26:39]
        Q4 = df1.iloc[:, 39:53]

        Q1 = Q1.loc[41:42]
        Q2 = Q2.loc[41:42]
        Q3 = Q3.loc[41:42]
        Q4 = Q4.loc[41:42]

        # Convert the data to integer type
        Q1 = Q1.astype(int)
        Q2 = Q2.astype(int)
        Q3 = Q3.astype(int)
        Q4 = Q4.astype(int)

        fig = plt.Figure(figsize=(10, 8))
        axs = fig.subplots(nrows=2, ncols=2)
        fig.subplots_adjust(hspace=0.5, wspace=0.3)

        plot_quarter_data(axs[0, 0], Q1, 'Q1')
        plot_quarter_data(axs[0, 1], Q2, 'Q2')
        plot_quarter_data(axs[1, 0], Q3, 'Q3')
        plot_quarter_data(axs[1, 1], Q4, 'Q4')

        # Existing canvas and button code
        self.canvas = FigureCanvasTkAgg(fig, self)
        self.canvas.get_tk_widget().grid(row=1, column=0, rowspan=4, sticky="nsew", pady=(25, 0), padx=(25, 0))

        self.graph_button.grid_remove()

        self.hide_data_button = tk.Button(self, text="Hide Data", command=self.hide_more_data, height=3, width=20, bg="darkgrey")
        self.hide_data_button.grid(row=1, column=1, pady=(10, 0), padx=(10, 0), sticky='ne')  # Changed position and added padx

    def hide_more_data(self):
        self.canvas.get_tk_widget().grid_remove()
        self.hide_data_button.grid_remove()

        # Re-display the graph_button
        self.graph_button.grid(row=1, column=1, rowspan=6, pady=10, padx=10, sticky='nsew')



class EUFrame(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.configure(bg='white')
        label = Label(self, text = 'EU', bg = 'white', font = ('Segoe UI', 20, 'bold'))
        label.grid()

root = tk.Tk()
api_key = 'GHwvGA29rHRE418p3YNLSUTEw6D43rSW6neIm'
app = MainWindow(root, api_key)
root.mainloop()
