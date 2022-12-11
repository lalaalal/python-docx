from docx import Document
from tkinter import *
from docx import controller
import json


class Umbrella(object):
    def __init__(self):
        self._root = Tk()
        self._controller = controller.Controller()

        var_add_frame = Frame(self._root)
        var_add_frame.grid(row=0, column=0, sticky=N+W+S)
        entry = Entry(var_add_frame)
        entry.grid(row=0, column=0, sticky=W+E+N+S)
        self._entry = entry
        add_btn = Button(var_add_frame, text="변수 추가", command=self._add_variable)
        add_btn.grid(row=0, column=1)
        self._add_btn = add_btn

        var_list = Listbox(var_add_frame)
        var_list.grid(row=1, column=0, columnspan=2, sticky=W+E+S)

        self._var_list = var_list

        data_adder = DataSetAdder(self._root, self._add_data)
        data_adder.grid(row=0, column=1, sticky=N+W)
        self._data_adder = data_adder

        data_list = Listbox(self._root)
        data_list.grid(row=0, column=2, sticky=N+E+S)
        self._data_list = data_list

        run_button = Button(self._root, text="실행", command=self._run)
        run_button.grid(row=1, column=2, sticky=E)

        clear_button = Button(self._root, text="초기화", command=self.clear)
        clear_button.grid(row=1, column=3, sticky=E)

        json_button = Button(self._root, text="파일에서", command=self._load_from_file)
        json_button.grid(row=1, column=4, sticky=E)

        self._root.mainloop()

    def _add_variable(self):
        """Add variable named in entry"""
        variable_name = self._entry.get()
        self.add_variable(variable_name)

    def add_variable(self, variable_name):
        """Add variable name"""
        self._controller.add_variable(variable_name)
        if variable_name not in self._var_list.get(0, END):
            index = self._var_list.size()
            self._var_list.insert(index, variable_name)
            self._data_adder.add_variable(variable_name)

    def _add_data(self):
        """Add a new data set from DataAdder"""
        data = self._data_adder.get_data()
        self.add_data(data)

    def add_data(self, data: dict):
        """Add a new data"""
        self._data_list.insert(0, str(data))
        self._controller.add_dataset(data)

    def _run(self):
        """Run macro"""
        index = 0
        for data in self._controller.data_list:
            template = Document("template.docx")
            print("data_set %d" % index)
            for key in data.keys():
                var_literal = "${%s}" % key
                print(var_literal + " to " + data[key])
                template.replace_word(var_literal, data[key])
            template.save(list(data.values())[0] + ".docx")

            index += 1

    def clear(self):
        """Remove all key, data"""
        self._controller.clear()
        self._var_list.delete(0, self._var_list.size())
        self._data_adder.clear()
        self._data_list.delete(0, self._data_list.size())

    def _load_from_file(self):
        """Clear and insert all data from input.json file"""
        self.clear()
        with open("input.json", "r", encoding="UTF-8") as file:
            data_list = json.load(file)
            keys = data_list[0].keys()
            for key in keys:
                self.add_variable(key)
            for data in data_list:
                self.add_data(data)


class DataSetAdder(Frame):
    def __init__(self, tk: Tk, command):
        super().__init__(tk)

        self._input_frame = Frame(self)
        self._input_frame.pack(side=BOTTOM)

        self._entries = dict()
        button = Button(self, text="데이터 추가", command=command)
        button.pack(fill=Y, side=BOTTOM, anchor=NW)

    def add_variable(self, variable_name):
        """Add new entry for given variable name"""
        row_index = len(self._entries)

        label = Label(self._input_frame, text=variable_name)
        label.grid(row=row_index, column=0)

        entry = Entry(self._input_frame)
        entry.grid(row=row_index, column=1)
        self._entries[variable_name] = entry

    def clear(self):
        for child in self._input_frame.winfo_children():
            child.destroy()
        self._entries.clear()

    def get_data(self) -> dict:
        """Returns dictionary mapped with variable name and entry value"""
        data = dict()
        for key in self._entries.keys():
            data[key] = self._entries[key].get()

        return data
