class Controller(object):
    def __init__(self):
        self._keys = []
        self.data_list = []

    def add_variable(self, name):
        self._keys.append(name)

    def add_dataset(self, dataset):
        data = dict()
        self.data_list.append(data)
        for key in self._keys:
            data[key] = dataset[key]

    def clear(self):
        self._keys.clear()
        self.data_list.clear()

    def print_data_list(self):
        for data in self.data_list:
            print(data)
