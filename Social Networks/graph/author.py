class Author(object):
    
    def __init__(self, author, department):
        self._author = author
        self._department = department
        self._work_num = 0

    @property
    def author(self):
        return self.author
    
    @author.setter
    def author(self, value):
        self._author = value

    @property
    def department(self):
        return self._department

    @department.setter
    def department(self, value):
        self._department = department

    @property
    def work_num(self):
        return self._work_num

    @work_num.setter
    def work_num(self, value):
        self._work_num = value

    def __lt__(self, other):
        return self._work_num < other._work_num

    def __eq__(self, other):
        return self._work_num == other._work_num