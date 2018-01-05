class Author(object):
    
    def __init__(self, author, department):
        self._author = author
        self._department = department
        self._coauthors_num = 0

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
    def coauthors_num(self):
        return self._coauthors_num

    @coauthors_num.setter
    def coauthors_num(self, value):
        self._coauthors_num = value
