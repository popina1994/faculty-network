class Work(object):
    class Work_Issue(object):
        def __init__(self, authors, year):
            self._authors = authors
            self._year = year

        @property
        def authors(self):
            return self._authors

        @property
        def year(self):
            return self._year

        def __lt__(self, other):
            return self._year < other._year

        def __eq__(self, other):
            return self.__year == other.__year

        def __str__(self):
            return (str(len(self.authors)) + ": " + 
                    str(self.year))

    def __init__(self, name, isConference):
        self._name = name
        self._isConference = isConference
        self._issues = []

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def isConference(self):
        return self._isConference

    @isConference.setter
    def isConference(self, value):
        self._isConference = value

    @property
    def issues(self):
        return self._issues

    def add_issue(self, authors, year):
        issue = Work.Work_Issue(authors, year)
        self._issues.append(issue)

    def __lt__(self, other):
        return self.name < other.name

    def __eq__(self, other):
        return self.name == other.name

    def sort_issues(self):
        self._issues.sort()

    def __str__(self):
        return (self.name + ", " +
                self.isConference )


