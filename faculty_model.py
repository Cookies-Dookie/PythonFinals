class Person:
    """Base class for all people in the system."""
    
    def __init__(self, faculty_id, full_name):
        self.faculty_id = faculty_id
        self.full_name = full_name

    def __str__(self):
        return f"{self.full_name} ({self.faculty_id})"

class Faculty(Person):
    """Represents a faculty member, inheriting from Person."""
    
    def __init__(self, faculty_id, full_name, department="General"):
        super().__init__(faculty_id, full_name)
        self.department = department
        
    def get_info(self):
        return f"{self.__str__()} - Dept: {self.department}"