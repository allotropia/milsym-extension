"""
Data structure for diagram data
Python port of DataOfDiagram.java
"""

from typing import List


class DataOfDiagram:
    """Container for diagram data with level and value pairs"""
    
    class Data:
        """Inner class representing a single data point"""
        
        def __init__(self, level: int, value: str):
            self.level = level
            self.value = value
        
        def get_level(self) -> int:
            return self.level
        
        def get_value(self) -> str:
            return self.value
        
        def increase(self):
            """Increment the level"""
            self.level += 1
    
    def __init__(self):
        self._datas: List[DataOfDiagram.Data] = []
    
    def add(self, level: int, value: str):
        """Add a new data point"""
        self._datas.append(self.Data(level, value))
    
    def get(self, index: int) -> 'DataOfDiagram.Data':
        """Get data point at index"""
        return self._datas[index]
    
    def is_one_first_level_data(self) -> bool:
        """Check if there's exactly one first level data point"""
        first_level_count = 0
        for data in self._datas:
            if data.get_level() == 0:
                first_level_count += 1
                if first_level_count > 1:
                    return False
        return True
    
    def increase_levels(self):
        """Increment all data levels by 1"""
        for data in self._datas:
            data.increase()
    
    def size(self) -> int:
        """Get number of data points"""
        return len(self._datas)
    
    def is_empty(self) -> bool:
        """Check if container is empty"""
        return len(self._datas) == 0
    
    def print_data(self):
        """Print all data points for debugging"""
        if not self.is_empty():
            for data in self._datas:
                print(f"{data.get_level()} {data.get_value()}")
        else:
            print("Datas of diagram is empty")
    
    def clear(self):
        """Clear all data"""
        self._datas.clear()