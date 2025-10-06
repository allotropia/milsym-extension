# OxygenOffice SmART Python Port

This is a Python port of the organization chart functionality from the OxygenOffice SmART LibreOffice extension. The original Java extension provides various diagram creation tools for LibreOffice, and this port focuses specifically on the organization chart (`orgchart`) component.

## Overview

The original Java codebase has been converted to Python while maintaining the same object-oriented structure and design patterns. The code provides:

- **Organization Chart Creation**: Create hierarchical organization charts from data
- **Tree Structure Management**: Handle parent-child relationships in organizational data
- **LibreOffice Integration**: Work with LibreOffice Draw documents via UNO API
- **Visual Customization**: Support for various colors, styles, and layouts

## Architecture

The Python port maintains the same class hierarchy as the original Java code:

```
Diagram (base class)
└── OrganizationChart (abstract base for org charts)
    └── OrgChart (main implementation)
        ├── OrgChartTree (manages tree structure)
        └── OrgChartTreeItem (individual tree nodes)
```

## Key Components

### Core Classes

- **`OrgChart`**: Main organization chart implementation
- **`OrgChartTree`**: Manages the tree structure and shape relationships
- **`OrgChartTreeItem`**: Represents individual items in the organization chart
- **`DataOfDiagram`**: Container for hierarchical data input
- **`SchemeDefinitions`**: Color scheme definitions

### Supporting Classes

- **`Controller`**: Main controller (stub implementation)
- **`Gui`**: GUI handling (stub implementation)
- **`Diagram`**: Base diagram functionality

## Installation

### Prerequisites

1. **LibreOffice**: Must be installed with Python UNO bindings
2. **Python 3.8+**: Required for type annotations and modern features

### Setup

1. Clone or download the code
2. Ensure LibreOffice UNO bindings are accessible:

```bash
# Ubuntu/Debian
export PYTHONPATH=$PYTHONPATH:/usr/lib/libreoffice/program

# Windows
set PYTHONPATH=%PYTHONPATH%;C:\Program Files\LibreOffice\program

# macOS
export PYTHONPATH=$PYTHONPATH:/Applications/LibreOffice.app/Contents/Resources
```

3. Install development dependencies (optional):

```bash
pip install -r requirements-dev.txt
```

## Usage

### Basic Example

```python
from oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.orgchart import OrgChart
from oxygenoffice.extensions.smart.diagram.data_of_diagram import DataOfDiagram
from oxygenoffice.extensions.smart.controller import Controller, Gui

# Create controller and GUI instances
controller = Controller()
gui = Gui()

# Create organization chart
org_chart = OrgChart(controller, gui, None)

# Create hierarchical data
data = DataOfDiagram()
data.add(0, "CEO")           # Level 0 (root)
data.add(1, "CTO")           # Level 1 (reports to CEO)
data.add(1, "CFO")           # Level 1 (reports to CEO)
data.add(2, "Dev Lead")      # Level 2 (reports to CTO)
data.add(2, "QA Lead")       # Level 2 (reports to CTO)

# Create the diagram
org_chart.create_diagram(data)
```

### Data Structure

The `DataOfDiagram` class expects hierarchical data where:
- **Level 0**: Root level (typically one item)
- **Level 1**: Direct reports to root
- **Level 2**: Second level, etc.

```python
data = DataOfDiagram()
data.add(level, "Display Text")
```

## LibreOffice UNO Integration

**Important**: The current implementation contains stub methods for UNO integration. To work with actual LibreOffice documents, you need to:

1. Replace stub imports with real UNO imports
2. Implement actual shape creation and manipulation
3. Handle document lifecycle and connection

See `UNO_INTEGRATION_GUIDE.py` for detailed instructions on converting stubs to working UNO code.

### Example UNO Integration

```python
# Real UNO imports (replace stubs)
import uno
from com.sun.star.awt import Point, Size
from com.sun.star.drawing import XShape
from com.sun.star.beans import XPropertySet

# Example real implementation
def create_shape_real(service_manager, draw_page, shape_type, x, y, width, height):
    shape = service_manager.createInstance(f"com.sun.star.drawing.{shape_type}")
    
    point = Point()
    point.X = x
    point.Y = y
    shape.setPosition(point)
    
    size = Size()
    size.Width = width
    size.Height = height
    shape.setSize(size)
    
    draw_page.add(shape)
    return shape
```

## Features

### Supported Functionality

- ✅ **Hierarchical Data Processing**: Convert flat data to tree structures
- ✅ **Tree Management**: Parent-child relationships, siblings, positioning
- ✅ **Layout Algorithms**: Horizontal and vertical layouts for different levels
- ✅ **Color Schemes**: Multiple predefined color schemes and themes
- ✅ **Shape Management**: Rectangle and connector shape handling
- ✅ **Dynamic Growth**: Add new nodes to existing diagrams

### Stub Implementations

- 🔶 **LibreOffice Integration**: UNO API calls (requires implementation)
- 🔶 **Shape Creation**: Actual LibreOffice shape creation
- 🔶 **Document Management**: LibreOffice document lifecycle
- 🔶 **GUI Integration**: User interface components

## Development

### Code Structure

```
python/
├── oxygenoffice/
│   └── extensions/
│       └── smart/
│           ├── controller.py              # Main controller (stub)
│           ├── diagram/
│           │   ├── diagram.py             # Base diagram class
│           │   ├── data_of_diagram.py     # Data container
│           │   ├── scheme_definitions.py  # Color schemes
│           │   └── organizationcharts/
│           │       ├── organization_chart.py          # Base org chart
│           │       ├── organization_chart_tree.py     # Base tree
│           │       ├── organization_chart_tree_item.py # Base tree item
│           │       └── orgchart/
│           │           ├── orgchart.py           # Main implementation
│           │           ├── orgchart_tree.py      # Tree implementation
│           │           └── orgchart_tree_item.py # Tree item implementation
│           └── gui/
├── UNO_INTEGRATION_GUIDE.py  # Guide for UNO integration
├── pyproject.toml             # Project configuration
├── requirements.txt           # Runtime requirements
└── requirements-dev.txt       # Development requirements
```

### Testing

```bash
# Run tests (after implementing test cases)
pytest

# Code formatting
black .

# Linting
flake8

# Type checking
mypy oxygenoffice/
```

## Migration from Java

The Python port maintains compatibility with the original Java design:

| Java Concept | Python Equivalent |
|--------------|-------------------|
| `ArrayList<T>` | `List[T]` |
| `abstract class` | `ABC` with `@abstractmethod` |
| `static` variables | Class variables |
| `UnoRuntime.queryInterface()` | `shape.queryInterface()` |
| `XPropertySet.setPropertyValue()` | Same UNO API |
| Java exceptions | Python exceptions |

### Key Differences

1. **Type System**: Python uses duck typing and optional type hints
2. **Memory Management**: Python has garbage collection vs Java's automatic memory management
3. **UNO Integration**: Slightly different syntax but same concepts
4. **Error Handling**: Python exceptions vs Java checked exceptions

## License

This code is ported from the OxygenOffice SmART extension and maintains the same licensing terms. Please refer to the original project for license details.

## Contributing

1. Implement UNO integration for actual LibreOffice connectivity
2. Add comprehensive test coverage
3. Improve error handling and validation
4. Add more diagram types from the original extension
5. Enhance documentation and examples

## Limitations

- **UNO Integration**: Requires manual implementation of LibreOffice UNO API calls
- **GUI Components**: Controller and GUI classes are stub implementations
- **Testing**: Limited testing without actual LibreOffice integration
- **Documentation**: Some advanced features may need additional documentation

## Related Projects

- **Original Java Extension**: OxygenOffice SmART LibreOffice extension
- **LibreOffice UNO**: LibreOffice Universal Network Objects API
- **Python UNO**: Python bindings for LibreOffice UNO API