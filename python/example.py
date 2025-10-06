"""
Simple usage example for the OxygenOffice SmART Python port
"""

from oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.orgchart import OrgChart
from oxygenoffice.extensions.smart.diagram.data_of_diagram import DataOfDiagram
from oxygenoffice.extensions.smart.controller import Controller, Gui


def create_simple_org_chart():
    """Create a simple organization chart"""
    print("Creating simple organization chart...")
    
    # Create controller and GUI instances (stub implementations)
    controller = Controller()
    gui = Gui()
    
    # Create organization chart
    org_chart = OrgChart(controller, gui, None)
    
    # Create hierarchical data
    data = DataOfDiagram()
    data.add(0, "CEO")                    # Level 0 (root)
    data.add(1, "Chief Technology Officer")  # Level 1
    data.add(1, "Chief Financial Officer")   # Level 1
    data.add(1, "Chief Marketing Officer")   # Level 1
    data.add(2, "Development Manager")       # Level 2 (under CTO)
    data.add(2, "QA Manager")               # Level 2 (under CTO)
    data.add(3, "Senior Developer")         # Level 3 (under Dev Manager)
    data.add(3, "Junior Developer")         # Level 3 (under Dev Manager)
    
    print(f"Created data with {data.size()} items")
    data.print_data()
    
    # Create the diagram (this calls stub methods)
    print("\nCreating diagram...")
    org_chart.create_diagram(data)
    
    print("Organization chart creation completed!")
    print("Note: This used stub implementations. For real LibreOffice integration,")
    print("see UNO_INTEGRATION_GUIDE.py for implementation details.")


def create_simple_numbered_chart():
    """Create a simple numbered organization chart"""
    print("\nCreating simple numbered organization chart...")
    
    controller = Controller()
    gui = Gui()
    org_chart = OrgChart(controller, gui, None)
    
    # Create a simple chart with 5 shapes (1 root + 4 children)
    print("Creating diagram with 5 shapes...")
    org_chart.create_diagram(5)
    
    print("Simple numbered chart creation completed!")


def demonstrate_tree_operations():
    """Demonstrate tree operations"""
    print("\nDemonstrating tree operations...")
    
    # Create sample data
    data = DataOfDiagram()
    data.add(0, "Root")
    data.add(1, "Child 1")
    data.add(1, "Child 2")
    data.add(2, "Grandchild 1")
    data.add(2, "Grandchild 2")
    
    print("Original data:")
    data.print_data()
    
    print(f"Has one first level data: {data.is_one_first_level_data()}")
    print(f"Data size: {data.size()}")
    print(f"Is empty: {data.is_empty()}")
    
    # Increase all levels
    print("\nAfter increasing levels:")
    data.increase_levels()
    data.print_data()


def demonstrate_color_schemes():
    """Demonstrate color scheme functionality"""
    print("\nDemonstrating color schemes...")
    
    from oxygenoffice.extensions.smart.diagram.scheme_definitions import SchemeDefinitions
    
    print("Available color schemes:")
    schemes = [
        ("Blue Scheme", SchemeDefinitions.BLUE_SCHEME),
        ("Red Scheme", SchemeDefinitions.RED_SCHEME),
        ("Green Scheme", SchemeDefinitions.GREEN_SCHEME),
        ("Fire Scheme", SchemeDefinitions.FIRE_SCHEME),
        ("Purple Scheme", SchemeDefinitions.PURPLE_SCHEME),
    ]
    
    for name, (light, dark) in schemes:
        print(f"  {name}: Light={hex(light)}, Dark={hex(dark)}")
    
    # Demonstrate gradient calculation
    print("\nGradient color examples:")
    base_color = 0xff0000  # Red
    for i in range(5):
        gradient_color = SchemeDefinitions.get_gradient_color(base_color, i, 5)
        print(f"  Step {i}: {hex(gradient_color)}")


if __name__ == "__main__":
    print("OxygenOffice SmART Python Port - Example Usage")
    print("=" * 50)
    
    # Run examples
    create_simple_org_chart()
    create_simple_numbered_chart()
    demonstrate_tree_operations()
    demonstrate_color_schemes()
    
    print("\n" + "=" * 50)
    print("Example completed!")
    print("\nNext steps:")
    print("1. Review UNO_INTEGRATION_GUIDE.py for LibreOffice integration")
    print("2. Implement real UNO API calls to replace stub methods")
    print("3. Test with actual LibreOffice documents")