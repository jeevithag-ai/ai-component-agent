# AI Component Agent - Basic Starter

# Sample component data (mock for now)
components = [
    {"name": "Capacitor", "value": "10uF", "price": 0.10},
    {"name": "Resistor", "value": "1kΩ", "price": 0.05},
    {"name": "Inductor", "value": "1mH", "price": 0.20}
]

# Function to calculate total BOM cost
def calculate_bom_cost(components):
    total = 0
    for comp in components:
        total += comp["price"]
    return total

# Run tool
if __name__ == "__main__":
    total_cost = calculate_bom_cost(components)
    print(f"Total BOM Cost: ${total_cost}")
