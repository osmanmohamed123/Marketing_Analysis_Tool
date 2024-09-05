import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

data = None  # Global variable to store loaded data

def load_data(file_path):
    """Load data from a CSV or Excel file."""
    global data
    try:
        if file_path.endswith('.csv'):
            data = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            data = pd.read_excel(file_path)
        else:
            print("Unsupported file format. Please upload a CSV or Excel file.")
        print("Data loaded successfully.")
    except Exception as e:
        print(f"Error loading data: {e}")

def analyze_performance():
    """Analyze performance metrics from the loaded data."""
    if data is None:
        print("No data loaded. Please load data first using the load_data command.")
        return

    try:
        # Example metric calculations
        open_rate = (data['emails_opened'] / data['emails_sent']) * 100
        click_through_rate = (data['clicks'] / data['emails_opened']) * 100
        conversion_rate = (data['conversions'] / data['emails_sent']) * 100

        # Adding metrics to the data frame for future use
        data['open_rate'] = open_rate
        data['click_through_rate'] = click_through_rate
        data['conversion_rate'] = conversion_rate

        print("Performance analysis complete.")
    except KeyError as e:
        print(f"Error in analyzing performance: Missing column {e}")
    except Exception as e:
        print(f"Error in analyzing performance: {e}")

def generate_report(report_name):
    """Generate a text-based report summarizing key metrics."""
    if data is None:
        print("No data loaded. Please load data first using the load_data command.")
        return
    
    try:
        with open(report_name, 'w') as f:
            f.write("Marketing Campaign Performance Report\n")
            f.write("=" * 40 + "\n")
            f.write(data[['open_rate', 'click_through_rate', 'conversion_rate']].describe().to_string())
        print(f"Report generated successfully: {report_name}")
    except Exception as e:
        print(f"Error generating report: {e}")

def generate_visualization():
    """Generate visualizations of the marketing data."""
    if data is None:
        print("No data loaded. Please load data first using the load_data command.")
        return
    
    try:
        sns.barplot(x=data.index, y='open_rate', data=data)
        plt.title('Open Rate by Campaign')
        plt.savefig('open_rate.png')
        plt.show()

        sns.lineplot(x=data.index, y='click_through_rate', data=data)
        plt.title('Click-Through Rate Over Time')
        plt.savefig('click_through_rate.png')
        plt.show()

        print("Visualizations generated and saved successfully.")
    except Exception as e:
        print(f"Error generating visualizations: {e}")


file_path = 'C:\\Users\\Acer\\OneDrive\\Documents\\MyPythonProject\\campaign email.csv'

# Step 1: Load the data
load_data(file_path)

# Step 2: Analyze performance metrics
analyze_performance()

# Step 3: Generate a report
report_name = 'performance_report.txt'
generate_report(report_name) 

# Step 4: Generate visualizations
generate_visualization()