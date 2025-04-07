"""
Excel Analyzer MCP Server
This application helps analyze Excel files using the Model Context Protocol.
"""

from mcp.server.fastmcp import FastMCP, Image
import pandas as pd
import matplotlib.pyplot as plt
import io
import os
from pathlib import Path

# Create the MCP server
mcp = FastMCP("Excel Analyzer")

# Define a tool to list Excel files in a directory
@mcp.tool()
def list_excel_files(directory_path: str = str(Path.home())) -> str:
    """
    List all Excel files in the specified directory.
    
    Args:
        directory_path: Path to the directory to scan for Excel files
                      (defaults to user's home directory)
    
    Returns:
        A string listing all Excel files found
    """
    try:
        # Convert to Path object for better path handling
        directory = Path(directory_path)
        
        # Check if directory exists
        if not directory.exists() or not directory.is_dir():
            return f"Error: Directory {directory_path} doesn't exist or is not a directory."
        
        # Find all Excel files
        excel_files = list(directory.glob("*.xlsx")) + list(directory.glob("*.xls"))
        
        if not excel_files:
            return f"No Excel files found in {directory_path}"
        
        # Format the result
        result = f"Found {len(excel_files)} Excel files in {directory_path}:\n\n"
        for i, file in enumerate(excel_files, 1):
            result += f"{i}. {file.name} - {file.absolute()}\n"
            
        return result
    
    except Exception as e:
        return f"Error listing Excel files: {str(e)}"

# Define a tool to load and summarize an Excel file
@mcp.tool()
def summarize_excel(file_path: str) -> str:
    """
    Load an Excel file and provide a summary of its content.
    
    Args:
        file_path: Full path to the Excel file
    
    Returns:
        A summary of the Excel file content
    """
    try:
        # Convert to Path object and check if file exists
        file = Path(file_path)
        if not file.exists():
            return f"Error: File {file_path} doesn't exist."
        
        # Load the Excel file
        df = pd.read_excel(file_path)
        
        # Generate summary
        summary = f"Excel File Analysis: {file.name}\n"
        summary += f"=" * 50 + "\n\n"
        summary += f"File Path: {file.absolute()}\n"
        summary += f"Total Sheets: {len(pd.ExcelFile(file_path).sheet_names)}\n"
        summary += f"Sheet Names: {', '.join(pd.ExcelFile(file_path).sheet_names)}\n\n"
        
        # Default to the first sheet
        summary += f"Analysis of first sheet:\n"
        summary += f"Total Rows: {len(df)}\n"
        summary += f"Total Columns: {len(df.columns)}\n\n"
        
        # Column information
        summary += "Column Information:\n"
        for col in df.columns:
            data_type = df[col].dtype
            non_null = df[col].count()
            null_percentage = (len(df) - non_null) / len(df) * 100 if len(df) > 0 else 0
            summary += f"- {col} (Type: {data_type}): {non_null} non-null values ({null_percentage:.1f}% missing)\n"
        
        # Data preview
        summary += "\nData Preview (First 5 rows):\n"
        summary += df.head(5).to_string() + "\n\n"
        
        # Numerical statistics if available
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            summary += "Statistical Summary of Numerical Columns:\n"
            summary += df[numeric_cols].describe().to_string() + "\n"
        
        return summary
    
    except Exception as e:
        return f"Error analyzing Excel file: {str(e)}"

# Define a tool to answer questions about Excel data
@mcp.tool()
def query_excel(file_path: str, query: str, sheet_name: str = None) -> str:
    """
    Answer questions about the Excel data using Pandas operations.
    
    Args:
        file_path: Full path to the Excel file
        query: Natural language query about the data
        sheet_name: Optional sheet name to analyze (defaults to first sheet)
    
    Returns:
        Answer to the query based on Excel data
    """
    try:
        # Load the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Process different types of queries
        query_lower = query.lower()
        
        # Handle count queries
        if "how many" in query_lower or "count" in query_lower:
            return f"Total rows: {len(df)}\nTotal columns: {len(df.columns)}"
        
        # Handle average/mean queries
        elif "average" in query_lower or "mean" in query_lower:
            # Try to identify which column
            numeric_cols = df.select_dtypes(include=['number']).columns
            results = []
            
            for col in numeric_cols:
                if col.lower() in query_lower:
                    avg = df[col].mean()
                    results.append(f"The average of {col} is {avg:.2f}")
            
            if results:
                return "\n".join(results)
            else:
                # If no specific column mentioned, return all averages
                return f"Average values for all numeric columns:\n{df[numeric_cols].mean().to_string()}"
        
        # Handle sum queries
        elif "sum" in query_lower or "total of" in query_lower:
            numeric_cols = df.select_dtypes(include=['number']).columns
            results = []
            
            for col in numeric_cols:
                if col.lower() in query_lower:
                    total = df[col].sum()
                    results.append(f"The sum of {col} is {total:.2f}")
            
            if results:
                return "\n".join(results)
            else:
                return f"Sum values for all numeric columns:\n{df[numeric_cols].sum().to_string()}"
        
        # Handle max/min queries
        elif "maximum" in query_lower or "max" in query_lower:
            numeric_cols = df.select_dtypes(include=['number']).columns
            results = []
            
            for col in numeric_cols:
                if col.lower() in query_lower:
                    max_val = df[col].max()
                    results.append(f"The maximum value of {col} is {max_val}")
            
            if results:
                return "\n".join(results)
            else:
                return f"Maximum values for all numeric columns:\n{df[numeric_cols].max().to_string()}"
        
        elif "minimum" in query_lower or "min" in query_lower:
            numeric_cols = df.select_dtypes(include=['number']).columns
            results = []
            
            for col in numeric_cols:
                if col.lower() in query_lower:
                    min_val = df[col].min()
                    results.append(f"The minimum value of {col} is {min_val}")
            
            if results:
                return "\n".join(results)
            else:
                return f"Minimum values for all numeric columns:\n{df[numeric_cols].min().to_string()}"
        
        # Default response - basic information
        else:
            return (f"Your query: '{query}' couldn't be specifically interpreted.\n\n"
                   f"Basic information about the data:\n"
                   f"- Rows: {len(df)}\n"
                   f"- Columns: {', '.join(df.columns)}\n\n"
                   f"Try asking about counts, averages, sums, maximums, or minimums of specific columns.")
    
    except Exception as e:
        return f"Error querying Excel file: {str(e)}"

# Define a tool to create charts from Excel data
@mcp.tool()
def create_chart(file_path: str, x_column: str, y_column: str, chart_type: str = "bar") -> Image:
    """
    Create a chart from Excel data.
    
    Args:
        file_path: Full path to the Excel file
        x_column: Column to use for x-axis
        y_column: Column to use for y-axis
        chart_type: Type of chart (bar, line, scatter, pie)
    
    Returns:
        An image of the chart
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)
        
        # Check if columns exist
        if x_column not in df.columns:
            return f"Error: Column '{x_column}' not found in the Excel file."
        if y_column not in df.columns:
            return f"Error: Column '{y_column}' not found in the Excel file."
        
        # Create figure and axis
        plt.figure(figsize=(10, 6))
        
        # Create the appropriate chart type
        if chart_type.lower() == "bar":
            df.plot(kind="bar", x=x_column, y=y_column)
        elif chart_type.lower() == "line":
            df.plot(kind="line", x=x_column, y=y_column)
        elif chart_type.lower() == "scatter":
            df.plot(kind="scatter", x=x_column, y=y_column)
        elif chart_type.lower() == "pie":
            # For pie charts, we typically need aggregated data
            df.groupby(x_column)[y_column].sum().plot(kind="pie")
        else:
            # Default to bar
            df.plot(kind="bar", x=x_column, y=y_column)
        
        # Add labels and title
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.title(f"{chart_type.capitalize()} Chart: {y_column} by {x_column}")
        plt.tight_layout()
        
        # Convert plot to image
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        
        # Return as MCP Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        return f"Error creating chart: {str(e)}"

# Run the server when script is executed
if __name__ == "__main__":
    print("Starting Excel Analyzer MCP Server...")
    mcp.run()