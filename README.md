# Excel Analyzer with Model Context Protocol (MCP)

A desktop application that demonstrates the power of the Model Context Protocol (MCP) for analyzing Excel files through AI assistance. This project showcases how MCP enables AI systems to securely interact with local files while maintaining clear data boundaries.

## What is Model Context Protocol (MCP)?

Model Context Protocol (MCP) is an open-source protocol from Anthropic that establishes a standardized way for AI models to interact with contextual information sources like files, databases, and APIs. MCP follows a client-host-server architecture that enables:

- **Secure Data Access**: AI models can interact with files without directly accessing them
- **Standardized Interface**: Consistent API for accessing different types of data sources
- **Clear Security Boundaries**: Control over what data is accessible to the AI

In our application, MCP serves as the bridge between the Excel files on your computer and the AI assistant, allowing for intelligent analysis while maintaining data security.

## Key Features

### File Discovery & Analysis through MCP
- Scan directories for Excel files using MCP file operations
- Select files for analysis through the MCP interface
- View file structure and metadata through MCP-managed access

### Data Analysis with MCP Tools
- Comprehensive file summaries using MCP-enabled data extraction
- Statistical analysis of Excel data through MCP queries
- Flexible data filtering with pandas-compatible syntax

### AI-Powered Insights
- Natural language interface for Excel data analysis
- Integration with OpenAI API, connected through MCP
- Contextual awareness of file structure and content

## Technical Implementation

### MCP Architecture in Excel Analyzer

Our application implements the MCP architecture with these components:

1. **MCP Server**: Provides Excel file access tools and resources
2. **Streamlit Interface**: Acts as the host application and user interface
3. **MCP Client**: Connects the AI model (via OpenAI API) to the Excel resources

This architecture ensures that:
- Excel files remain secure on your local machine
- AI analysis has controlled access to file data
- The user maintains full control over which files are analyzed

### Technologies Used

- **Python**: Core programming language
- **MCP SDK**: For implementing the Model Context Protocol
- **Streamlit**: Web application framework
- **Pandas & Matplotlib**: Data processing and visualization
- **OpenAI API**: AI-powered analysis

## Installation and Setup

### Prerequisites

- Python 3.9 or higher
- OpenAI API key

### Installation Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-analyzer-mcp.git
   cd excel-analyzer-mcp
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up your OpenAI API key in a `.env` file:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

### Running the Application

Launch the application with:

```bash
python -m src.mcp_excel_analyzer --app
```

## How It Works

1. **MCP Discovery**: The application scans directories for Excel files using MCP file system tools
2. **Secure Analysis**: When you select a file, MCP provides controlled access for analysis
3. **AI Integration**: Questions about the data are processed by OpenAI through the MCP interface
4. **Visualization**: Charts and data views are generated using data accessed via MCP

## Why Use MCP?

MCP provides several advantages for applications like our Excel Analyzer:

1. **Security**: Files remain on your system and are accessed through controlled interfaces
2. **Standardization**: A consistent protocol for accessing different data sources
3. **Separation of Concerns**: Clean architecture separating UI, data access, and AI logic
4. **Future Extensibility**: The same approach can be extended to other data sources

This proof-of-concept demonstrates how MCP can provide a secure, standardized way for AI applications to interact with local resources like Excel files while maintaining proper boundaries and controls.