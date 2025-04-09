# Intelligent Inventory Planning Assistant

## Overview

This project builds an intelligent assistant capable of analyzing complex Excel-based inventory planning workbooks and answering business questions through natural language. By combining the Model Context Protocol (MCP) with direct LLM API calls, we create a powerful hybrid system that leverages the strengths of both approaches.

The system allows users to ask questions like "How can I increase my profits by a million dollars?" and receive thoughtful, data-driven answers based on their inventory planning workbooks.

## System Architecture

Our solution employs a hybrid architecture combining local Excel processing via MCP with specialized reasoning through direct LLM API calls:

```
┌──────────────────────────────────────────────────────────────────────────┐
│                          User Interface Layer                            │
└───────────────────────────────────┬──────────────────────────────────────┘
                                   │
                                   ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                        Orchestration Controller                          │
└───────┬─────────────────────┬──────────────────────┬─────────────────────┘
        │                     │                      │
        ▼                     ▼                      ▼
┌───────────────┐    ┌─────────────────┐    ┌─────────────────────┐
│   Question    │    │     Excel       │    │    Calculation &     │
│   Analysis    │    │   Processing    │    │      Reasoning       │
│  (LLM API)    │    │     (MCP)       │    │      (LLM API)       │
└───────────────┘    └─────────────────┘    └─────────────────────┘
        │                     │                      │
        └─────────────────────┼──────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                         Answer Generation                                │
└──────────────────────────────────────────────────────────────────────────┘
```

## Key Components

### 1. Question Analysis (LLM API)
Processes natural language questions using advanced reasoning models like GPT-4o to:
- Identify the business objective
- Determine required data points
- Plan calculation approaches
- Structure the analysis workflow

### 2. Excel Processing (MCP)
Handles all interactions with Excel workbooks through MCP tools:
- Scans workbook structure and sheet metadata
- Extracts relevant data from identified sheets
- Transforms raw data into structured formats
- Executes calculations based on analysis plans

```
┌────────────────────────────────────────────────────────────┐
│                    Excel Processing                        │
│                                                            │
│  ┌─────────────┐   ┌──────────────┐   ┌─────────────────┐  │
│  │  Workbook   │   │ Sheet Data   │   │ Data            │  │
│  │  Structure  ├──►│ Extraction   ├──►│ Transformation  │  │
│  │  Analysis   │   │              │   │                 │  │
│  └─────────────┘   └──────────────┘   └─────────────────┘  │
│                                                            │
└────────────────────────────────────────────────────────────┘
```

### 3. Calculation & Reasoning (LLM API)
Provides sophisticated reasoning capabilities:
- Develops calculation strategies for business questions
- Determines appropriate analytical approaches
- Evaluates data quality issues
- Assesses confidence in generated answers
- Provides explanation and business recommendations

### 4. Orchestration Controller
Manages the flow between components:
- Initiates the appropriate sequence of operations
- Passes data between LLM API and MCP components
- Tracks the state of the analysis process
- Handles errors and retries as needed

## Data Flow

The system processes data through the following flow:

```
    ┌─────────────┐
    │ User Query  │
    └──────┬──────┘
           │
           ▼
┌─────────────────────┐   ┌─────────────────────┐
│  Question Analysis  │   │ Workbook Structure  │
│     (LLM API)       ├──►│  Analysis (MCP)     │
└─────────┬───────────┘   └──────────┬──────────┘
          │                          │
          └──────────┬───────────────┘
                     │
                     ▼
         ┌─────────────────────┐
         │  Sheet Relevance    │
         │ Determination (LLM) │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │    Data Extraction  │
         │       (MCP)         │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │   Data Quality      │
         │  Assessment (LLM)   │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │ Data Transformation │
         │       (MCP)         │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │  Analysis Strategy  │
         │     (LLM API)       │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │ Execute Calculations│
         │       (MCP)         │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │ Confidence Assessment│
         │     (LLM API)       │
         └──────────┬──────────┘
                    │
                    ▼
         ┌─────────────────────┐
         │  Final Answer with  │
         │    Explanations     │
         └─────────────────────┘
```

## Implementation Details

### MCP Tools Extension

We extend the existing `excel_analyzer.py` with additional MCP tools:

```python
@mcp.tool()
def analyze_workbook_structure(file_path: str) -> dict:
    """Analyze overall structure of an Excel workbook"""
    # Implementation details...

@mcp.tool()
def extract_sheet_data(file_path: str, sheet_name: str, columns: list = None) -> dict:
    """Extract data from a specific sheet with optional column filtering"""
    # Implementation details...

@mcp.tool()
def transform_data(data: dict, transformations: dict) -> dict:
    """Apply transformations to extracted data"""
    # Implementation details...

@mcp.tool()
def execute_calculation(data: dict, calculation_instructions: dict) -> dict:
    """Execute calculations based on provided instructions"""
    # Implementation details...

@mcp.tool()
def answer_inventory_question(file_path: str, question: str) -> dict:
    """Complete pipeline to answer inventory questions"""
    # Main orchestration implementation...
```

### LLM API Integration

We integrate direct LLM API calls for specialized reasoning tasks:

```python
def analyze_question(question: str) -> dict:
    """Use reasoning model to analyze inventory questions"""
    # Implementation using OpenAI API or similar

def determine_relevant_sheets(workbook_structure: dict, question_analysis: dict) -> list:
    """Determine which sheets contain relevant data"""
    # Implementation using reasoning model

def assess_data_quality(extracted_data: dict) -> dict:
    """Evaluate data quality and recommend transformations"""
    # Implementation using reasoning model

def develop_calculation_strategy(question: str, available_data: dict) -> dict:
    """Develop strategy for answering inventory questions"""
    # Implementation using reasoning model

def evaluate_answer(question: str, answer: dict, data_used: dict) -> dict:
    """Evaluate confidence in generated answer"""
    # Implementation using reasoning model
```

### Orchestration

The main orchestration function ties everything together:

```python
@mcp.tool()
def answer_inventory_question(file_path: str, question: str) -> dict:
    """
    End-to-end process to answer an inventory planning question.
    
    Args:
        file_path: Path to the inventory planning workbook
        question: User's natural language question
        
    Returns:
        Comprehensive answer with confidence assessment
    """
    # 1. Analyze question
    question_analysis = analyze_question(question)
    
    # 2. Scan workbook structure
    workbook_structure = analyze_workbook_structure(file_path)
    
    # 3. Determine relevant sheets
    relevant_sheets = determine_relevant_sheets(workbook_structure, question_analysis)
    
    # 4. Extract and process data
    extracted_data = {}
    for sheet_info in relevant_sheets:
        sheet_data = extract_sheet_data(
            file_path, 
            sheet_info["sheet_name"], 
            sheet_info.get("columns")
        )
        extracted_data[sheet_info["sheet_name"]] = sheet_data
    
    # 5. Assess data quality
    quality_assessment = assess_data_quality(extracted_data)
    
    # 6. Transform data if needed
    transformed_data = transform_data(extracted_data, quality_assessment["transformations"])
    
    # 7. Develop calculation strategy
    calculation_strategy = develop_calculation_strategy(question, transformed_data)
    
    # 8. Execute calculations
    calculation_results = execute_calculation(transformed_data, calculation_strategy)
    
    # 9. Evaluate answer and confidence
    evaluation = evaluate_answer(question, calculation_results, transformed_data)
    
    # 10. Format and return comprehensive response
    return {
        "question": question,
        "answer": calculation_results["answer"],
        "explanation": calculation_results["explanation"],
        "confidence": evaluation["confidence"],
        "confidence_explanation": evaluation["explanation"],
        "limitations": evaluation["limitations"],
        "recommendations": evaluation["recommendations"]
    }
```

## Usage Examples

### Simple Profit Analysis Question

```
User: "How can I increase my profits by a million dollars?"

System: [Process overview]
1. Analyzes question → identifies profit focus
2. Scans workbook → identifies sales, costs, pricing sheets 
3. Extracts relevant data → sales figures, costs, margins
4. Performs calculations → current profits, improvement scenarios
5. Evaluates options → pricing changes, cost reductions, volume increases

Response:
"Based on your current inventory planning data, you could increase profits 
by $1 million through a combination of strategies:

1. Price Optimization: Increasing prices on high-volume, low-margin products 
   (particularly Widget X and Gadget Y) by 2.3% would generate approximately 
   $425,000 in additional profit with minimal impact on volume.

2. Inventory Carrying Cost Reduction: Your current inventory levels for 
   seasonal products are 22% above optimal levels. Reducing this excess 
   inventory could save approximately $315,000 in carrying costs.

3. Supplier Consolidation: Analysis of your current supplier pricing shows 
   a 5% potential saving through consolidating orders with suppliers A and B, 
   generating approximately $260,000 in cost savings.

Confidence: High (85%)
This analysis is based on complete sales and cost data from the past 12 months 
with clear margin calculations. The main limitation is uncertainty about 
elasticity of demand for price increases, which may require testing before 
full implementation."
```

### Complex Inventory Optimization Question

```
User: "What would happen if we reduced our safety stock by 15%?"

System: [Process overview]
1. Analyzes question → identifies inventory risk focus
2. Scans workbook → identifies inventory, sales, stockout sheets
3. Extracts relevant data → current safety stock, historical demand, stockout costs
4. Performs calculations → service level impacts, cost savings, risk assessment
5. Evaluates tradeoffs → cost savings vs. service level impacts

Response:
"Reducing safety stock by 15% would have the following impacts:

1. Cost Savings: Approximately $342,500 annual reduction in inventory 
   carrying costs based on your current weighted average cost of capital (8.2%).

2. Service Level Impact: Your overall service level would decrease from 
   98.3% to approximately 94.7%, based on historical demand patterns.

3. Stockout Risk: The probability of stockouts would increase for 13 SKUs, 
   with Products X, Y, and Z facing the highest risk (>25% increased risk).

4. Financial Impact: The estimated cost of potential stockouts ($295,000) 
   nearly offsets the carrying cost savings, resulting in a net benefit 
   of only $47,500.

Confidence: Medium (72%)
This analysis accounts for historical demand patterns but cannot predict 
future demand spikes. The calculation assumes your historical stockout 
costs are accurate, which could be refined with more detailed data."
```

## Development Roadmap

We recommend implementing this system in progressive phases:

### Phase 1: Core Excel Analysis Foundation
- Implement basic MCP tools for Excel interaction
- Build question analysis for simple inventory questions
- Create basic data extraction and transformation capabilities

### Phase 2: Enhanced Data Processing
- Add data quality assessment and transformation
- Implement business logic extraction from formulas
- Create basic calculation capabilities for common questions

### Phase 3: Advanced Reasoning Integration
- Integrate LLM reasoning for question analysis
- Add calculation strategy development
- Implement confidence assessment

### Phase 4: Comprehensive Optimization
- Add multi-sheet analysis capabilities
- Implement business recommendation generation
- Create visualization tools for complex analyses

## Considerations and Limitations

### Data Privacy and Security
- All processing occurs locally through MCP
- No Excel data is transmitted to LLM services except metadata
- Consider implementing additional controls for sensitive data

### Excel Complexity
- Complex Excel files with extensive macros may require specialized handling
- Linked workbooks or external data sources need additional implementation
- Handling proprietary formulas may require domain-specific knowledge

### Performance Considerations
- Large workbooks may require optimized processing strategies
- Consider implementing caching for repeated analyses
- Balance between local processing and API calls for optimal performance

### Business Logic Accuracy
- Inventory planning involves domain-specific business rules
- Consider implementing validation checks for critical calculations
- Allow for user feedback to refine calculation strategies

## Getting Started

### Prerequisites
- Python 3.8+
- MCP SDK
- Access to OpenAI API or other LLM service

### Installation
```bash
# Clone the repository
git clone

# Install dependencies
pip install -r requirements.txt

# Configure API keys
cp .env.example .env
# Edit .env with your API keys
```

### Configuration
Edit the `config.json` file to specify:
- LLM API endpoints and parameters
- Default file locations
- Logging preferences

## Conclusion

This Intelligent Inventory Planning Assistant demonstrates the power of combining MCP's local data processing capabilities with the advanced reasoning of LLM APIs. By creating a hybrid architecture, we leverage the strengths of both approaches to provide users with insightful, data-driven answers to complex inventory planning questions.

The system respects data privacy by processing Excel files locally while still delivering the sophisticated reasoning capabilities of advanced language models. Through this approach, we create a valuable tool that can help businesses optimize their inventory planning processes and improve profitability.