# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is the **银图PMC智能分析平台** (Yingtu PMC Intelligent Analysis Platform) - a comprehensive production planning and material control system for a hair dryer manufacturing factory. The system provides data-driven decision support for procurement, production scheduling, and financial analysis.

## Architecture & System Components

### Core Analysis Systems
- **silverPlan_analysis.py** (Primary Engine): Comprehensive PMC analysis system using LEFT JOIN architecture with ComprehensivePMCAnalyzer class
- **精准供应商物料分析系统.py**: Legacy supplier material matching with multi-supplier selection logic
- **精准订单物料分析系统.py**: Legacy order-material analysis with ROI calculations  
- **streamlit_dashboard.py**: Interactive web-based PMC analysis dashboard with file upload and real-time visualization

### Data Sources & Flow
The system integrates data from multiple Excel files with flexible input methods:
- **Order Data**: `order-amt-89.xlsx` (国内), `order-amt-89-c.xlsx` (柬埔寨) - 4 worksheets total (8月/9月 domestic, 8月-柬/9月-柬 Cambodia)
- **Material Shortage**: `mat_owe_pso.xlsx` - Material shortage tracking
- **Supplier Data**: `supplier.xlsx` - Supplier pricing and information  
- **Inventory**: `inventory_list.xlsx` - Current inventory with pricing

### Upload & Processing Architecture
- **File Upload Interface**: Streamlit-based drag-and-drop for all 5 required Excel files
- **Temporary Processing**: Files saved to temp directory during analysis
- **Real-time Analysis**: Automatic execution of silverPlan_analysis.py on uploaded data
- **Dynamic Report Generation**: Timestamped Excel reports with automatic dashboard refresh

### Key Business Logic
- **Hybrid Analysis Approach**: Combines precise matching (order→shortage→supplier) with estimation for unmatched orders
- **Supplier Selection Algorithm**: Multi-criteria scoring based on latest modification date (40%), lowest price (35%), and supplier stability (25%)
- **ROI Optimization**: Investment return calculations to guide procurement decisions
- **Multi-currency Support**: RMB, USD, HKD, EUR with configurable exchange rates

## Common Development Commands

### Environment Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Alternative: Install as package with entry points
pip install -e .
```

### Running the Analysis Systems
```bash
# Primary comprehensive analysis system (recommended)
python silverPlan_analysis.py

# Legacy individual analysis systems
python 精准供应商物料分析系统.py
python 精准订单物料分析系统.py

# Launch interactive dashboard with upload functionality
streamlit run streamlit_dashboard.py
```

### Development Tools
```bash
# Code formatting
black *.py

# Type checking
mypy streamlit_dashboard.py silverPlan_analysis.py

# Code quality analysis
pylint *.py

# Run tests (if available)
pytest
```

### Dashboard Deployment
```bash
# Local development server
streamlit run streamlit_dashboard.py --server.port 8501

# Production deployment with file watcher
streamlit run streamlit_dashboard.py --server.fileWatcherType none --server.port 8501
```

## Key File Locations

### Python Analysis Scripts
- `silverPlan_analysis.py` - **Primary comprehensive analysis engine** (ComprehensivePMCAnalyzer class)
- `streamlit_dashboard.py` - Web dashboard with file upload, real-time processing, and visualization
- `精准供应商物料分析系统.py` - Legacy supplier analysis engine
- `精准订单物料分析系统.py` - Legacy order-focused analysis
- `生成回款分析报告.py` - Payment return analysis
- `订单物料汇总分析.py` - Order material summarization
- `缺料清单分析报告.py` - Material shortage analysis
- `merge_order_files.py` - Utility for merging order files
- `setup.py` - Package installation configuration

### Configuration & Dependencies
- `requirements.txt` - Python dependencies with specific versions for stability
- `数据源.md` - Data source specifications and file format requirements
- `output field.json` - Expected output field structure and mapping

### Data Processing Architecture
- Files can be uploaded via Streamlit interface or placed in root directory
- Generated reports save with timestamps (format: `银图PMC综合物料分析报告_YYYYMMDD_HHMMSS.xlsx`)
- Multi-sheet Excel output: '综合物料分析明细' (main) + '汇总统计' (summary)
- Automatic report discovery by timestamp in dashboard

## System Features

### Dashboard Capabilities (`streamlit_dashboard.py`)
- **File Upload Interface**: Drag-and-drop upload for 5 required Excel files with progress tracking
- **Real-time Processing**: Automatic analysis execution with progress indicators and status updates
- **KPI Monitoring**: Total shortage amount, order count, supplier count, emergency purchases, ROI analysis
- **Interactive Filtering**: By month, date range, ROI thresholds, data completeness, supplier selection
- **Multi-view Tabs**: Management overview, procurement lists (order/supplier views), deep analysis
- **Export Functions**: CSV export with proper encoding (GBK for Excel compatibility)
- **Expandable Order Views**: Detailed material breakdown with supplier tabs and ROI calculations

### Core Analysis Engine (`silverPlan_analysis.py`)
- **LEFT JOIN Architecture**: Orders as primary table, ensuring all orders appear in results
- **Comprehensive Data Integration**: Orders + shortage + inventory + supplier data with conservative filling
- **Lowest Price Supplier Selection**: Multi-criteria algorithm with price optimization
- **ROI Calculations**: Investment return analysis (order amount ÷ shortage amount) with "无需投入" for zero shortage
- **Currency Conversion**: Multi-currency support (RMB, USD, HKD, EUR) with configurable rates
- **Data Quality Tracking**: Completeness markers and data filling audit trails

## Development Guidelines

### Data Handling
- Always handle missing data gracefully with `.fillna()` and error handling
- Use `pd.to_numeric(errors='coerce')` for numeric conversions
- Maintain string data types for order numbers and IDs to preserve leading zeros
- Implement proper encoding (GBK/UTF-8) for Chinese text compatibility

### Excel Integration
- Use `openpyxl` engine for modern Excel files
- Support both `.xlsx` and `.xls` formats where needed
- Generate multi-sheet reports with descriptive sheet names
- Include summary statistics sheets for management reporting

### Performance Considerations
- Use `@st.cache_data` for expensive data loading operations
- Implement progress tracking for large dataset processing and file uploads
- Batch process materials in groups of 100 for user feedback
- Optimize DataFrame operations with vectorized calculations
- Temporary file cleanup after analysis to prevent disk space issues

### File Upload & Integration Patterns
- Use `tempfile.mkdtemp()` for secure temporary file storage during processing
- Implement custom data loading methods by replacing `analyzer.load_all_data` for uploaded files
- Handle file validation and format checking before analysis execution
- Provide clear user feedback during upload and processing stages
- Automatic cache clearing and page refresh after successful analysis

## Business Context

This system serves a hair dryer manufacturing company that needs to:
- Optimize cash flow and return on investment (target 2.8-3.0x ROI)
- Reduce material shortage rates by 30-40%  
- Enable minute-level production scheduling decisions
- Provide 5-minute anomaly detection and response

The target ROI improvement from current 2.0-2.5x to 2.8-3.0x represents a 20% efficiency gain, which drives the algorithmic prioritization throughout the system.