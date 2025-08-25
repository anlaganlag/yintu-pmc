# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is the **银图PMC智能分析平台** (Yingtu PMC Intelligent Analysis Platform) - a comprehensive production planning and material control system for a hair dryer manufacturing factory. The system provides data-driven decision support for procurement, production scheduling, and financial analysis.

## Architecture & System Components

### Core Analysis Systems
- **精准供应商物料分析系统** (`精准供应商物料分析系统.py`): Supplier material matching with multi-supplier selection logic
- **精准订单物料分析系统** (`精准订单物料分析系统.py`): Order-material analysis with ROI calculations  
- **Streamlit Dashboard** (`streamlit_dashboard.py`): Interactive web-based PMC analysis dashboard with visualization

### Data Sources & Flow
The system integrates data from multiple Excel files:
- **Order Data**: `(全部)8月9月订单.xlsx` - Monthly order sheets (August/September)
- **Material Shortage**: `mat_owe_pso.xlsx`, `owe-mat.xls` - Material shortage tracking
- **Supplier Data**: `supplier.xlsx` - Supplier pricing and information
- **Inventory**: `inventory_list.xlsx` - Current inventory with pricing

### Key Business Logic
- **Hybrid Analysis Approach**: Combines precise matching (order→shortage→supplier) with estimation for unmatched orders
- **Supplier Selection Algorithm**: Multi-criteria scoring based on latest modification date (40%), lowest price (35%), and supplier stability (25%)
- **ROI Optimization**: Investment return calculations to guide procurement decisions
- **Multi-currency Support**: RMB, USD, HKD, EUR with configurable exchange rates

## Common Development Commands

### Running the Analysis Systems
```bash
# Run supplier material analysis
python 精准供应商物料分析系统.py

# Run order material analysis  
python 精准订单物料分析系统.py

# Launch interactive dashboard
streamlit run streamlit_dashboard.py
```

### Running the Dashboard
The Streamlit dashboard provides real-time visualization and can be launched with:
```bash
streamlit run streamlit_dashboard.py --server.port 8501
```

## Key File Locations

### Python Analysis Scripts
- `精准供应商物料分析系统.py` - Main supplier analysis engine
- `精准订单物料分析系统.py` - Order-focused analysis
- `streamlit_dashboard.py` - Web dashboard with interactive features
- `生成回款分析报告.py` - Payment return analysis
- `订单物料汇总分析.py` - Order material summarization
- `缺料清单分析报告.py` - Material shortage analysis

### Documentation
- `银图PMC系统PRD.md` - Product requirements document
- `银图开发计划.md` - Development roadmap
- `汇总表需求.md` - Summary table requirements
- `排期系统.md` - Production scheduling system specs

### Data Processing
- All scripts expect Excel data files in the root directory
- Generated reports save to `D:\yingtu-PMC\` with timestamps
- Multi-sheet Excel output with detailed analysis tables

## System Features

### Dashboard Capabilities (`streamlit_dashboard.py`)
- **Real-time KPI Monitoring**: Total shortage amount, order count, supplier count, emergency purchases
- **Investment ROI Analysis**: Return on investment calculations with high-return project identification
- **Interactive Filtering**: By month, date range, ROI thresholds, data completeness
- **Export Functions**: CSV export with proper encoding (GBK for Excel compatibility)
- **Multi-view Organization**: Management overview, procurement lists, deep analysis tabs

### Analysis Engine Features
- **Precise Matching**: Direct order-to-supplier mapping when shortage data exists
- **Intelligent Estimation**: ML-based cost estimation for orders without precise shortage data
- **Multi-supplier Optimization**: Backup supplier options with price comparison
- **Currency Conversion**: Automatic RMB conversion with configurable rates
- **Data Integrity Tracking**: Completeness markers for audit trails

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
- Implement progress tracking for large dataset processing
- Batch process materials in groups of 100 for user feedback
- Optimize DataFrame operations with vectorized calculations

## Business Context

This system serves a hair dryer manufacturing company that needs to:
- Optimize cash flow and return on investment (target 2.8-3.0x ROI)
- Reduce material shortage rates by 30-40%  
- Enable minute-level production scheduling decisions
- Provide 5-minute anomaly detection and response

The target ROI improvement from current 2.0-2.5x to 2.8-3.0x represents a 20% efficiency gain, which drives the algorithmic prioritization throughout the system.