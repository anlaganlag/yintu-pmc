# PMC Project Log

This log captures lessons learned and patterns specific to the PMC (Production Material Control) analysis system.

## Log Entry Types

### COMPLETED TASK Entries
```markdown
---
## COMPLETED TASK: [YYYY-MM-DD HH:MM:SS]
**Task File:** tasks/xxx_task_name.md
**Total Steps:** [number]
**Summary:** [list of completed actions]
**PMC Impact:** [manufacturing/analysis improvements]
---
```

### LESSON LEARNED Entries
```markdown
---
## LESSON LEARNED: [YYYY-MM-DD HH:MM:SS]
**Error Type:** [Data/Excel/ROI/Performance/etc.]
**Problem:** [what went wrong]
**Root Cause:** [underlying reason]
**Solution:** [what was changed]
**Prevention Rule:** [rule for future PMC work]
**Files Changed:** [list of files]
---
```

### PATTERN LEARNED Entries
```markdown
---
## PATTERN LEARNED: [YYYY-MM-DD HH:MM:SS]
**Task Type:** [Excel/Dashboard/ROI/Analysis/etc.]
**When Working On:** [PMC-specific scenario]
**Files That Always Need Changes:** [consistent file patterns]
**Common Steps:** [recurring PMC implementation steps]
**Business Impact:** [manufacturing efficiency gains]
---
```

## PMC-Specific Patterns

### Excel Processing Pattern
**When Working On:** Excel file upload and data integration
**Files That Always Need Changes:** [silverPlan_analysis.py, streamlit_dashboard.py, data validation modules]
**Common Steps:**
1. Validate Excel file format and encoding (GBK/UTF-8)
2. Preserve data integrity during LEFT JOIN operations
3. Handle multi-currency conversion accurately
4. Maintain order completeness (no lost orders)
5. Test with actual manufacturing data files

### ROI Calculation Pattern
**When Working On:** Investment return analysis and optimization
**Files That Always Need Changes:** [ROI calculation modules, dashboard KPI displays, reporting functions]
**Common Steps:**
1. Validate order amount and shortage amount data quality
2. Handle edge cases (zero shortage = "无需投入")
3. Apply correct currency conversion rates
4. Ensure calculation supports management decision-making
5. Test across different order types (PSO, RSO, MSO, TSO)

### Dashboard Performance Pattern
**When Working On:** Management dashboard and visualization updates
**Files That Always Need Changes:** [streamlit_dashboard.py, visualization components, export functions]
**Common Steps:**
1. Ensure 5-minute response time requirement met
2. Optimize for large dataset processing (6000+ order records)
3. Maintain interactive filtering and real-time updates
4. Test export functionality with proper encoding
5. Verify mobile/tablet accessibility for management use

---
## LESSON LEARNED: 2025-09-02 00:48:00
**Error Type:** Runtime/KeyError
**Problem:** KeyError: '物項編號_清理' during LEFT JOIN inventory analysis - code tried to access old column name after field name unification
**Root Cause:** Incomplete field name mapping during refactoring - created new standardized column names but debug code still referenced old column names
**Solution:** 
1. Fixed all column name references to use new standardized names ('物料编号_清理' instead of '物項編號_清理')
2. Added comprehensive column existence validation with fallback logic
3. Added validate_required_columns() function to prevent similar issues
4. Enhanced merge operations to handle missing optional columns gracefully
**Prevention Rule:** Always validate column existence before accessing DataFrame columns, especially after field name unification operations
**Files Changed:** silverPlan_analysis.py (lines 362-453: column validation, field name mapping, merge operations)
---

---
## COMPLETED TASK: 2025-09-02 00:35:00
**Task File:** 001_clarify_silverplan_processing_logic_refactor.md
**Total Steps:** 15
**Summary:** 
- Fixed Excel data uniqueness by adding 数量Pcs dimension to groupby logic
- Enhanced material code matching with standardization and field name unification  
- Improved system reliability with comprehensive error handling and data quality reporting
- Added material match statistics tracking and high-risk order identification
- Validated all changes with actual data (455 orders, 10K+ shortage records, 100K+ inventory items)
**PMC Impact:** 
- Eliminated investment ratio calculation errors from duplicate order counting
- Improved material code matching accuracy, reducing price lookup failures
- Enhanced management decision support with data quality insights and high-risk order alerts
- Maintained 100% backward compatibility while improving system reliability
---

## Usage Guidelines for PMC Work

- **For Excel Issues:** Review Excel Processing Pattern and encoding lessons
- **For ROI Problems:** Check ROI Calculation Pattern and edge case handling
- **For Performance:** Apply Dashboard Performance Pattern and optimization techniques
- **For Data Integrity:** Always use LEFT JOIN patterns to preserve order completeness