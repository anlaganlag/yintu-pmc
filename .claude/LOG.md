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

## Usage Guidelines for PMC Work

- **For Excel Issues:** Review Excel Processing Pattern and encoding lessons
- **For ROI Problems:** Check ROI Calculation Pattern and edge case handling
- **For Performance:** Apply Dashboard Performance Pattern and optimization techniques
- **For Data Integrity:** Always use LEFT JOIN patterns to preserve order completeness