# Action: Debug and Fix Error

## OBJECTIVE:
Perform comprehensive error analysis and provide immediate fix, root cause resolution, and prevention measures.

**INPUT:** Error description as `$ARGUMENTS` (e.g., "ROI calculation returns negative values")

**SUCCESS CRITERIA:** Error is fixed, root cause identified, and lesson learned is logged to prevent recurrence.

---

### STEP 1: IMMEDIATE ERROR ANALYSIS

**GATHER CONTEXT:**
- [ ] Read error message/description from `$ARGUMENTS`
- [ ] Identify affected files and functions
- [ ] Reproduce error if possible
- [ ] Analyze recent code changes that might be related

**PMC-SPECIFIC DEBUGGING PRIORITIES:**
- [ ] Check data accuracy (manufacturing decisions depend on correct data)
- [ ] Verify Excel file processing integrity
- [ ] Validate ROI calculation logic
- [ ] Confirm currency conversion accuracy
- [ ] Check LEFT JOIN integrity (ensure no orders lost)

---

### STEP 2: THREE-LEVEL SOLUTION APPROACH

#### **Level 1: IMMEDIATE FIX**
Stop the error/crash right now:
- [ ] Identify minimal change to prevent error
- [ ] Apply immediate fix with clear comments
- [ ] Test immediate fix works

#### **Level 2: ROOT CAUSE FIX** 
Prevent this category of error systemically:
- [ ] Identify underlying cause
- [ ] Implement systemic solution
- [ ] Add validation/checks to prevent similar errors

#### **Level 3: PREVENTION LESSON**
Update practices to avoid similar issues:
- [ ] Document lesson learned
- [ ] Update code patterns if needed
- [ ] Add to LOG.md for future reference

---

### STEP 3: IMPLEMENTATION

**IMMEDIATE EXECUTION:**
- Apply fixes using atomic operations
- Test thoroughly in PMC context
- Verify data accuracy maintained

**DOCUMENTATION:**
```markdown
## LESSON LEARNED: {YYYY-MM-DD HH:MM:SS}
**Error Type:** {Runtime/Logic/Integration/Data}
**Problem:** {what went wrong}
**Root Cause:** {underlying reason}  
**Solution:** {what was changed}
**Prevention Rule:** {rule for future}
**Files Changed:** {list of files}
```

---

### PMC DOMAIN SPECIFIC DEBUGGING

**Data Integrity Checks:**
- Verify order counts before/after processing
- Confirm ROI calculations match expected ranges
- Check currency conversion accuracy
- Validate supplier matching logic

**Performance Validation:**
- Ensure 5-minute response time maintained
- Check memory usage for large datasets
- Verify Excel processing efficiency

**Business Logic Validation:**
- Confirm LEFT JOIN preserves all orders
- Validate ROI calculation edge cases (zero shortage, zero amount)
- Check multi-currency handling