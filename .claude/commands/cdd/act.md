# Action: Execute Implementation Plan Atomically  

## OBJECTIVE:
Execute a complete implementation plan from a task file atomically, ensuring all changes succeed together or fail cleanly.

**INPUT:** Task file path as `$ARGUMENTS` (e.g., "001_add_email_validation.md")

**SUCCESS CRITERIA:** All checklist items completed successfully, or clean failure with system in consistent state.

---

### STEP 1: LOAD COMPLETE CONTEXT

**CONTEXT LOADING:**
- [ ] Read specified task file for implementation plan
- [ ] Load `.claude/VISION.md` for strategic alignment  
- [ ] Load `.claude/CLAUDE.md` for technical patterns
- [ ] Scan `.claude/LOG.md` for relevant lessons learned
- [ ] Analyze existing codebase patterns for consistency

**PMC CONTEXT PRIORITIES:**
- [ ] Data accuracy requirements (manufacturing decisions critical)
- [ ] 5-minute response time constraints  
- [ ] Excel integration patterns
- [ ] ROI calculation integrity
- [ ] LEFT JOIN architecture preservation

---

### STEP 2: ATOMIC EXECUTION PREPARATION

**PRE-EXECUTION VALIDATION:**
- [ ] Verify task plan is complete and specific
- [ ] Check all required files exist and are accessible
- [ ] Ensure no external dependencies are missing
- [ ] Validate that checklist items are actionable

**ATOMIC OPERATION SETUP:**
- [ ] Prepare all file modifications in memory first
- [ ] Plan rollback strategy for any failures
- [ ] Ensure no partial states possible

---

### STEP 3: EXECUTE IMPLEMENTATION PLAN

**SYSTEMATIC EXECUTION:**

Execute checklist items in order:
- [ ] Process each checklist item completely
- [ ] Apply PMC domain patterns (data integrity, performance)
- [ ] Maintain Excel processing compatibility
- [ ] Preserve ROI calculation accuracy
- [ ] Test each component as implemented

**QUALITY GATES:**
- [ ] Code follows existing patterns
- [ ] PMC business rules respected
- [ ] No data integrity compromises
- [ ] Performance requirements maintained

---

### STEP 4: INTEGRATION TESTING

**PMC-SPECIFIC TESTING:**
- [ ] Test with actual Excel files if data processing involved
- [ ] Verify ROI calculations with sample data
- [ ] Check currency conversion accuracy
- [ ] Validate dashboard display correctness
- [ ] Ensure 5-minute response time maintained

**SYSTEM INTEGRATION:**
- [ ] Verify all components work together
- [ ] Test error handling and edge cases  
- [ ] Check data flow integrity
- [ ] Validate user interface consistency

---

### STEP 5: COMPLETION AND LOGGING

**SUCCESS PATH:**
- [ ] Mark task as completed in task file
- [ ] Update task status and completion time
- [ ] Log any lessons learned to LOG.md
- [ ] Clean up temporary files

**FAILURE HANDLING:**
- [ ] Rollback all changes to clean state
- [ ] Document failure reason in task file
- [ ] Identify what needs revision in plan
- [ ] Recommend using `/cdd:revise` for plan updates

**AUTOMATIC LOGGING:**
```markdown
---
## COMPLETED TASK: {YYYY-MM-DD HH:MM:SS}
**Task File:** {task_file_name}
**Total Steps:** {number}
**Summary:** {list of completed actions}
**PMC Impact:** {manufacturing/analysis improvements}
---
```

---

### PMC DOMAIN EXECUTION PATTERNS

#### **Data Processing Tasks**
- Always preserve order integrity (LEFT JOIN pattern)
- Validate calculations at each step
- Test with multiple currency scenarios
- Ensure Excel compatibility maintained

#### **Dashboard/UI Tasks** 
- Test display with real PMC data
- Verify KPI calculations accuracy
- Check responsive design for management use
- Validate export functionality

#### **Analysis Engine Tasks**
- Maintain ROI calculation precision
- Preserve supplier selection logic
- Test performance with full dataset
- Verify multi-file processing integrity

---

### ATOMIC OPERATION GUARANTEES

**ALL-OR-NOTHING PRINCIPLE:**
- Either complete implementation succeeds entirely
- Or system returns to exact previous state
- No partial implementations left behind
- Clear success/failure reporting

**CONSISTENCY MAINTENANCE:**
- File changes batched and applied together
- Database/data integrity preserved
- PMC business rules never violated
- System remains in working state throughout