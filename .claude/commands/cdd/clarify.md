# Action: Clarify and Execute Simple Changes

## OBJECTIVE:
Quickly clarify simple requirements through adaptive questioning and execute changes directly without creating task files.

**INPUT:** Change description as `$ARGUMENTS` (e.g., "Add email validation to registration form")

**SUCCESS CRITERIA:** Change is implemented correctly with minimal clarification overhead.

---

### ADAPTIVE CLARIFICATION STRATEGY

#### **For Very Detailed Input** (minimal clarification needed)
*Example: "Change error message in src/auth/login.js line 45 from 'Invalid' to 'Invalid email or password'"*

**Action:** Execute directly with minimal confirmation
- [ ] Confirm the specific change
- [ ] Implement immediately
- [ ] Test the change

#### **For Medium Detailed Input** (2-3 targeted questions)
*Example: "Add email validation to registration form"*

**Ask targeted questions:**
- Which validation approach? (HTML5, regex, external service)
- What error message format?
- Where should validation trigger? (on blur, submit, etc.)

#### **For Low Detailed Input** (multiple clarification questions)
*Example: "Fix the login issue"*

**Systematic clarification:**
- What specific problem are users experiencing?
- Where does the issue occur? (login form, authentication, redirect, etc.)
- Any error messages or symptoms?
- What should the correct behavior be?

---

### PMC-SPECIFIC CLARIFICATION PATTERNS

#### **Data Display Changes**
*Example: "Show shortage percentages on dashboard"*

**Quick clarifications:**
- Percentage of what? (total shortage/total order value?)
- Where on dashboard? (KPI cards, table, chart?)
- Number format? (0.5% or 0.005?)

#### **Excel Processing Changes**
*Example: "Fix Excel upload issues"*

**Essential questions:**
- What specific issue? (fails to load, missing data, wrong format?)
- Which file types affected? (orders, suppliers, inventory?)
- Error occurs when? (upload, processing, display?)

#### **ROI Calculation Changes**
*Example: "Improve ROI calculation accuracy"*

**Critical clarifications:**
- What's currently inaccurate? (formula, currency conversion, edge cases?)
- Expected behavior for zero shortage orders?
- Currency conversion rates need updating?

---

### EXECUTION PATTERNS

#### **UI/Display Changes**
1. Identify affected components
2. Make minimal changes
3. Test visual output
4. Verify data accuracy maintained

#### **Data Processing Changes**  
1. Ensure data integrity preserved
2. Apply changes with validation
3. Test with sample data
4. Verify performance maintained

#### **Excel Integration Changes**
1. Test with actual Excel files
2. Verify data mapping correctness
3. Check encoding (UTF-8/GBK) handling
4. Validate output format

---

### DIRECT IMPLEMENTATION

**No task file creation** - execute changes immediately:
- [ ] Load relevant context from CLAUDE.md patterns
- [ ] Apply PMC domain rules (data accuracy priority)
- [ ] Implement change atomically
- [ ] Test in PMC context
- [ ] Document if significant lesson learned

### SUCCESS CRITERIA

**Efficient completion when:**
- ✅ Change implemented with minimal back-and-forth
- ✅ PMC business rules respected (data accuracy, performance)
- ✅ No task file overhead for simple changes
- ✅ User gets immediate, working solution