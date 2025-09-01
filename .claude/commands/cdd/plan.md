# Action: Create Task and Generate Detailed Implementation Plan

## OBJECTIVE:

To create a new task file and generate the most detailed and specific implementation plan possible by clarifying all ambiguous requirements through targeted questioning.

**INPUT:** Task description as `$ARGUMENTS` (e.g., "Add user authentication with JWT tokens")

**SUCCESS CRITERIA:** Plan is complete when task file is created, all requirements are clarified, and checklist items are so specific that any competent developer could execute them without further questions.

**CORE PRINCIPLE:** When something is uncertain or vague, always ask the user relevant questions to make sure things are clear before proceeding.

---

### STAGE 1: TASK FILE CREATION

**AUTOMATIC TASK CREATION:**
1. **Generate filename:**
   - Create descriptive snake_case name from `$ARGUMENTS` (max 50 chars)
   - Check `.claude/cddtasks/` directory for next number (001_, 002_, etc.)
   - Format: `{number}_{task_name}.md`
   
2. **Create task file** at `.claude/cddtasks/{filename}` with this structure:
```markdown
# Task: {ARGUMENTS}
**Created:** {YYYY-MM-DD HH:MM:SS}
**Status:** Planning

## Requirements
[To be filled during clarification]

## Implementation Plan
[To be generated after requirements are clear]

## Checklist
[To be created with specific, actionable items]
```

---

### STAGE 2: LOAD CONTEXT FOR INFORMED PLANNING

**CONTEXT LOADING:**
- [ ] Read `.claude/VISION.md` for strategic direction
- [ ] Read `.claude/CLAUDE.md` for technical constraints and patterns
- [ ] Scan `.claude/LOG.md` for similar tasks and lessons learned
- [ ] Analyze existing codebase structure if task involves code changes
- [ ] Review package.json, requirements.txt, or similar for available dependencies

**Purpose:** Ensure all recommendations align with project goals and leverage existing capabilities.

---

### STAGE 3: STRATEGIC DIALOGUE FOR REQUIREMENT CLARIFICATION

**CLARIFICATION STRATEGY:**

Instead of making assumptions about implementation details, engage in **structured dialogue** to explore different approaches:

#### **Present Multiple Approaches**
For complex tasks, present 2-3 different implementation strategies with trade-offs:

**Example for "Add caching system":**
- **Approach A (Redis):** External caching, shared across instances, requires Redis setup
- **Approach B (In-memory):** Simple, fast, but lost on restart
- **Approach C (File-based):** Persistent, no external deps, but slower

#### **Ask Strategic Questions**
Focus on decisions that significantly impact the implementation:

**Architecture Questions:**
- "Should this integrate with existing auth system or be standalone?"
- "Do you prefer a lightweight solution or full-featured with configuration options?"
- "Should this handle edge cases automatically or fail fast with clear errors?"

**Scope Questions:**
- "Should we handle [specific edge case] now or address it later?"
- "Do you want [related feature] included or kept separate?"
- "Should this be configurable or use sensible defaults?"

**Technical Questions:**
- "Any preference between [option A] vs [option B] for [specific aspect]?"
- "Should this follow [existing pattern X] or [existing pattern Y] from the codebase?"

#### **AVOID over-clarification:**
- Don't ask about obvious implementation details
- Don't present more than 3 options (causes decision paralysis)
- Don't ask questions where the answer doesn't affect the plan

---

### STAGE 4: DETAILED PLAN GENERATION

After clarification, generate a detailed implementation plan:

#### **Requirements Section**
Update task file with:
```markdown
## Requirements
### Functional Requirements
- [ ] [Specific capability 1]
- [ ] [Specific capability 2]

### Technical Constraints
- [ ] [Must use existing X]
- [ ] [Must be compatible with Y]

### Success Criteria
- [ ] [Measurable outcome 1]
- [ ] [Measurable outcome 2]
```

#### **Implementation Plan**
Generate specific, ordered steps:
```markdown
## Implementation Plan

### Phase 1: [Preparation/Setup]
1. [Specific preparation step]
2. [Another specific step]

### Phase 2: [Core Implementation]
1. [Specific implementation step]
2. [Another specific step]

### Phase 3: [Integration & Testing]
1. [Specific testing approach]
2. [Integration steps]
```

#### **Checklist Generation**
Create atomic, verifiable checklist items:

**GOOD checklist items:**
- [ ] Create `UserAuth.js` with login() and logout() methods
- [ ] Add JWT token validation middleware to `/api/*` routes  
- [ ] Update login form to call AuthService.login() on submit
- [ ] Add error handling for expired tokens with redirect to login

**BAD checklist items:**
- [ ] Implement authentication *(too vague)*
- [ ] Handle edge cases *(not specific)*
- [ ] Make it work *(not actionable)*

---

### STAGE 5: FINAL VALIDATION

Before completing:

**SANITY CHECKS:**
- [ ] Every checklist item is specific and actionable
- [ ] Plan addresses all clarified requirements  
- [ ] Technical approach aligns with project patterns
- [ ] No major assumptions left unvalidated
- [ ] Plan is feasible given project constraints

**SAVE UPDATED TASK FILE** with complete plan.

---

### EXAMPLE WORKFLOWS

#### **Simple Task Example:**
```
User: "/cdd:plan Add email validation to registration form"

1. Create task file: 001_add_email_validation.md
2. Load context (find existing validation patterns)
3. Ask: "Should this use HTML5 validation, custom regex, or external service?"
4. Generate plan with specific validation implementation
5. Create actionable checklist
```

#### **Complex Task Example:**
```
User: "/cdd:plan Implement caching system"

1. Create task file: 002_implement_caching_system.md  
2. Load context (check existing architecture)
3. Present caching options (Redis, in-memory, file-based)
4. Ask about scope, performance requirements, persistence needs
5. Generate detailed multi-phase implementation plan
6. Create comprehensive checklist covering setup, implementation, testing
```

---

### SUCCESS INDICATORS

**You've succeeded when:**
- ✅ Task file exists with complete requirements and plan
- ✅ All major architectural decisions are resolved
- ✅ Checklist items are specific enough for `/cdd:act` to execute
- ✅ Plan leverages existing project patterns and capabilities
- ✅ No significant assumptions or ambiguities remain

**This creates the foundation for successful execution via `/cdd:act`.**