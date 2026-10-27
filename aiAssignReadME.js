function aiAssignReadME() {

/**
┌─────────────────────────────────────┐
│          runAssignAll()             │
│                                     │
│ Main quarter-level controller       │
└─────────────────┬───────────────────┘
                  │
                  ▼
        ┌──────────────────────┐
        │ Get first sheet tab  │
        │ getSheets()[0]       │
        └──────────┬───────────┘
                   │
                   ▼
        ┌──────────────────────┐
        │   runForceAssign()   │
        │                      │
        │ 1. clearTaskAssignments()
        │ 2. resolveForcedAssignments()
        │ 3. writeForcedAssignments()
        │ 4. build initial
        │    historicalCount
        └──────────┬───────────┘
                   │
                   │ returns:
                   │ historicalCount
                   ▼
        ┌──────────────────────────┐
        │ getScheduleRowBounds()   │
        │                          │
        │ find firstRow / lastRow  │
        └──────────┬───────────────┘
                   │
                   ▼
       ┌────────────────────────────┐
       │ FOR row = firstRow..lastRow│
       └─────────────┬──────────────┘
                     │
                     ▼
          ┌───────────────────────┐
          │    runRowAssign()     │
          │                       │
          │ Assign ONE Sunday     │
          └───────────┬───────────┘
                      │
                      ▼
          ┌────────────────────────┐
          │   aiScheduleData()     │
          │                        │
          │ Build current row data │
          │                        │
          │ • emptyTasks           │
          │ • options              │
          │ • historicalCount      │
          │ • futureCount          │
          │ • occupiedRoles        │
          │ • blockingWorkers      │
          └───────────┬────────────┘
                      │
                      ▼
              ┌─────────────────┐
              │ Which method?   │
              └───────┬─────────┘
                      │
          ┌───────────┴────────────┐
          │                        │
          ▼                        ▼
┌───────────────────┐    ┌─────────────────────┐
│      No AI        │    │      AI Tool        │
│                   │    │                     │
│ runGASAssign()    │    │ aiCallEngine()      │
└─────────┬─────────┘    └──────────┬──────────┘
          │                         │
          │                         ▼
          │              ┌────────────────────┐
          │              │ Selected provider  │
          │              │                    │
          │              │ ChatGPT            │
          │              │ DeepSeek           │
          │              │ Claude             │
          │              │ Gemini             │
          │              └─────────┬──────────┘
          │                        │
          │                        ▼
          │              ┌────────────────────┐
          │              │     aiParser()     │
          │              │                    │
          │              │ raw response       │
          │              │       ↓            │
          │              │ parsed assignments │
          │              └─────────┬──────────┘
          │                        │
          └────────────┬───────────┘
                       │
                       ▼
             ┌─────────────────────┐
             │    aiValidator()    │
             │                     │
             │ Validate BOTH paths │
             │                     │
             │ • valid task        │
             │ • valid worker      │
             │ • blocking rules    │
             │ • duplicates        │
             └──────────┬──────────┘
                        │
              ┌─────────┴─────────┐
              │                   │
           INVALID              VALID
              │                   │
              ▼                   ▼
   ┌────────────────────┐   ┌──────────────────────┐
   │ AI path only:      │   │ writeAIAssignments()│
   │ aiRepair()         │   │                      │
   │                    │   │ write row to sheet   │
   └─────────┬──────────┘   └──────────┬───────────┘
             │                         │
             ▼                         ▼
      aiValidator() again     SpreadsheetApp.flush()
             │                         │
             └──────────────┬──────────┘
                            │
                            ▼
                ┌──────────────────────────┐
                │ Update historicalCount  │
                └────────────┬─────────────┘
                             │
              ┌──────────────┴──────────────┐
              │                             │
          AI path                      No-AI path
              │                             │
              ▼                             ▼
   Update after whole row       Already updated inside
   assignments are written      runGASAssign(), one task
                                at a time
              │                             │
              └──────────────┬──────────────┘
                             │
                             ▼
                  ┌────────────────────┐
                  │ Next Sunday / row  │
                  └────────────────────┘
--
For AI:

aiScheduleData()
      ↓
AI sees all empty tasks for current Sunday
      ↓
AI returns assignments
      ↓
validate
      ↓
write
      ↓
historicalCount updated afterward
--
For NO AI

runGASAssign()

Task 1
  ↓
choose:
lowest historicalCount
  ↓
tie → lowest futureCount
  ↓
tie → first option
  ↓
historicalCount++
  ↓
if blocking:
blockingWorkers.add(worker)
  ↓

Task 2 uses UPDATED history
and UPDATED blockingWorkers
  ↓
repeat
--

| Function                     | Responsibility                                     |
| ---------------------------- | -------------------------------------------------- |
| `runAssignAll()`             | Controls the whole quarter                         |
| `runForceAssign()`           | Handles all forced assignments and initial history |
| `clearTaskAssignments()`     | Clears task cells                                  |
| `resolveForcedAssignments()` | Finds rule/availability/single-option assignments  |
| `writeForcedAssignments()`   | Writes forced assignments                          |
| `getScheduleRowBounds()`     | Finds first/last schedule row                      |
| `runRowAssign()`             | Controls one Sunday                                |
| `aiScheduleData()`           | Builds one Sunday's assignment data                |
| `runGASAssign()`             | Deterministic No-AI scheduler                      |
| `aiCallEngine()`             | Routes to selected AI provider                     |
| `aiParser()`                 | Parses AI response                                 |
| `aiValidator()`              | Validates both AI and GAS assignments              |
| `aiRepair()`                 | Repairs invalid AI results only                    |
| `writeAIAssignments()`       | Writes final assignments                           |

historicalCount
    └── survives across the whole quarter

blockingWorkers
    └── exists only for the current Sunday

futureCount
    └── calculated by aiScheduleData() for each task

options
    └── task-specific valid candidate list

**/

}


