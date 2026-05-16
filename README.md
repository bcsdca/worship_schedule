# Cantonese Worship Duty Reminder & Scheduling System

A comprehensive Google Apps Script (GAS) automation platform designed to manage quarterly Cantonese worship scheduling, worker assignments, worship duty reminders, and worship coordination workflows.

This project serves as the central quarterly scheduling, communication, and coordination system for the quarterly Cantonese worship ministry.

The system automates:
- Quarterly worship scheduling
- Worker availability collection
- Smart A/V duty assignments
- Automatic detection of any worker who should not get assigned that week base on the unavialibiity sheet
- Automatic checking/detection of any worker who was double booked in any given week
- Weekly reminder emails
- Worship coordination notifications
- Worship dashboard generation
- Livestream statistics collection
- Integration with downstream slide-generation workflows

---

# ✨ Core Features

## 📅 Quarterly Worship Schedule Management

At the beginning of every quarter:
1. A new worship schedule spreadsheet is created
2. New worship dates and sermon assignments are entered
3. Preachers and worship leaders are assigned
4. Worker availability collection begins

The system supports ongoing schedule maintenance throughout the quarter.

---

# 📧 Worker Availability Collection

Before scheduling assignments started:
- Availability request emails are automatically sent to workers
- Workers submit unavailable dates
- Availability data is collected into the spreadsheet

This availability information becomes the foundation for automated worker scheduling.

---

# 🤖 Smart Automatic Worker Assignment

One of the key features of this project is the `smartAssignWorker()` scheduling engine.

The scheduling algorithm automatically assigns weekly worship workers based on:
- Individual worker availability
- Existing workload balance
- Worship expertise/roles
- Fair workload distribution
- Historical assignment counts

Supported worship roles include:
- Audio
- Livestream
- Projection/PPT
- Camera
- Technical support
- Other A/V ministry duties

The goal is to ensure:
- Balanced workload across all workers
- Fair scheduling rotation
- Reduced manual coordination effort

---

# 👥 Schedule Review & Confirmation Workflow

Before the worship schedule goes live:
- Workers are given an opportunity to review assignments
- Adjustments can be made manually
- Final confirmation is performed

This hybrid workflow combines:
- Automated scheduling
- Human review and flexibility

---

# 🔄 Weekly Worship Duty Reminder Emails

Every week, the GAS automatically sends worship reminder emails to all workers assigned for that Sunday's worship service.

Reminder emails include:
- Worship date
- Assigned duties
- Worship team information
- Sermon details
- Scripture passages
- Important reminders

This significantly reduces manual follow-up work for coordinators.

---

# 🖼 Tuesday Sermon Information PNG Attachment

Every Tuesday, the worship reminder email may include a PNG image attachment containing:
- This week's sermon title
- Sermon scripture passage

### Conditional Behavior
- If the sermon information file is received before **11:00 AM Tuesday morning**, the PNG image is included in the reminder email.
- Otherwise, the reminder email will not contain the PNG attachment.

This helps worship workers prepare earlier in the week.

---

# 📣 Schedule Change Notifications

During the quarter, worship schedules occasionally change.

Whenever schedule modifications occur:
- Additional reminder emails are automatically generated
- Only affected workers are notified
- Updated assignment information is included

This minimizes confusion and ensures workers stay informed about last-minute changes.

---

# 📊 Cantonese Livestream Statistics Collection

The GAS also automatically collects weekly Cantonese livestream statistics for:
- Historical analysis
- Worship reporting
- Trend tracking
- Future processing and analytics

This data can later be used for ministry reporting and livestream planning.

---

# 📱 Saturday Morning Text Reminder System

The project previously included a Saturday morning SMS/text reminder feature.

The text reminder system:
- Successfully operated for approximately one year
- Automatically reminded worship workers before Sunday service

However:
- The functionality now appears to be blocked by Google gmail to SMS tool
- It is suspected that the messages may have been reported as spam
- As a result, outbound text reminders are currently restricted or disabled

The email reminder workflow remains fully operational.

---

# 📈 Worship Dashboard

The project automatically generates a worship dashboard that provides:

## Worker Workload Tracking
Displays:
- Total assignments per worker
- Role distribution
- Quarterly workload balancing
- Assignment statistics

## Worker Contact Directory
Acts as the central source of truth for:
- Phone numbers
- Email addresses
- Worship expertise areas
- Ministry roles

## Assignment Coordination
Provides worship coordinators with:
- Quick visibility into worker assignments
- Staffing balance insights
- Ministry planning support

This dashboard serves as the operational control center for the worship ministry.

---

# 📂 Centralized Source of Truth

The system serves as the primary centralized database for:
- Worker contact information
- Worker availability
- Ministry expertise
- Worship schedules
- Duty assignments
- Assignment history

This ensures consistency across all worship coordination workflows.

---

# 🔗 Integration with Other GAS Projects

This project integrates with other worship automation systems, including:
- `slide_src` GAS project
- Worship PowerPoint generation workflows
- Worship PPT email distribution systems

For example:
- The assigned PPT worker information is stored in the `weeklyShare` spreadsheet by this weekly reminder GAS
- Downstream GAS projects (slide_src) later use this information to distribute generated worship presentation links

---

# 🏗 High-Level Workflow

```text
Quarterly Schedule Creation
              ↓
 Worker Availability Collection
              ↓
 smartAssignWorker() Scheduling
              ↓
 Worker Review & Confirmation
              ↓
 Weekly Reminder Emails
              ↓
 Schedule Change Notifications
              ↓
 Livestream Statistics Collection
              ↓
 Dashboard Reporting & Analytics
```

---

# 🛠 Technologies Used

- Google Apps Script (GAS)
- GmailApp
- SpreadsheetApp
- Google Triggers
- Google Drive Services
- Email Automation
- PNG Attachment Processing
- Dashboard Automation

---

# 📋 Main Functional Areas

| Function Area | Purpose |
|---|---|
| Quarterly Schedule Creation | Create new worship schedules |
| Availability Collection | Gather worker unavailable dates |
| smartAssignWorker() | Automatic balanced scheduling |
| Reminder Emails | Weekly worship duty reminders |
| PNG Sermon Attachments | Tuesday sermon information images |
| Schedule Change Alerts | Notify affected workers |
| Livestream Statistics | Collect weekly livestream data |
| Dashboard Generation | Workload and contact reporting |
| Worker Directory | Centralized contact information |
| Trigger Management | Automated scheduling tasks |

---

# 🚀 Benefits

- Reduces manual scheduling effort
- Balances ministry workload fairly
- Improves communication efficiency
- Automates reminder workflows
- Centralizes worship coordination data
- Reduces scheduling conflicts
- Improves worship ministry organization
- Enables scalable volunteer coordination

---

# 📌 Example Weekly Workflow

Every week:
1. Worship assignments already exist in the quarterly schedule
2. Tuesday reminder emails are sent
3. Sermon PNG image may be attached
4. Workers receive duty reminders
5. Any schedule changes trigger updated notifications
6. Livestream statistics are collected after worship service

Result:
- Minimal manual coordination required
- Better volunteer communication
- Improved scheduling consistency

---

# 🔮 Possible Future Improvements

- Restore SMS/text reminder functionality
- Better anti-spam handling for notifications
- Mobile-friendly dashboard
- AI-assisted worker scheduling
- Calendar integration
- Livestream analytics visualization
- Integration with Planning Center or ProPresenter

---

# 🙏 Acknowledgements

This project was created to streamline Cantonese worship ministry coordination and reduce repetitive administrative work for worship leaders and A/V ministry coordinators.

---

# 📜 License

This project is open source and intended for church/community ministry usage.
