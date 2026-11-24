# CleanTrack Kawempe

A communityv oriented waste collection reporting system for Kawempe Division, Kampala.  Residents report missed collections, overflowing collection points and illegal dumping; coordinators assign collection teams and follow each case to closure.

The project has two technical parts: a **relational database** (Microsoft Access) and a **network model** (Cisco Packet Tracer), documented in a Word report.

---

## Project files

| File | Purpose |
|---|---|
| `CleanTrack_Project_Kawempe.accdb` | Access database (data tier) |
| `cleantracer_wa_fixed.pkt` | Corrected Packet Tracer network model |
| `CleanTrack_Kawempe_updated.docx` | Project report (design, screenshots, verification) |
| `CleanTrack_Access_Completion.bas` | VBA module that completes missing data and relationships in the database |
| `README.md` | This file |

The original, unfixed `cleantracer_wa.pkt` is superseded by `cleantracer_wa_fixed.pkt`.

**Requirements:** Microsoft Access (2013 or later), Cisco Packet Tracer 9.0 or later, Microsoft Word.

---

## 1. Database (Microsoft Access)

### Tables

| Table | Purpose | Key fields |
|---|---|---|
| `Users` | Residents, coordinators, collection team, admins | UserID, FullName, PhoneNumber, UserRole, AccountStatus |
| `ServiceAreas` | Kawempe parishes and zones | AreaID, ParishName, ZoneName, CollectionDay |
| `Reports` | Waste reports submitted by residents | ReportID, ReferenceNo (e.g. `CT-KAW-2026-001`), ReporterID, AreaID, Category, Priority, CurrentStatus, CreatedAt |
| `Assignments` | Coordinator dispatch of a team to a report | AssignmentID, ReportID, CoordinatorID, TeamMemberID, AssignedAt, TargetDate |
| `StatusEvents` | Audit trail of every status change | EventID, ReportID, PreviousStatus, NewStatus, UpdatedBy, EventTimeStamp |
| `Notifications` | Alerts sent to the reporting resident | NotificationID, ReportID, RecipientID, MessageText, DeliveryState |

Workflow statuses: Submitted, Under Review, Assigned, In Progress, Completed, Closed.

### Other objects

- **Forms:** `frmSubmitReport`, `frmCoordinatorDispatch`
- **Queries:** `qryOpenReportsByParish`, `qrySummaryByCategory`
- **Report:** `rptKawempeServiceSummary`

### Completing the database

Back up the `.accdb`, then:

1. Open the database and press **Alt+F11**.
2. **File > Import File...** and choose `CleanTrack_Access_Completion.bas`.
3. Press **Ctrl+G**, type `CompleteCleanTrack` and press Enter.

The script normalises role names (`Residents` to `Resident`), adds two sample collection-team users, fills `CollectionDay`, creates sample assignments for reports 1 to 3, and adds the missing Users relationships. Each step checks first, so it is safe to run twice. The names and phone numbers it adds are sample data.

### Known gaps

- `Category` and `NewStatus` are multi-value (lookup) fields, whereas the report describes them as plain text.
- Some sample reports are filed under Coordinator/Admin users rather than residents.

---

## 2. Network (Cisco Packet Tracer)

Open `cleantracer_wa_fixed.pkt`. It models three sites joined through an ISP router.

| Site | Subnet | Devices |
|---|---|---|
| A: Coordination Office | 192.168.10.0/24 | R1-Kawempe, SW1-Office, Coordinator-PC1, Admin-PC2, Dispatcher-Laptop |
| B: Cloud Data Center | 192.168.20.0/24 | R2-DataCenter, SW2-DC, CleanTrack-Web-Server, CleanTrack-DB-Server |
| C: Community Wireless Zone | 192.168.30.0/24 | WRT300N-Kawempe (SSID `CleanTrack_Community_WiFi`), Resident-Smartphone1, CollectionTeam-Tablet1 |

### Addressing

| Segment | Network | Addresses |
|---|---|---|
| Office LAN | 192.168.10.0/24 | Gateway .1 (R1 Fa0/0); PCs .10, .11, .12; SW1 management .2 |
| Data center LAN | 192.168.20.0/24 | Gateway .1 (R2 Fa0/0); Web server .10, DB server .20; SW2 management .2 |
| Wireless LAN | 192.168.30.0/24 | Gateway .1 (WRT300N); DHCP pool .100 to .149 |
| WAN 1 | 10.0.1.0/30 | R1 S0/3/0 = .1, ISP S0/3/0 = .2 |
| WAN 2 | 10.0.2.0/30 | R2 S0/3/0 = .1, ISP S0/1/0 = .2 |
| WAN 3 | 10.0.3.0/30 | ISP Fa0/0 = .1, WRT300N Internet port = .2 |

### Services

- **Routing:** static. R1 and R2 use a default route to the ISP; ISP-Cloud has routes to all three LANs.
- **DHCP:** WRT300N serves the wireless LAN, with DNS 192.168.20.10.
- **DNS** (web server): `cleantrack.kawempe.ug` and `www.cleantrack.kawempe.ug` resolve to 192.168.20.10; `db.cleantrack.kawempe.ug` resolves to 192.168.20.20.
- **HTTP/HTTPS** (web server): CleanTrack landing page.

### Fixes applied to the original file

- WRT300N Internet port was `0.0.0.0` with no gateway; now 10.0.3.2/30 with gateway 10.0.3.1.
- DNS service was disabled with no records; now enabled with the records above.
- Switches were both named "Switch" with no management address; now SW1-Office and SW2-DC with management IPs and default gateways.
- Device names aligned with the report.

### Verification tests

Run these in Packet Tracer and capture screenshots for the report.

| # | From | Test | Expected |
|---|---|---|---|
| 1 | Coordinator-PC1 | `ping 192.168.10.1` | Replies |
| 2 | Coordinator-PC1 | `ping 192.168.20.10` | Replies across the WAN |
| 3 | Dispatcher-Laptop | `ping 192.168.20.20` | Replies |
| 4 | Coordinator-PC1 | `ping cleantrack.kawempe.ug` | Resolves and replies |
| 5 | Coordinator-PC1 | Browser: `http://cleantrack.kawempe.ug` | Landing page shown |
| 6 | Resident-Smartphone1 | Browser: `http://192.168.20.10` | Page loads over Wi-Fi |
| 7 | CollectionTeam-Tablet1 | `ping 192.168.20.10` | Replies |
| 8 | CleanTrack-Web-Server | `ping 192.168.20.20` | Replies |

### Things to check after opening

- Web server, Services > DNS: confirm the three records are present; add them by hand if not.
- WRT300N, Setup: confirm the Internet connection type is Static IP with 10.0.3.2.
- The wireless router uses NAT, so wireless devices can reach the servers but sites A and B cannot initiate connections to wireless devices.

---

## 3. Report

`CleanTrack_Kawempe_updated.docx` covers the architecture, database design, network design and addressing. Section 3.3 holds the services and verification tests, with yellow placeholders where the cross-site test screenshots still need to be inserted. Section 3.4 lists the network corrections.

---

## Author
NAME: ELIJAH SENABULYA.