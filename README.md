# SPH Duty Roster Scripts 🏥

A Google Apps Script utility tailored to automate and optimize monthly on-call rotations for surgical residents at SPH. It replaces manual scheduling layouts with a randomized, constraint-based balancing script.

---

## 🚀 Core Features & Logic

* **Workload Balancing(Fair Share):** The script keeps a live tally of everyone's shifts. When filling an empty weekday slot, it automatically looks for the person who has worked the absolute least so far, ensuring the workload is split evenly.
* **Bias Protection(Tie-Breaker Shuffling):** If multiple residents are tied for the fewest number of shifts, the script automatically shuffles their names randomly. This prevents alphabetical bias so the same people don't consistently end up with the extra work.
* **Fatigue Control (No Back-to-Back Shifts):** Working two days in a row (like Friday and Saturday) is exhausting. The script scans the calendar 10 times over to spot these back-to-back conflicts and safely trades shifts with other available residents to give people a break.
* **Leave Protection:** The script cross-references a digital leave calendar before making any assignments. If a resident has requested a specific date off, the script automatically skips them and moves to the next eligible person.
* **Interactive Frontend Modals:** Instead of typing directly into the spreadsheet, shifts are managed through clean popup windows. These windows act as a safety gate—instantly catching typos, missing entries, or accidental double-bookings before anything changes on the main sheet.

---

---

## 🛠️ Setup & Installation

### 1. Paste the Code Script Files
1. Open your Google Sheet and navigate to **Extensions** ➔ **Apps Script**.
2. Create 5 Script files (`main.gs`, `builder.gs`, `personnel.gs`, `weekend.gs`, `weekday.gs`) and paste their corresponding repository `.js` codes into them.

### 2. Add the User Interface Templates
1. Click the **`+`** icon in the script editor and select **HTML**.
2. Create 3 HTML files named exactly: `leaveForm`, `weekendForm`, and `weekdayForm`. Paste their respective repository HTML layouts inside.
3. Save the project.

### 3. Run the Program
1. Refresh your main Google Sheet browser tab.
2. Click the new **`SPH Roster Menu`** toolbar drop-down option at the top of the window and select **Generate New Duty List**.

> 💡 **Note:** On the first run, grant the standard Google macro permission by clicking **Advanced** ➔ **Go to SPH Roster (unsafe)**.