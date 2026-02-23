# 📊 Analyzing SPSS (.sav) Files with GitHub Copilot CLI

A step-by-step guide on how to use **GitHub Copilot CLI** to explore, review, and analyze SPSS data files — no SPSS software required.

---

## 📁 What You Need

- A `.sav` SPSS data file (e.g., `Employee data.sav`)
- [Python](https://www.python.org/downloads/) installed
- [GitHub Copilot CLI](https://githubnext.com/projects/copilot-cli) installed and authenticated

---

## 🔧 Installation

### 1. Install GitHub Copilot CLI

```bash
# Install via npm
npm install -g @githubnext/github-copilot-cli

# Authenticate with your GitHub account
github-copilot-cli auth
```

> **Note:** You need an active GitHub Copilot subscription to use the CLI.

### 2. Python Dependencies

GitHub Copilot CLI will automatically install missing packages when needed, but you can also install them upfront:

```bash
pip install pyreadstat pandas
```

| Package | Purpose |
|---|---|
| `pyreadstat` | Reads SPSS `.sav` files into Python |
| `pandas` | Data manipulation and pivot tables |

---

## 🗂️ Folder Setup

Place your `.sav` file in a working directory:

```
D:\SPSS_PS\
├── Employee data.sav     ← your SPSS data file
```

Open a terminal in that folder and launch GitHub Copilot CLI.

---

## 💬 Example Session

Below is a real conversation showing how to analyze an SPSS file using Copilot CLI.

---

### Step 1 — Ask Copilot to Read and Review the File

**You:**
> There is a SPSS file, can you read and review it?

**Copilot CLI:**
Copilot detected the `Employee data.sav` file in the directory, installed `pyreadstat` automatically, and returned a full review:

```
Shape: (474 rows, 10 columns)

Variables:
  id        – Employee Code
  gender    – Gender (m/f)
  bdate     – Date of Birth
  educ      – Educational Level (years)
  jobcat    – Employment Category (1=Clerical, 2=Custodial, 3=Manager)
  salary    – Current Salary
  salbegin  – Beginning Salary
  jobtime   – Months since Hire
  prevexp   – Previous Experience (months)
  minority  – Minority Classification (0=No, 1=Yes)

Missing values: 1 (date of birth only)
```

**Key Stats Summary:**

| Variable | Min | Mean | Max |
|---|---|---|---|
| Salary | $15,750 | $34,420 | $135,000 |
| Beginning Salary | $9,000 | $17,016 | $79,980 |
| Education | 8 yrs | 13.5 yrs | 21 yrs |
| Tenure | 63 mo | 81 mo | 98 mo |
| Minority | 22% | — | — |

---

### Step 2 — Generate a Cross-tabulation Table

**You:**
> Can you generate a table: Beginning Salary by gender and education?

**Copilot CLI:**
Copilot created a Python script and generated the following pivot table:

```
Beginning Salary by Gender and Education
=====================================================================================
Education       Mean Female   N Female      Mean Male     N Male     Mean Total    N Total
-------------------------------------------------------------------------------------
8 yrs               $11,720         30        $14,817         23        $13,064         53
12 yrs              $12,335        128        $15,114         62        $13,242        190
14 yrs                    -          -        $15,625          6        $15,625          6
15 yrs              $13,768         33        $16,343         83        $15,611        116
16 yrs              $17,928         24        $25,363         35        $22,338         59
17 yrs              $12,750          1        $28,320         10        $26,905         11
18 yrs                    -          -        $32,240          9        $32,240          9
19 yrs                    -          -        $34,764         27        $34,764         27
20 yrs                    -          -        $36,240          2        $36,240          2
21 yrs                    -          -        $37,500          1        $37,500          1
Total               $13,092        216        $20,301        258        $17,016        474
```

**Key Findings:**
- 📌 Males earn **$7,209 more** on average in beginning salary than females
- 📈 Higher education correlates with higher beginning salary for both genders
- 👥 Most employees (190) have 12 years of education (high school level)
- 🚩 Women are absent from education levels 14, 18, 19, 20, and 21 years

---

## 🔁 What Copilot CLI Does Behind the Scenes

```
User asks question
      ↓
Copilot reads directory → finds .sav file
      ↓
Generates Python code using pyreadstat + pandas
      ↓
Runs the code in the terminal
      ↓
Returns clean, readable output with insights
```

---

## 💡 More Questions You Can Ask

Here are other questions you can ask Copilot CLI about your `.sav` file:

| Question | What it does |
|---|---|
| "Show me the distribution of salary by job category" | Bar chart or summary table |
| "Are there significant differences in salary between genders?" | t-test or ANOVA |
| "Show correlations between education, experience and salary" | Correlation matrix |
| "How many employees are in each job category?" | Frequency table |
| "Plot salary vs education" | Scatter plot |
| "Check for outliers in salary" | Boxplot or IQR analysis |
| "Predict salary using regression" | Linear regression model |

---

## 📦 Required Python Packages (Summary)

```bash
pip install pyreadstat pandas matplotlib scipy
```

| Package | Use Case |
|---|---|
| `pyreadstat` | Read `.sav` files |
| `pandas` | Tables, pivot, groupby |
| `matplotlib` | Charts and plots |
| `scipy` | Statistical tests (t-test, ANOVA) |

---

## 📝 Notes

- Copilot CLI works in **plain English** — no coding required from you
- It can handle SPSS **variable labels** and **value labels** automatically
- The generated Python scripts are saved in your folder and reusable
- Works on Windows, macOS, and Linux

---

*Generated with GitHub Copilot CLI — your AI pair programmer in the terminal.*
