# QC360 - AWS & Python Powered Data Validation Framework


QC360 is a Python & AWS-based, configuration-driven Data Quality Validation Framework that automates large-scale QC execution using AWS Athena. The framework dynamically generates parameterized queries from Excel configurations, executes them in parallel, retrieves results from Athena/S3, and generates validation outputs in Excel, CSV, tab-delimited, pipe-delimited, or custom file formats.

Built using Python, boto3, pandas, openpyxl, and xlwings, QC360 enables users to maintain QC logic, mappings, filters, environments, templates, and reporting parameters directly within a centralized Excel workbook (`reporting_config.xlsm`) without requiring significant code changes.

The framework supports:

- Dynamic parameterized query execution
- AWS Athena integration
- Parallel query processing
- Excel-based configuration management
- Dynamic worksheet and template creation
- Automated report generation
- Multi-process QC support
- Configurable output formats
- Detailed execution logging
- One-click execution through Excel Macro or VS Code

---

# Prerequisites

### Python

Install:

```text
Python 3.11.9 or higher
```

Recommended Source:

```text
Microsoft Store
```

### Visual Studio Code

Download:

```text
https://code.visualstudio.com/
```

---

# Repository Setup

Clone the repository:

```bash
git clone <repository-url>
cd QC360
```

---

# Virtual Environment Setup

Open VS Code Terminal and execute:

### Create Virtual Environment

```bash
python -m venv venv
```

### Activate Virtual Environment (CMD)

```bash
venv\Scripts\activate
```

### Activate Virtual Environment (PowerShell)

```powershell
.\venv\Scripts\Activate.ps1
```

---

# Install Required Libraries

```bash
pip install boto3==1.34.69 botocore==1.34.69 et_xmlfile==2.0.0 jmespath==1.0.1 numpy==1.26.4 openpyxl==3.1.5 pandas==2.2.2 python-dateutil==2.9.0.post0 pytz==2024.1 pywin32==306 s3transfer==0.10.1 six==1.16.0 urllib3==2.2.2 xlwings==0.30.12
```

---

# Configuration

Open:

```text
reporting_config.xlsm
```

Navigate to:

```text
team-mkt-chnl_cfg
```

Configure:

- Environment Details
- AWS Access Key
- AWS Secret Key
- Bucket
- S3 Path
- Region
- Data Date
- Process Configuration
- Product & Market Mappings
- QC Parameters
- Validation Rules

QC360 dynamically reads all runtime parameters directly from this worksheet.

---

# Running QC360

## Option 1: One-Click Macro Execution

1. Open `reporting_config.xlsm`
2. Navigate to the Configuration Worksheet
3. Click the Macro Run Button
4. Wait for completion
5. Review generated outputs

## Option 2: Execute from VS Code

Ensure the virtual environment is activated.

Run:

```bash
python index.py
```

or simply use:

```text
Run Python File
```

from VS Code.

---

# Framework Workflow

```text
Configure reporting_config.xlsm
            ↓
Load Runtime Parameters
            ↓
Generate Dynamic Athena Queries
            ↓
Execute Queries in Parallel
            ↓
Retrieve Results from Athena/S3
            ↓
Perform QC Validation
            ↓
Populate Dynamic Templates
            ↓
Generate Excel/File Outputs
            ↓
Log Execution Results
```

---

# Output Formats

QC360 supports:

- Excel (.xlsx)
- CSV (.csv)
- Tab Delimited Files
- Pipe Delimited Files
- Custom Delimited Outputs

Outputs are generated automatically based on configuration settings.

---

# Logging

Execution logs are generated automatically and include:

- Query Execution Details
- Athena Execution IDs
- Query Failures
- Template Processing Details
- Runtime Exceptions
- Process Completion Status

---

# Repository Structure

```text
QC360/
│
├── index.py
├── input_config.py
├── Query.py
├── OVR.py
├── sales_file.py
├── xl_df_mapper.py
├── reporting_config.xlsm
│
├── logs/
├── outputs/
└── README.md
```

---

# Customization

QC360 is designed to be easily forked and customized.

You can:

- Add new QC checks
- Add new Athena queries
- Create custom report templates
- Extend business rules
- Add new output formats
- Modify mappings and configurations
- Build process-specific validation modules

---

# Quick Start

```text
1. Clone Repository
2. Create Virtual Environment
3. Install Dependencies
4. Configure reporting_config.xlsm
5. Execute Macro or Run index.py
6. Review Outputs
```

---

### Tagline

**QC360 — Configure Once. Validate Everywhere. Automate Every QC.**
