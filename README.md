# RVTools Cluster Summary Generator

A tool for Nutanix SEs to quickly extract sizing data from **RVTools** Excel exports.

> 📦 **Looking for the Nutanix Collector version?** See [Collector-Formatter-8000](https://github.com/noko7/Collector-Formatter-8000)

---

## Quick Start

1. **Export from RVTools** - Use "Export All to Excel" in RVTools
2. **Drop the `.xlsx` file** into the `input_files/` folder
3. **Run the script**: `python cluster_summary.py`
4. **Get your results** in the `results/` folder

That's it. Each RVTools export generates its own summary Excel file.

---
<img> https://github.com/Noko7/rvtools-formatter-9000/blob/main/results.png </img>
## What You Get

For each RVTools file, the tool generates:

| Output | Description |
|--------|-------------|
| `cluster_summary_*.xlsx` | Formatted Excel with per-cluster hardware details |
| `validation_*.txt` | Processing log with row counts and data mapping info |

### The Excel Output Contains:

- **Total Summary Table** - All clusters aggregated into one view (top of file)
- **Per-Cluster Sections** - Each cluster gets its own table with:
  - Cluster-level metrics (VMs, vCPU, RAM, storage, ratios)
  - Host hardware details (vendor, model, CPU, RAM, storage per host)

---

## RVTools Compatibility

This tool is specifically designed for **RVTools exports**. It automatically handles:

| RVTools Column | Mapped To |
|----------------|-----------|
| `vHost` sheet | Host data |
| `vInfo` sheet | VM info |
| `Host` (hostname) | VM-to-Host mapping |
| `Cluster` (name) | Cluster grouping |
| `# CPU`, `# Cores`, `# Memory` | Hardware specs |
| `VM UUID` | VM identification |

### Supported RVTools Versions
- RVTools 4.x ✓
- Older versions should work but may need column mapping updates

---

## When to Use This

✅ **Pre-sales sizing** - Quick hardware inventory before creating a Sizer scenario  
✅ **Multi-cluster environments** - Get a consolidated view across clusters  
✅ **Hardware refresh projects** - Document existing host specs for replacement planning  
✅ **Proposal prep** - Extract key metrics for customer presentations  
✅ **Customer doesn't have Collector access** - RVTools is often easier to get

---

## When NOT to Use This

❌ **Performance sizing** - This is capacity data only, not performance/utilization metrics  
❌ **Storage deep-dive** - Use the raw RVTools vDatastore tab  
❌ **Single-VM analysis** - This aggregates at cluster/host level, not per-VM  

---

## Important Notes on the Data

### What's Included
- Host hardware specs (vendor, model, CPU model, cores, RAM, NICs)
- VM counts and vCPU totals per cluster
- VM memory allocation (not utilization)
- vDisk provisioned and consumed storage
- vCPU to pCore ratios

### What's NOT Included
- CPU/Memory utilization percentages  
- IOPS or throughput metrics
- Network bandwidth data
- Historical trends

### Storage Numbers
- **Provisioned Storage** = Total vDisk size allocated to VMs
- **Consumed Storage** = Actual space used on datastore  
- **Capacity vDisk** = vDisk capacity (may differ from provisioned for thin disks)
- All values in TiB (binary terabytes)

### RAM Numbers
- Host RAM shown in GiB (converted from MiB as reported by RVTools)
- VM RAM shown in GiB (converted from MiB)

---

## Folder Structure

```
rvtools-formatter/
├── cluster_summary.py      # The main script
├── input_files/            # DROP RVTOOLS EXPORTS HERE
│   └── (your RVTools .xlsx files)
└── results/                # OUTPUT GOES HERE
    └── cluster_summary_*.xlsx
```

---

## Requirements

- Python 3.8+
- pandas (`pip install pandas`)
- openpyxl (`pip install openpyxl`)

First-time setup:
```bash
pip install pandas openpyxl
```

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No input files found" | Make sure Excel files are in `input_files/` folder |
| Missing required sheets | Collector must have vCluster, vHosts, vInfo, vCPU, vMemory, vDisk, vPartition sheets |
| Permission denied | Close any open Excel files in the results folder |
| openpyxl warning | Run `pip install openpyxl` |

---

## Legacy Mode

If you have individual CSV files (vCluster.csv, vHosts.csv, etc.) instead of a collector Excel file, just place them in the same folder as the script. The tool will process them directly.

---

## Questions?

This tool uses MOID-based joins (not cluster/VM names) so it handles obfuscated collector data correctly. The output is meant as a starting point for Sizer - always validate key numbers with the customer.
