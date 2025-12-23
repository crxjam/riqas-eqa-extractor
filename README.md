cat > README.md <<'EOF'
# RIQAS EQA PDF → Excel Extractor

Parses RIQAS PDF reports, extracts analyte metrics, applies risk logic (bias/trend/history escalation),
and writes into an Excel template with:
- Header Information
- Result Summary
- Cycle_History

## Install
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
