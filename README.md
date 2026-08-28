# PV Automation Hub

This package displays both existing Streamlit applications inside one parent application.
The two application codes remain separate:

- `pages/triage_app.py`
- `pages/quality_reviewer.py`

## Common files

Place these shared masters in the single `data/` folder:

- `data/MedDRA.xlsx` or `data/MedDRA.csv`
- `data/Listedness_CX.xlsx`

Both pages resolve their `data` folder from the project root when deployed with this layout.
If a page currently builds its data path from its own file location, change its BASE_DIR line to import
`BASE_DIR` and `DATA_DIR` from `shared_context`.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Linking the two applications

Both pages share the same Streamlit session. Use:

```python
from shared_context import set_shared_value, get_shared_value

set_shared_value("case_id", sender_id)
case_id = get_shared_value("case_id", "")
```

Recommended linked workflow:

1. Parse multiple XMLs in Triage.
2. Store the chosen case ID and XML bytes in shared session state.
3. Open Quality Reviewer from the sidebar.
4. Pre-populate its upload/selection controls from shared state.
5. Store its mismatch summary back in shared state for export from Triage.

## Important

The original business logic has not been merged. Only the page configuration call was removed from each
page because Streamlit requires it in the common entrypoint. Test all workflows before production use.
